/**
 * Replays the google.script.run calls Index.html actually makes.
 *
 * test/ui_calls.json is recorded by driving Index.html in Chromium against this
 * same server code (see record_ui_calls.js). This suite re-runs every recorded
 * call against Code.js — no browser needed — and checks that each flow still gets
 * what the page does with the answer, so a tightened endpoint that breaks the UI
 * fails here rather than in production.
 */

const fs = require('fs');
const path = require('path');
const { createEnv } = require('./apps_script_env');
const { USERS, sheets } = require('./fixtures');
const recording = require('./ui_calls.json');

exports.name = 'Index.html flows (recorded google.script.run calls)';

exports.run = function ({ test, assert }) {
  const flows = {};
  recording.flows.forEach(f => { flows[f.name] = f; });

  /** Replays one recorded flow and returns the responses, keyed by call. */
  function replay(flowName) {
    const flow = flows[flowName];
    if (!flow) throw new Error('No recorded flow named ' + flowName);
    const env = createEnv({ activeUser: flow.signedInAs, sheets: sheets() });

    return flow.calls.map(call => {
      const result = env.attempt(call.fn, ...call.args);
      if (result.ok !== call.ok) {
        throw new Error(`${call.fn}(${JSON.stringify(call.args)}) was recorded as ` +
          `${call.ok ? 'allowed' : 'refused'} but is now ${result.ok ? 'allowed' : 'refused'}` +
          (result.ok ? '' : ': ' + result.error.message));
      }
      return { fn: call.fn, args: call.args, value: result.value };
    });
  }

  const responseTo = (responses, fn) => (responses.find(r => r.fn === fn) || {}).value;

  test('the browser run made no failed calls and threw no page errors', () => {
    recording.flows.forEach(f => {
      assert.deepEqual(f.handledErrors, [], `${f.name} hit a failure handler`);
      assert.deepEqual(f.pageErrors, [], `${f.name} threw in the page`);
    });
  });

  // ---- Teacher -------------------------------------------------------------

  test('teacher flow: every recorded call still succeeds', () => {
    const responses = replay('teacher');
    assert.ok(responses.length >= 3, 'the teacher flow should exercise several endpoints');
  });

  test("teacher flow: the Submit dropdown still gets colleagues' names", () => {
    const initial = responseTo(replay('teacher'), 'getInitialData');
    const names = initial.staffData.map(s => s.name).filter(Boolean);

    assert.ok(names.length > 1, 'renderTeacherSubmit maps staffData to <option> labels');
    assert.deepEqual(flows.teacher.observations.subbedForOptions.slice().sort(),
      names.slice().sort(), 'what the browser rendered matches what the server returns');
  });

  test('teacher flow: My Report still has the numbers the KPI cards show', () => {
    const initial = responseTo(replay('teacher'), 'getInitialData');
    const me = initial.staffData.find(s => s.email.toLowerCase() === USERS.omsTeacher);

    // openStaffDetail renders Carry Over / Earned / Used / Paid Out / Balance.
    ['carryOver', 'earned', 'used', 'paidOut', 'total'].forEach(field => {
      assert.equal(typeof me[field], 'number', `${field} must be a number, not "${me[field]}"`);
    });
    assert.deepEqual(flows.teacher.observations.kpis,
      ['My TST Summary', '3.00', '2.50', '1.00', '1.00', '3.50'],
      'the rendered KPI cards');
  });

  test('teacher flow: the history call still returns the teacher\'s own rows', () => {
    const history = responseTo(replay('teacher'), 'getTeacherHistory');
    assert.ok(history.length > 0);
    assert.ok(history.some(h => h.type === 'Earned'));
  });

  test('teacher flow: the Schedule tab still renders the saved checkboxes', () => {
    const schedule = responseTo(replay('teacher'), 'getScheduleData');
    const mine = (schedule['September'] || []).filter(r => r.email.toLowerCase() === USERS.omsTeacher);

    assert.equal(mine.length, 1);
    assert.deepEqual(flows.teacher.observations.checkedBoxes,
      ['Period 1 - 8:10 - 8:57_Mon', 'Period 1 - 8:10 - 8:57_Tue'],
      'the availability grid the teacher sees');
  });

  test('teacher flow: nothing in any response carries another staff member\'s balance', () => {
    replay('teacher').forEach(({ fn, value }) => {
      const rows = fn === 'getInitialData' ? value.staffData : (Array.isArray(value) ? value : []);
      rows.forEach(row => {
        if (!row || !row.email) return;
        if (row.email.toString().toLowerCase() === USERS.omsTeacher) return;
        assert.equal(row.carryOver, undefined, `${fn} leaked Carry Over for ${row.email}`);
        assert.equal(row.paidOut, undefined, `${fn} leaked Paid Out for ${row.email}`);
      });
    });
  });

  // ---- Admin ---------------------------------------------------------------

  test('admin flow: every recorded call still succeeds', () => {
    const responses = replay('admin');
    assert.ok(responses.some(r => r.fn === 'getPendingEarned'));
    assert.ok(responses.some(r => r.fn === 'getStaffDirectoryData'));
  });

  test('admin flow: the Directory still shows balances', () => {
    const dir = responseTo(replay('admin'), 'getStaffDirectoryData');
    const teacher = dir.find(s => s.email.toLowerCase() === USERS.omsTeacher);

    assert.equal(teacher.carryOver, 3);
    assert.equal(teacher.earned, 2.5);
    assert.ok(flows.admin.observations.directoryHasTeacher, 'and the table rendered them');
  });

  test('admin flow: the badges match the queues', () => {
    const responses = replay('admin');
    const counts = responseTo(responses, 'getDashboardCounts');
    const earned = responseTo(responses, 'getPendingEarned');
    const used = responseTo(responses, 'getPendingUsed');

    assert.equal(counts.earned, earned.length);
    assert.equal(counts.used, used.length);
    assert.equal(flows.admin.observations.badges.earned, String(earned.length));
    assert.equal(flows.admin.observations.badges.used, String(used.length));
    assert.equal(flows.admin.observations.earnedRows, earned.length, 'rows rendered in the queue');
  });

  test('admin flow: the master schedule still lists the whole building', () => {
    const schedule = responseTo(replay('admin'), 'getScheduleData');
    const emails = (schedule['September'] || []).map(r => r.email.toLowerCase());

    assert.ok(emails.includes(USERS.omsTeacher));
    assert.ok(emails.includes(USERS.omsTeacher2), 'an admin needs everyone, not just themselves');
  });

  // ---- View As -------------------------------------------------------------

  test('View As flow: every recorded call still succeeds', () => {
    const viewAs = responseTo(replay('admin-view-as-teacher'), 'getViewAsData');
    assert.equal(viewAs.email, USERS.omsTeacher);
    assert.equal(viewAs.role, 'Teacher');
  });

  test('View As flow: the admin still gets a usable teacher payload', () => {
    const viewAs = responseTo(replay('admin-view-as-teacher'), 'getViewAsData');
    const me = viewAs.staffData.find(s => s.email.toLowerCase() === USERS.omsTeacher);

    assert.equal(me.total, 3.5, 'the usage form caps on this');
    assert.ok(viewAs.staffData.length > 1, 'the Submit dropdown needs colleagues');
    assert.deepEqual(flows['admin-view-as-teacher'].observations.kpis,
      ['My TST Summary', '3.00', '2.50', '1.00', '1.00', '3.50']);
    assert.ok(/Viewing as Tina Teacher/.test(flows['admin-view-as-teacher'].observations.banner));
  });

  // ---- Calls the recording does not cover ------------------------------------
  //
  // ui_calls.json only holds the flows the browser script walked through, so an
  // endpoint used on a screen nobody recorded is unguarded until production. The
  // obvious fix — parse every google.script.run chain out of Index.html — needs a
  // real JavaScript parser: the page nests template literals and contains regex
  // literals with quotes in them, and anything less drifts and reports phantom
  // failures. These three checks stay within what can be asserted soundly.

  test('Index.html never calls a private server function', () => {
    const html = fs.readFileSync(path.join(__dirname, '..', 'Index.html'), 'utf8');
    const env = createEnv({ activeUser: USERS.omsAdmin, sheets: sheets() });

    const underscored = new Set(
      [...html.matchAll(/\.([A-Za-z_$][\w$]*_)\s*\(/g)].map(m => m[1])
    );
    underscored.forEach(name => {
      assert.ok(typeof env.context[name] !== 'function',
        `Index.html calls ${name}(), which Apps Script never exposes to a client — ` +
        'it would fail silently in the browser');
    });
  });

  test('the endpoints the page dispatches by computed name all exist', () => {
    const env = createEnv({ activeUser: USERS.omsAdmin, sheets: sheets() });
    // batch${Action}${Type} in the multiselect toolbar, and the revert pair in the
    // staff detail modal. A rename on the server would be silent otherwise.
    ['batchApproveEarned', 'batchDenyEarned', 'batchApproveUsed', 'batchDeleteUsed',
     'revertEarnedToPending', 'revertUsedToPending'].forEach(fn => {
      assert.equal(typeof env.context[fn], 'function', `Code.js no longer defines ${fn}()`);
    });
  });

  test('the assignment endpoints the page needs are among them', () => {
    const html = fs.readFileSync(path.join(__dirname, '..', 'Index.html'), 'utf8');
    ['assignCoverage', 'getAssignments', 'getMyAssignments', 'recordAssignment',
     'cancelAssignment', 'remindAssignment',
     'sendTestCalendarEvent', 'getCalendarTestResult',
     'getEmailServiceStatus', 'setAuthorizeUrl'].forEach(fn => {
      assert.ok(html.includes('.' + fn + '('), `Index.html should call ${fn}()`);
    });
  });
};
