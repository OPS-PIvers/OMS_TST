/**
 * Server-side authorization on the read endpoints.
 *
 * Every public (non-underscore) function is callable from the browser console by
 * any signed-in domain user, so these tests ask the same question of each one:
 * what does a Teacher get, what does an admin from another building get, and does
 * the legitimate caller still get everything they need?
 */

const { createEnv } = require('./apps_script_env');
const { USERS, sheets, STAFF_HEADER, STAFF_ROWS, APPROVALS_HEADER } = require('./fixtures');

const envFor = (email, overrides) => createEnv({ activeUser: email, sheets: sheets(overrides) });

exports.name = 'Read endpoint authorization';

exports.run = function ({ test, assert }) {

  // ---- getStaffDirectoryData -------------------------------------------------

  test('teacher gets colleagues by name/email only — no balances, carry over or paid out', () => {
    const env = envFor(USERS.omsTeacher);
    const dir = env.run('getStaffDirectoryData', 'OMS');

    const others = dir.filter(s => s.email.toLowerCase() !== USERS.omsTeacher);
    assert.ok(others.length > 0, 'the roster should still list colleagues');
    others.forEach(s => {
      assert.ok(s.name, `${s.email} should keep a display name`);
      assert.equal(s.carryOver, undefined, `${s.email} leaked Carry Over`);
      assert.equal(s.paidOut, undefined, `${s.email} leaked Paid Out`);
      assert.equal(s.earned, undefined, `${s.email} leaked Earned`);
      assert.equal(s.used, undefined, `${s.email} leaked Used`);
      assert.equal(s.total, undefined, `${s.email} leaked a running total`);
      assert.equal(s.rowIndex, undefined, `${s.email} leaked its directory row`);
    });
  });

  test('teacher still gets their own full record (the usage form is capped by it)', () => {
    const env = envFor(USERS.omsTeacher);
    const me = env.run('getStaffDirectoryData', 'OMS')
      .find(s => s.email.toLowerCase() === USERS.omsTeacher);

    assert.ok(me, 'the teacher must appear in their own directory payload');
    assert.equal(me.carryOver, 3);
    assert.equal(me.paidOut, 1);
    assert.equal(me.earned, 2.5, 'approved earned hours');
    assert.equal(me.used, 1, 'approved used hours');
    assert.equal(me.total, 3 + 2.5 - 1 - 1);
  });

  test("a teacher cannot read another building's directory", () => {
    const env = envFor(USERS.omsTeacher);
    const dir = env.run('getStaffDirectoryData', 'OHS');
    const emails = dir.map(s => s.email.toLowerCase());

    assert.ok(!emails.includes(USERS.ohsTeacher), 'OHS-only staff must not appear');
    assert.ok(emails.includes(USERS.omsTeacher2), 'the call falls back to the caller\'s own building');
  });

  test('a null building no longer returns the whole district', () => {
    const env = envFor(USERS.omsAdmin);
    const emails = env.run('getStaffDirectoryData', null).map(s => s.email.toLowerCase());

    assert.ok(emails.includes(USERS.omsTeacher), 'OMS staff are in scope');
    assert.ok(!emails.includes(USERS.ohsTeacher), 'OHS staff are not');
  });

  test("an OMS admin asking for OHS gets OMS, not OHS", () => {
    const env = envFor(USERS.omsAdmin);
    const emails = env.run('getStaffDirectoryData', 'OHS').map(s => s.email.toLowerCase());

    assert.ok(!emails.includes(USERS.ohsTeacher), 'OHS-only staff must not appear');
    assert.ok(emails.includes(USERS.omsTeacher), 'fell back to the admin\'s own building');
  });

  test('a multi-building admin can read either of their buildings', () => {
    const env = envFor(USERS.dualAdmin);
    const ohs = env.run('getStaffDirectoryData', 'OHS').map(s => s.email.toLowerCase());
    assert.ok(ohs.includes(USERS.ohsTeacher), 'OHS is one of Dana\'s buildings');

    const oms = env.run('getStaffDirectoryData', 'OMS').map(s => s.email.toLowerCase());
    assert.ok(oms.includes(USERS.omsTeacher));
  });

  test('an admin still gets balances, and Super Admins can pick any building', () => {
    const admin = envFor(USERS.omsAdmin).run('getStaffDirectoryData', 'OMS')
      .find(s => s.email.toLowerCase() === USERS.omsTeacher);
    assert.equal(admin.carryOver, 3);
    assert.equal(admin.earned, 2.5);

    const superAdmin = envFor(USERS.superAdmin).run('getStaffDirectoryData', 'OHS')
      .map(s => s.email.toLowerCase());
    assert.ok(superAdmin.includes(USERS.ohsTeacher));
  });

  test('a multi-building teacher keeps combined balances and can switch buildings', () => {
    const env = envFor(USERS.multiTeacher);

    const atOms = env.run('getStaffDirectoryData', 'OMS');
    const meAtOms = atOms.find(s => s.email.toLowerCase() === USERS.multiTeacher);
    assert.equal(meAtOms.earned, 2, 'earned at OMS + OHS are combined');
    assert.ok(atOms.some(s => s.email.toLowerCase() === USERS.omsTeacher), 'sees OMS colleagues');

    const atOhs = env.run('getStaffDirectoryData', 'OHS');
    assert.ok(atOhs.some(s => s.email.toLowerCase() === USERS.ohsTeacher), 'sees OHS colleagues');
    assert.ok(!atOhs.some(s => s.email.toLowerCase() === USERS.omsTeacher2),
      'OMS-only colleagues stay out of the OHS roster');
    const meAtOhs = atOhs.find(s => s.email.toLowerCase() === USERS.multiTeacher);
    assert.equal(meAtOhs.carryOver, 4, 'own record is complete in either building');
  });

  test('someone outside the directory gets no one else\'s data', () => {
    const env = envFor(USERS.stranger);
    const dir = env.run('getStaffDirectoryData', 'OMS');
    dir.forEach(s => assert.equal(s.carryOver, undefined, `${s.email} leaked Carry Over to a stranger`));
  });

  test('archived staff stay hidden from the teacher roster', () => {
    const env = envFor(USERS.omsTeacher);
    const emails = env.run('getStaffDirectoryData', 'OMS', null, true).map(s => s.email.toLowerCase());
    assert.ok(!emails.includes(USERS.archivedTeacher),
      'includeArchived must not let a teacher pull archived rows');
  });

  // ---- getInitialData (the same directory rules, on the payload the app boots with)

  test('getInitialData gives a teacher the same redacted roster', () => {
    const data = envFor(USERS.omsTeacher).run('getInitialData');
    const others = data.staffData.filter(s => s.email.toLowerCase() !== USERS.omsTeacher);

    assert.ok(others.length > 0);
    others.forEach(s => {
      assert.ok(s.name, 'names are what the Submit dropdown needs');
      assert.equal(s.carryOver, undefined, `${s.email} leaked Carry Over on boot`);
      assert.equal(s.paidOut, undefined, `${s.email} leaked Paid Out on boot`);
    });

    const me = data.staffData.find(s => s.email.toLowerCase() === USERS.omsTeacher);
    assert.equal(me.total, 3.5, 'the teacher still boots with their own balance');
  });

  test('getInitialData gives an admin the full building', () => {
    const data = envFor(USERS.omsAdmin).run('getInitialData');
    const teacher = data.staffData.find(s => s.email.toLowerCase() === USERS.omsTeacher);
    assert.equal(teacher.carryOver, 3);
    assert.equal(teacher.earned, 2.5);
  });

  test('someone outside the directory boots with an empty roster', () => {
    const data = envFor(USERS.stranger).run('getInitialData');
    assert.deepEqual(data.staffData, []);
    assert.equal(data.role, 'Guest');
  });

  // ---- View As ---------------------------------------------------------------

  test('View As still hands the admin a working teacher payload', () => {
    const data = envFor(USERS.omsAdmin).run('getViewAsData', USERS.omsTeacher, 'OMS');

    assert.equal(data.email, USERS.omsTeacher);
    assert.equal(data.role, 'Teacher');
    const me = data.staffData.find(s => s.email.toLowerCase() === USERS.omsTeacher);
    assert.equal(me.total, 3.5, 'the viewed teacher\'s balance drives the usage form');
    assert.ok(data.staffData.some(s => s.email.toLowerCase() === USERS.omsTeacher2),
      'colleagues are still there for the Submit dropdown');
  });

  test('View As is still refused across buildings', () => {
    assert.rejected(envFor(USERS.ohsAdmin).attempt('getViewAsData', USERS.omsTeacher, 'OMS'),
      /own building/i);
  });

  test('a teacher cannot View As anyone', () => {
    assert.rejected(envFor(USERS.omsTeacher).attempt('getViewAsData', USERS.omsTeacher2, 'OMS'),
      /admin access required/i);
  });

  // ---- Admin queues ----------------------------------------------------------

  ['getPendingEarned', 'getPendingUsed', 'getDashboardCounts'].forEach(fn => {
    test(`${fn} refuses a teacher`, () => {
      assert.rejected(envFor(USERS.omsTeacher).attempt(fn, 'OMS'), /admin access required/i);
    });

    test(`${fn} refuses someone outside the directory`, () => {
      assert.rejected(envFor(USERS.stranger).attempt(fn, 'OMS'), /admin access required/i);
    });
  });

  test('an OMS admin cannot read the OHS queues', () => {
    const env = envFor(USERS.omsAdmin);

    const earned = env.run('getPendingEarned', 'OHS');
    assert.ok(earned.every(r => r.building === 'OMS'), 'only OMS rows come back');
    assert.ok(earned.some(r => r.email.toLowerCase() === USERS.omsTeacher2), 'the OMS queue is intact');

    const used = env.run('getPendingUsed', 'OHS');
    assert.ok(used.every(r => r.building === 'OMS'));
  });

  test('pending earned includes rows with a blank building (they default to OMS)', () => {
    const rows = envFor(USERS.omsAdmin).run('getPendingEarned', 'OMS');
    assert.equal(rows.length, 2, 'the blank-building row counts as OMS');
  });

  test('an OHS admin sees only the OHS queues', () => {
    const env = envFor(USERS.ohsAdmin);
    const earned = env.run('getPendingEarned');
    assert.equal(earned.length, 1);
    assert.equal(earned[0].email.toLowerCase(), USERS.ohsTeacher);
  });

  test('a multi-building admin can read both queues by switching building', () => {
    const env = envFor(USERS.dualAdmin);
    assert.equal(env.run('getPendingEarned', 'OHS').length, 1, 'OHS queue');
    assert.equal(env.run('getPendingEarned', 'OMS').length, 2, 'OMS queue');
  });

  test('a building added through Settings (not config.js) is still readable by its admin', () => {
    // saveBuildingConfig can append a building code that BUILDING_CONFIG never had;
    // the directory assignment is what authorizes the read.
    const staff = STAFF_ROWS.map(r => r.slice());
    staff.find(r => r[1] === USERS.dualAdmin)[8] = 'OMS, OPS';
    staff.push(['Pat Prairie', 'pat.prairie@orono.k12.mn.us', 'Teacher', '', '', 0, 0, '', 'OPS', '', '', '']);

    const env = createEnv({
      activeUser: USERS.dualAdmin,
      sheets: sheets({
        'Staff Directory': [STAFF_HEADER, ...staff],
        'TST Approvals (New)': [APPROVALS_HEADER,
          ['pat.prairie@orono.k12.mn.us', 'Pat Prairie', 'Someone', '', '2025-09-20', 'Period 1', 'Full Period', 1, false, '', false, '', '', 'OPS']]
      })
    });

    const rows = env.run('getPendingEarned', 'OPS');
    assert.equal(rows.length, 1, 'the OPS admin can see the OPS queue');
    assert.ok(env.run('getStaffDirectoryData', 'OPS')
      .some(s => s.email === 'pat.prairie@orono.k12.mn.us'), 'and the OPS directory');
  });

  test('dashboard counts match the queues they badge', () => {
    const env = envFor(USERS.dualAdmin);
    ['OMS', 'OHS'].forEach(b => {
      const counts = env.run('getDashboardCounts', b);
      assert.equal(counts.earned, env.run('getPendingEarned', b).length, `${b} earned badge`);
      assert.equal(counts.used, env.run('getPendingUsed', b).length, `${b} used badge`);
    });
  });

  test('an OMS admin cannot get OHS counts by passing a building', () => {
    const env = envFor(USERS.omsAdmin);
    assert.deepEqual(env.run('getDashboardCounts', 'OHS'), env.run('getDashboardCounts', 'OMS'));
  });

  // ---- getScheduleData -------------------------------------------------------

  test('a teacher gets only their own availability rows', () => {
    const env = envFor(USERS.omsTeacher);
    const schedule = env.run('getScheduleData', 'OMS');

    const all = Object.keys(schedule).reduce((acc, m) => acc.concat(schedule[m]), []);
    assert.ok(all.length > 0, 'the Schedule tab still needs the teacher\'s own rows');
    all.forEach(r => assert.equal(r.email.toLowerCase(), USERS.omsTeacher,
      `leaked ${r.email}'s availability`));
  });

  test("the teacher Schedule tab renders the same grid it did before", () => {
    // renderTeacherSchedule reads scheduleData[month] and filters to its own email.
    const schedule = envFor(USERS.omsTeacher).run('getScheduleData', 'OMS');
    const september = schedule['September'] || [];
    const mine = september.filter(d => d.email.toLowerCase() === USERS.omsTeacher);

    assert.equal(mine.length, 1);
    assert.equal(mine[0].period, 'Period 1 - 8:10 - 8:57');
    assert.equal(mine[0].days, 'Mon,Tue');
  });

  test('an admin still gets the whole building schedule', () => {
    const schedule = envFor(USERS.omsAdmin).run('getScheduleData', 'OMS');
    const emails = (schedule['September'] || []).map(r => r.email.toLowerCase());

    assert.ok(emails.includes(USERS.omsTeacher));
    assert.ok(emails.includes(USERS.omsTeacher2));
    assert.ok(!emails.includes(USERS.ohsTeacher), 'other buildings stay out');
  });

  test('an OMS admin asking for the OHS schedule gets OMS', () => {
    const schedule = envFor(USERS.omsAdmin).run('getScheduleData', 'OHS');
    const emails = (schedule['September'] || []).map(r => r.email.toLowerCase());
    assert.ok(!emails.includes(USERS.ohsTeacher));
  });

  test("a teacher's schedule rows carry no one else's pending requests", () => {
    const schedule = envFor(USERS.omsTeacher).run('getScheduleData', 'OMS');
    const all = Object.keys(schedule).reduce((acc, m) => acc.concat(schedule[m]), []);
    all.forEach(r => {
      (r.pendingRequests || []).forEach(p => {
        assert.ok(r.email.toLowerCase() === USERS.omsTeacher,
          'pending requests only ride along on the caller\'s own row');
      });
    });
  });

  // ---- Write paths that read the directory internally ------------------------
  // (they call the private staffDirectoryData_, which must stay unrestricted)

  test('a teacher can still submit earned time, tagged to their building', () => {
    const env = envFor(USERS.omsTeacher, {
      'Form Responses 1': [['Timestamp', 'Email', 'Subbed For', 'Other', 'Date', 'Period', 'Type', 'Decimal']]
    });

    assert.allowed(env.attempt('submitEarned', {
      email: USERS.omsTeacher, subbedForName: 'Ted Teacher', subbedForType: 'Staff',
      date: '2025-11-03', period: 'Period 3 - 9:52 - 10:39', amountType: 'Full Period', amountDecimal: 1
    }));

    const rows = env.sheet('TST Approvals (New)').values;
    const added = rows[rows.length - 1];
    assert.equal(added[0], USERS.omsTeacher);
    assert.equal(added[1], 'Tina Teacher', 'the name came from the directory lookup');
    assert.equal(added[13], 'OMS', 'tagged with the submitter\'s primary building');
  });

  test('a teacher can still submit usage', () => {
    const env = envFor(USERS.omsTeacher);
    assert.allowed(env.attempt('submitUsage', {
      email: USERS.omsTeacher, name: 'Tina Teacher', date: '2025-11-04', amount: 1
    }));

    const rows = env.sheet('TST Usage (New)').values;
    assert.equal(rows[rows.length - 1][0], USERS.omsTeacher);
    assert.equal(rows[rows.length - 1][7], 'OMS');
  });

  test('a teacher still cannot submit for someone else', () => {
    const env = envFor(USERS.omsTeacher);
    assert.rejected(env.attempt('submitUsage', {
      email: USERS.omsTeacher2, name: 'Ted Teacher', date: '2025-11-04', amount: 1
    }), /admin access required|own building/i);
  });

  // ---- Maintenance endpoints -------------------------------------------------

  test('syncMissingSubmissions refuses a teacher', () => {
    const env = envFor(USERS.omsTeacher, {
      'Form Responses 1': [['Timestamp', 'Email', 'Subbed For', 'Other', 'Date', 'Period', 'Type', 'Decimal']]
    });
    assert.rejected(env.attempt('syncMissingSubmissions'), /admin access required/i);
  });

  test('syncMissingSubmissions still runs for an admin', () => {
    const env = envFor(USERS.omsAdmin, {
      'Form Responses 1': [['Timestamp', 'Email', 'Subbed For', 'Other', 'Date', 'Period', 'Type', 'Decimal']]
    });
    assert.equal(assert.allowed(env.attempt('syncMissingSubmissions')), 'Synced 0 missing submissions.');
  });

  test('processEmailQueue refuses a teacher calling it from the client', () => {
    const env = envFor(USERS.omsTeacher, { 'Email Queue': emailQueue() });
    assert.rejected(env.attempt('processEmailQueue'), /admin access required/i);
    assert.equal(env.sentEmails.length, 0, 'nothing may be sent');
  });

  test('a forged trigger event does not get a teacher past the check', () => {
    const env = envFor(USERS.omsTeacher, { 'Email Queue': emailQueue() });
    // google.script.run can only send JSON, so this is the best a client can do.
    assert.rejected(env.attempt('processEmailQueue', { authMode: 'FULL', triggerUid: '123' }),
      /admin access required/i);
    assert.equal(env.sentEmails.length, 0);
  });

  test('the installed trigger still drains the queue', () => {
    const env = envFor(USERS.omsAdmin, { 'Email Queue': emailQueue() });
    assert.allowed(env.attempt('processEmailQueue', { authMode: env.AuthMode.FULL, triggerUid: '123' }));
    assert.equal(env.sentEmails.length, 1, 'the queued OMS email was sent');
    assert.equal(env.sentEmails[0].to, USERS.omsTeacher);
  });

  test('an admin can flush the queue by hand', () => {
    const env = envFor(USERS.omsAdmin, { 'Email Queue': emailQueue() });
    assert.allowed(env.attempt('processEmailQueue'));
    assert.equal(env.sentEmails.length, 1);
  });

  test("the queue still only sends the caller's own buildings", () => {
    const env = envFor(USERS.ohsAdmin, { 'Email Queue': emailQueue() });
    assert.allowed(env.attempt('processEmailQueue'));
    assert.equal(env.sentEmails.length, 0, 'the OMS row belongs to the OMS admin');
  });

  // ---- Private helpers are off the client surface ----------------------------

  ['staffDirectoryData_', 'calculateDynamicBalances_',
   'getPendingEarnedMap_', 'pendingEarnedFor_', 'pendingUsedFor_', 'scheduleData_',
   'processEmailQueue_'].forEach(fn => {
    test(`${fn} is not reachable from google.script.run`, () => {
      const env = envFor(USERS.omsTeacher);
      assert.rejected(env.attempt(fn), /private/i);
      assert.equal(typeof env.context[fn], 'function', 'but server code can still call it');
    });
  });
};

function emailQueue() {
  return [
    ['Timestamp', 'Recipient', 'Subject', 'Body', 'Building', 'Status', 'LastUpdated', 'Options'],
    [new Date(), USERS.omsTeacher, 'TST Report', '<p>hi</p>', 'OMS', 'Pending', '', '{}']
  ];
}
