/**
 * TST Coverage Assignments.
 *
 * Coverage is assigned rather than requested, and the assignment row — not an email
 * in someone's inbox — is the record. These tests ask the same three questions of
 * every new endpoint (what does a Teacher get, what does an admin from another
 * building get, what does a Super Admin get) and then check the behaviour the
 * design turns on: the duplicate guard, the claim-before-submit recording path, the
 * 14-day link expiry, cancellation against approved hours, and the one nudge.
 */

const { createEnv } = require('./apps_script_env');
const { USERS, sheets } = require('./fixtures');

const envFor = (email, overrides) => createEnv({ activeUser: email, sheets: sheets(overrides) });

/** A yyyy-MM-dd key `offset` days from today — negative is in the past. */
function ymd(offset) {
  const d = new Date();
  d.setDate(d.getDate() + (offset || 0));
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

const OMS_PERIOD = 'Period 1 - 8:10 - 8:57';
const OMS_PERIOD_2 = 'Period 2 - 9:01 - 9:48';

/** The payload the Assign Coverage modal sends. */
function payload(extra) {
  return Object.assign({
    teacherEmail: USERS.omsTeacher,
    teacherName: 'Tina Teacher',
    subbedFor: 'Ted Teacher',
    coveredForEmail: USERS.omsTeacher2,
    date: ymd(3),
    period: OMS_PERIOD,
    amount: 1,
    amountType: 'Full Period',
    building: 'OMS'
  }, extra || {});
}

/** Rows currently sitting in the Email Queue sheet, as objects. */
function queued(env) {
  const sheet = env.sheet('Email Queue');
  if (!sheet) return [];
  return sheet.values.slice(1).map(r => ({
    recipient: r[1], subject: r[2], body: r[3], building: r[4], status: r[5]
  }));
}

function assignmentSheetRows(env) {
  const sheet = env.sheet('TST Assignments');
  return sheet ? sheet.values.slice(1) : [];
}

/** Creates one assignment as the OMS admin and returns its id. */
function assign(env, extra) {
  const res = env.run('assignCoverage', payload(extra));
  if (!res || !res.id) throw new Error('assignCoverage did not create a row: ' + JSON.stringify(res));
  return res.id;
}

exports.name = 'Coverage assignments';

exports.run = function ({ test, assert }) {

  // ---- assignCoverage: who may assign ---------------------------------------

  test('a teacher cannot assign coverage', () => {
    const env = envFor(USERS.omsTeacher);
    assert.rejected(env.attempt('assignCoverage', payload()), /admin access required/i);
    assert.equal(assignmentSheetRows(env).length, 0, 'nothing may be written');
    assert.equal(queued(env).length, 0, 'and nothing may be emailed');
  });

  test('an admin from another building cannot assign OMS staff', () => {
    const env = envFor(USERS.ohsAdmin);
    assert.rejected(env.attempt('assignCoverage', payload()), /your own building/i);
    assert.equal(assignmentSheetRows(env).length, 0);
  });

  test('the building admin can assign their own staff', () => {
    const env = envFor(USERS.omsAdmin);
    const res = assert.allowed(env.attempt('assignCoverage', payload()));
    assert.ok(res.assigned);
    assert.equal(assignmentSheetRows(env).length, 1);
  });

  test('a Super Admin can assign in a building they are not assigned to', () => {
    const env = envFor(USERS.superAdmin);
    assert.allowed(env.attempt('assignCoverage', payload({
      teacherEmail: USERS.ohsTeacher,
      teacherName: 'Hank Teacher',
      subbedFor: 'Otto Admin',
      coveredForEmail: USERS.ohsAdmin,
      period: 'Period 1',
      building: 'OHS'
    })));
    assert.equal(assignmentSheetRows(env).length, 1);
  });

  test('an admin cannot assign someone to cover for themselves', () => {
    const env = envFor(USERS.omsAdmin);
    assert.rejected(
      env.attempt('assignCoverage', payload({ coveredForEmail: USERS.omsTeacher, subbedFor: 'Tina Teacher' })),
      /cover for themselves/i);
  });

  test('a zero or missing duration is refused', () => {
    const env = envFor(USERS.omsAdmin);
    assert.rejected(env.attempt('assignCoverage', payload({ amount: 0 })), /greater than zero/i);
    assert.rejected(env.attempt('assignCoverage', payload({ amount: null })), /greater than zero/i);
  });

  // ---- The duplicate guard ---------------------------------------------------

  test('the same person, date and period twice is refused outright', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env);
    assert.rejected(env.attempt('assignCoverage', payload()), /already assigned/i);
    assert.equal(assignmentSheetRows(env).length, 1, 'the second attempt writes nothing');
  });

  test('a different period on the same day comes back as a conflict to confirm', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env);

    const res = assert.allowed(env.attempt('assignCoverage', payload({ period: OMS_PERIOD_2 })));
    assert.ok(res.conflict, 'covering twice in a day is legitimate, so it asks rather than refuses');
    assert.equal(res.existing.length, 1);
    assert.equal(res.existing[0].period, OMS_PERIOD);
    assert.equal(assignmentSheetRows(env).length, 1, 'a conflict must not write a row');
    assert.equal(queued(env).length, 3, 'nor send a second set of emails');
  });

  test('force: true goes through the same-day conflict', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env);
    assert.allowed(env.attempt('assignCoverage', payload({ period: OMS_PERIOD_2, force: true })));
    assert.equal(assignmentSheetRows(env).length, 2);
  });

  test('a cancelled assignment no longer blocks re-assigning the same slot', () => {
    const env = envFor(USERS.omsAdmin);
    const id = assign(env);
    env.run('cancelAssignment', id);
    assert.allowed(env.attempt('assignCoverage', payload()));
  });

  // ---- The emails ------------------------------------------------------------

  test('assigning queues three emails and sends none directly', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env);

    const rows = queued(env);
    assert.equal(rows.length, 3, 'the sub, the covered-for teacher, and the admin copy');
    assert.deepEqual(rows.map(r => r.recipient),
      [USERS.omsTeacher, USERS.omsTeacher2, USERS.omsAdmin]);
    rows.forEach(r => assert.equal(r.building, 'OMS', 'every row is tagged for the building admin'));
    assert.equal(env.sentEmails.length, 0,
      'nothing bypasses the queue — that is what makes the building admin the sender');
  });

  test('free-text coverage emails nobody about a person who does not exist', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env, { subbedFor: 'Activity Bus', coveredForEmail: '' });

    const rows = queued(env);
    assert.equal(rows.length, 2, 'just the sub and the admin copy');
    assert.deepEqual(rows.map(r => r.recipient), [USERS.omsTeacher, USERS.omsAdmin]);
  });

  test("the sub's email says assigned, offers no decline, and carries a Record link", () => {
    const env = envFor(USERS.omsAdmin);
    assign(env);
    const subEmail = queued(env)[0];

    assert.ok(/You have been assigned for TST Coverage/.test(subEmail.body));
    assert.ok(/Record My TST Time/.test(subEmail.body));
    assert.ok(/If you have a conflict, contact/.test(subEmail.body));
    assert.ok(!/Decline/i.test(subEmail.body), 'there is no decline path any more');
    assert.ok(/action=record/.test(subEmail.body), 'the button points at the signed Record link');
    assert.ok(!/TST Calendar/.test(subEmail.body),
      'Phase 1 has no calendar, so the email must not claim one');
  });

  test('the covered-for teacher gets their own email, not a copy of the sub\'s', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env);
    const coveredEmail = queued(env)[1];

    assert.ok(/has been assigned to cover your/.test(coveredEmail.body));
    assert.ok(!/action=record/.test(coveredEmail.body),
      'only the person covering gets a Record link');
  });

  test('the note goes only where the admin ticked it', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env, { note: 'Plans are on the desk', noteToSub: true, noteToCovered: false });
    const rows = queued(env);

    assert.ok(/Plans are on the desk/.test(rows[0].body), 'the sub was ticked');
    assert.ok(!/Plans are on the desk/.test(rows[1].body), 'the covered-for teacher was not');
    assert.ok(/Plans are on the desk/.test(rows[2].body), 'the admin copy always shows it');
  });

  test('a note nobody was ticked for is not emailed to anyone', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env, { note: 'internal reminder' });
    const rows = queued(env);
    assert.ok(!/internal reminder/.test(rows[0].body));
    assert.ok(!/internal reminder/.test(rows[1].body));
  });

  test('a note with HTML in it is escaped, not rendered', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env, { note: '<img src=x onerror=alert(1)>', noteToSub: true });
    const body = queued(env)[0].body;
    assert.ok(!/<img/.test(body), 'the tag must not survive into the email');
    assert.ok(/&lt;img/.test(body));
  });

  // ---- getAssignments --------------------------------------------------------

  test('a teacher cannot read the assignments queue', () => {
    const env = envFor(USERS.omsTeacher);
    assert.rejected(env.attempt('getAssignments', 'OMS'), /admin access required/i);
  });

  test("an admin asking for another building's queue falls back to their own", () => {
    const env = envFor(USERS.omsAdmin);
    assign(env);
    const list = env.run('getAssignments', 'OHS');
    assert.ok(list.every(a => a.building === 'OMS'), 'no read is ever district-wide');
  });

  test('a Super Admin may name any building', () => {
    const env = envFor(USERS.superAdmin);
    assign(env, {
      teacherEmail: USERS.ohsTeacher, teacherName: 'Hank Teacher', subbedFor: 'Otto Admin',
      coveredForEmail: USERS.ohsAdmin, period: 'Period 1', building: 'OHS'
    });
    assert.equal(env.run('getAssignments', 'OHS').length, 1);
    assert.equal(env.run('getAssignments', 'OMS').length, 0);
  });

  // ---- The badge -------------------------------------------------------------

  test('the badge counts only coverage that already happened with nothing recorded', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env, { date: ymd(3) });                        // upcoming — not actionable
    assign(env, { date: ymd(-2), period: OMS_PERIOD_2, force: true }); // overdue

    assert.equal(env.run('getDashboardCounts', 'OMS').assignments, 1);
  });

  test('the badge matches the list under it', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env, { date: ymd(-2) });
    const outstanding = env.run('getAssignments', 'OMS').filter(a => a.outstanding).length;
    assert.equal(env.run('getDashboardCounts', 'OMS').assignments, outstanding);
  });

  test('recording an assignment clears it from the badge', () => {
    const env = envFor(USERS.omsAdmin);
    const id = assign(env, { date: ymd(-2) });
    assert.equal(env.run('getDashboardCounts', 'OMS').assignments, 1);

    env.run('recordAssignment', id, USERS.omsTeacher);
    assert.equal(env.run('getDashboardCounts', 'OMS').assignments, 0);
  });

  // ---- getMyAssignments ------------------------------------------------------

  test('a teacher sees what they are covering and what is covered for them', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env);

    const subView = createEnv({ activeUser: USERS.omsTeacher, sheets: env.spreadsheet.sheets
      .reduce((acc, s) => { acc[s.name] = s.values; return acc; }, {}) });
    const mine = subView.run('getMyAssignments', USERS.omsTeacher);
    assert.equal(mine.length, 1);
    assert.equal(mine[0].role, 'covering');

    const coveredView = createEnv({ activeUser: USERS.omsTeacher2, sheets: env.spreadsheet.sheets
      .reduce((acc, s) => { acc[s.name] = s.values; return acc; }, {}) });
    const theirs = coveredView.run('getMyAssignments', USERS.omsTeacher2);
    assert.equal(theirs.length, 1);
    assert.equal(theirs[0].role, 'covered');
  });

  test("a teacher cannot read another teacher's assignments", () => {
    const env = envFor(USERS.omsTeacher);
    assert.rejected(env.attempt('getMyAssignments', USERS.omsTeacher2), /admin access required/i);
  });

  test('an admin can read them for a teacher they manage (View As)', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env);
    assert.allowed(env.attempt('getMyAssignments', USERS.omsTeacher));
  });

  test('an admin from another building cannot', () => {
    const env = envFor(USERS.ohsAdmin);
    assert.rejected(env.attempt('getMyAssignments', USERS.omsTeacher), /your own building/i);
  });

  // ---- recordAssignment ------------------------------------------------------

  test('recording files a pending earned request and marks the row', () => {
    const env = envFor(USERS.omsAdmin);
    const id = assign(env, { date: ymd(-1) });
    const before = env.sheet('TST Approvals (New)').values.length;

    env.run('recordAssignment', id, USERS.omsTeacher);

    const approvals = env.sheet('TST Approvals (New)').values;
    assert.equal(approvals.length, before + 1, 'exactly one earned row');
    const row = approvals[approvals.length - 1];
    assert.equal(row[0], USERS.omsTeacher);
    assert.equal(row[5], OMS_PERIOD);
    assert.equal(row[7], 1, 'the hours the admin assigned');
    assert.equal(row[8], false, 'still pending your approval');

    assert.equal(env.run('getAssignments', 'OMS')[0].status, 'Recorded');
  });

  test('recording twice cannot create a second request', () => {
    const env = envFor(USERS.omsAdmin);
    const id = assign(env, { date: ymd(-1) });
    const before = env.sheet('TST Approvals (New)').values.length;

    env.run('recordAssignment', id, USERS.omsTeacher);
    env.run('recordAssignment', id, USERS.omsTeacher);

    assert.equal(env.sheet('TST Approvals (New)').values.length, before + 1);
  });

  test('a teacher cannot record an assignment that belongs to someone else', () => {
    const env = envFor(USERS.omsAdmin);
    const id = assign(env);

    const other = createEnv({ activeUser: USERS.omsTeacher2, sheets: env.spreadsheet.sheets
      .reduce((acc, s) => { acc[s.name] = s.values; return acc; }, {}) });
    assert.rejected(other.attempt('recordAssignment', id, USERS.omsTeacher2),
      /belongs to another staff member/i);
    assert.rejected(other.attempt('recordAssignment', id, USERS.omsTeacher),
      /admin access required/i);
  });

  test('an admin may record on a teacher\'s behalf, and the trail says so', () => {
    const env = envFor(USERS.omsAdmin);
    const id = assign(env, { date: ymd(-20) }); // past the emailed link's 14 days
    assert.allowed(env.attempt('recordAssignment', id, USERS.omsTeacher));

    const row = assignmentSheetRows(env)[0];
    assert.equal(row[16], 'Recorded');
    assert.ok(/^admin:/.test(row[18]), 'recorded-by distinguishes the admin from the teacher');
  });

  test('an admin from another building may not record it', () => {
    const env = envFor(USERS.omsAdmin);
    const id = assign(env);

    const otto = createEnv({ activeUser: USERS.ohsAdmin, sheets: env.spreadsheet.sheets
      .reduce((acc, s) => { acc[s.name] = s.values; return acc; }, {}) });
    assert.rejected(otto.attempt('recordAssignment', id, USERS.omsTeacher), /your own building/i);
  });

  // ---- cancelAssignment ------------------------------------------------------

  test('cancelling marks the row and emails both staff', () => {
    const env = envFor(USERS.omsAdmin);
    const id = assign(env);
    env.sheet('Email Queue').values = env.sheet('Email Queue').values.slice(0, 1);

    assert.allowed(env.attempt('cancelAssignment', id));
    assert.equal(assignmentSheetRows(env)[0][16], 'Cancelled');

    const rows = queued(env);
    assert.deepEqual(rows.map(r => r.recipient), [USERS.omsTeacher, USERS.omsTeacher2]);
    assert.ok(/has been cancelled/.test(rows[0].body));
  });

  test('cancelling a recorded assignment removes its pending request', () => {
    const env = envFor(USERS.omsAdmin);
    const id = assign(env, { date: ymd(-1) });
    env.run('recordAssignment', id, USERS.omsTeacher);
    const withRequest = env.sheet('TST Approvals (New)').values.length;

    assert.allowed(env.attempt('cancelAssignment', id));
    assert.equal(env.sheet('TST Approvals (New)').values.length, withRequest - 1);
  });

  test('cancelling refuses once the hours are approved, and changes nothing', () => {
    const env = envFor(USERS.omsAdmin);
    const id = assign(env, { date: ymd(-1) });
    env.run('recordAssignment', id, USERS.omsTeacher);

    const approvals = env.sheet('TST Approvals (New)');
    const rowIndex = approvals.values.length; // 1-based: the row just appended
    env.run('approveEarnedRow', rowIndex, { send: false });

    assert.rejected(env.attempt('cancelAssignment', id), /already approved/i);
    assert.equal(assignmentSheetRows(env)[0][16], 'Recorded', 'the assignment is untouched');
    assert.equal(approvals.values.length, rowIndex, 'and the approved hours survive');
  });

  test('a teacher cannot cancel, and an admin cannot cancel another building', () => {
    const env = envFor(USERS.omsAdmin);
    const id = assign(env);
    const all = env.spreadsheet.sheets.reduce((acc, s) => { acc[s.name] = s.values; return acc; }, {});

    const tina = createEnv({ activeUser: USERS.omsTeacher, sheets: all });
    assert.rejected(tina.attempt('cancelAssignment', id), /admin access required/i);

    const otto = createEnv({ activeUser: USERS.ohsAdmin, sheets: all });
    assert.rejected(otto.attempt('cancelAssignment', id), /your own building/i);
  });

  // ---- Reminders -------------------------------------------------------------

  test('the nudge only picks up coverage that has passed with nothing recorded', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env, { date: ymd(2) });                                     // upcoming
    assign(env, { date: ymd(-2), period: OMS_PERIOD_2, force: true }); // due a nudge
    const recorded = assign(env, { date: ymd(-3), period: 'Period 3 - 9:52 - 10:39', force: true });
    env.run('recordAssignment', recorded, USERS.omsTeacher);
    env.sheet('Email Queue').values = env.sheet('Email Queue').values.slice(0, 1);

    assert.equal(env.callInternal('nudgeOutstandingAssignments_'), 1);
    const rows = queued(env);
    assert.equal(rows.length, 1);
    assert.ok(/Reminder: Record your TST time/.test(rows[0].subject));
  });

  test('the nudge is one per assignment, not one per day', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env, { date: ymd(-2) });
    assert.equal(env.callInternal('nudgeOutstandingAssignments_'), 1);
    assert.equal(env.callInternal('nudgeOutstandingAssignments_'), 0, 'the second run finds nothing');
  });

  test('the nudge stops once the emailed link has expired', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env, { date: ymd(-30) });
    assert.equal(env.callInternal('nudgeOutstandingAssignments_'), 0,
      'chasing a link that no longer works would only confuse people');
  });

  test('a teacher cannot trigger the nudge from the client', () => {
    const env = envFor(USERS.omsTeacher);
    assert.rejected(env.attempt('nudgeOutstandingAssignments'), /admin access required/i);
    assert.rejected(env.attempt('nudgeOutstandingAssignments', { authMode: 'FULL', triggerUid: '1' }),
      /admin access required/i);
  });

  test('the installed trigger may run it', () => {
    const env = envFor(USERS.omsAdmin);
    assert.allowed(env.attempt('nudgeOutstandingAssignments',
      { authMode: env.AuthMode.FULL, triggerUid: '1' }));
  });

  test('the manual Remind button re-sends and marks the row', () => {
    const env = envFor(USERS.omsAdmin);
    const id = assign(env, { date: ymd(-1) });
    env.sheet('Email Queue').values = env.sheet('Email Queue').values.slice(0, 1);

    assert.allowed(env.attempt('remindAssignment', id));
    assert.equal(queued(env).length, 1);
    assert.equal(env.callInternal('nudgeOutstandingAssignments_'), 0,
      'a manual reminder counts as the one nudge');
  });

  test('Remind refuses an assignment that is already recorded', () => {
    const env = envFor(USERS.omsAdmin);
    const id = assign(env, { date: ymd(-1) });
    env.run('recordAssignment', id, USERS.omsTeacher);
    assert.rejected(env.attempt('remindAssignment', id), /already recorded/i);
  });

  // ---- The signed Record link ------------------------------------------------

  function recordLinkParams(env, id, email) {
    // Built the same way sendAssignmentEmails_ builds it.
    const params = { action: 'record', id: id, tEmail: email };
    params.sig = env.callInternal('assignmentSignature_', params);
    return params;
  }

  test('a valid link records the assignment for the signed-in teacher', () => {
    const admin = envFor(USERS.omsAdmin);
    const id = assign(admin, { date: ymd(-1) });
    const all = admin.spreadsheet.sheets.reduce((acc, s) => { acc[s.name] = s.values; return acc; }, {});

    const tina = createEnv({ activeUser: USERS.omsTeacher, sheets: all });
    const page = tina.run('doGet', { parameter: recordLinkParams(tina, id, USERS.omsTeacher) });
    assert.ok(/TST Time Recorded/.test(page.getContent()));
    assert.equal(tina.run('getMyAssignments', USERS.omsTeacher)[0].status, 'Recorded');
  });

  test('a tampered link is refused', () => {
    const admin = envFor(USERS.omsAdmin);
    const id = assign(admin, { date: ymd(-1) });
    const all = admin.spreadsheet.sheets.reduce((acc, s) => { acc[s.name] = s.values; return acc; }, {});

    const tina = createEnv({ activeUser: USERS.omsTeacher, sheets: all });
    const params = recordLinkParams(tina, id, USERS.omsTeacher);
    params.tEmail = USERS.omsTeacher2; // signature no longer covers this
    const page = tina.callInternal('doGet', { parameter: params });
    assert.ok(/Link not valid/.test(page.getContent()));
  });

  test('opening someone else\'s link while signed in as yourself is refused', () => {
    const admin = envFor(USERS.omsAdmin);
    const id = assign(admin, { date: ymd(-1) });
    const all = admin.spreadsheet.sheets.reduce((acc, s) => { acc[s.name] = s.values; return acc; }, {});

    const ted = createEnv({ activeUser: USERS.omsTeacher2, sheets: all });
    const page = ted.callInternal('doGet', { parameter: recordLinkParams(ted, id, USERS.omsTeacher) });
    assert.ok(/Wrong account/.test(page.getContent()));
  });

  test('the link stops working 14 days after the coverage date', () => {
    const admin = envFor(USERS.omsAdmin);
    const id = assign(admin, { date: ymd(-20) });
    const all = admin.spreadsheet.sheets.reduce((acc, s) => { acc[s.name] = s.values; return acc; }, {});

    const tina = createEnv({ activeUser: USERS.omsTeacher, sheets: all });
    const page = tina.callInternal('doGet', { parameter: recordLinkParams(tina, id, USERS.omsTeacher) });
    assert.ok(/expired/i.test(page.getContent()));
    assert.equal(tina.run('getMyAssignments', USERS.omsTeacher)[0].status, 'Assigned',
      'an expired link must not record anything');
  });

  test('the in-app button still works after the link has expired', () => {
    const admin = envFor(USERS.omsAdmin);
    const id = assign(admin, { date: ymd(-20) });
    const all = admin.spreadsheet.sheets.reduce((acc, s) => { acc[s.name] = s.values; return acc; }, {});

    // Different risks: the link is a bearer token in an inbox, this is an
    // authenticated action by that teacher on a row the admin created.
    const tina = createEnv({ activeUser: USERS.omsTeacher, sheets: all });
    assert.allowed(tina.attempt('recordAssignment', id, USERS.omsTeacher));
  });

  test('a cancelled assignment cannot be recorded through the link', () => {
    const admin = envFor(USERS.omsAdmin);
    const id = assign(admin, { date: ymd(-1) });
    admin.run('cancelAssignment', id);
    const all = admin.spreadsheet.sheets.reduce((acc, s) => { acc[s.name] = s.values; return acc; }, {});

    const tina = createEnv({ activeUser: USERS.omsTeacher, sheets: all });
    const page = tina.callInternal('doGet', { parameter: recordLinkParams(tina, id, USERS.omsTeacher) });
    assert.ok(/cancelled/i.test(page.getContent()));
  });

  test('the old Accept and Decline links say so instead of failing as forged', () => {
    const env = envFor(USERS.omsTeacher);
    ['accept', 'reject'].forEach(action => {
      const page = env.callInternal('doGet', { parameter: { action: action, tEmail: USERS.omsTeacher } });
      assert.ok(/no longer valid/i.test(page.getContent()),
        `${action} should explain itself, not look broken`);
    });
  });

  // ---- The queue sends as the building admin ---------------------------------

  test("a Super Admin's trigger no longer sends another building's mail", () => {
    const env = createEnv({
      activeUser: USERS.superAdmin,
      sheets: sheets({
        'Email Queue': [
          ['Timestamp', 'Recipient', 'Subject', 'Body', 'Building', 'Status', 'LastUpdated', 'Options'],
          [new Date(), USERS.ohsTeacher, 'TST Coverage Assignment', '<p>x</p>', 'OHS', 'Pending', '', '{}']
        ]
      })
    });

    // Sam Super is assigned to OMS. Sweeping OHS used to race Otto's trigger every
    // minute and make the From name a coin flip.
    assert.allowed(env.attempt('processEmailQueue'));
    assert.equal(env.sentEmails.length, 0);
  });

  test("the building's own admin still sends it", () => {
    const env = createEnv({
      activeUser: USERS.ohsAdmin,
      sheets: sheets({
        'Email Queue': [
          ['Timestamp', 'Recipient', 'Subject', 'Body', 'Building', 'Status', 'LastUpdated', 'Options'],
          [new Date(), USERS.ohsTeacher, 'TST Coverage Assignment', '<p>x</p>', 'OHS', 'Pending', '', '{}']
        ]
      })
    });

    assert.allowed(env.attempt('processEmailQueue'));
    assert.equal(env.sentEmails.length, 1);
    assert.equal(env.sentEmails[0].name, 'Otto Admin', 'the From name is the building admin');
    assert.equal(env.sentEmails[0].replyTo, USERS.ohsAdmin);
  });

  // ---- Year end --------------------------------------------------------------

  test('finalizing a year sweeps that building\'s assignment rows into the archive', () => {
    const env = envFor(USERS.superAdmin);
    assign(env, { date: ymd(-5) });
    assign(env, {
      teacherEmail: USERS.ohsTeacher, teacherName: 'Hank Teacher', subbedFor: 'Otto Admin',
      coveredForEmail: USERS.ohsAdmin, period: 'Period 1', building: 'OHS', date: ymd(-5)
    });

    env.run('finalizeSchoolYear', '2025-2026', 'OMS', true);

    assert.equal(assignmentSheetRows(env).length, 1, 'only OMS was finalized');
    assert.equal(assignmentSheetRows(env)[0][2], 'OHS');
    const arch = env.sheet('TST Assignments Archive');
    assert.equal(arch.values.length, 2, 'header + the archived OMS row');
    assert.equal(arch.values[1][arch.values[1].length - 1], '2025-2026');
  });

  // ---- Private helpers stay off the client surface ---------------------------

  ['recordAssignment_', 'assignmentsFor_', 'assignmentRows_', 'findAssignment_',
   'sendAssignmentEmails_', 'nudgeOutstandingAssignments_', 'handleAssignmentRecord_',
   'archiveAssignmentsForBuilding_'].forEach(fn => {
    test(`${fn} is not reachable from google.script.run`, () => {
      const env = envFor(USERS.omsTeacher);
      assert.rejected(env.attempt(fn), /private/i);
      assert.equal(typeof env.context[fn], 'function', 'but server code can still call it');
    });
  });
};
