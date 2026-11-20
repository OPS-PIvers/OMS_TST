/**
 * Editing a coverage assignment from the Assignments tab.
 *
 * Same three questions as every endpoint (a Teacher, an admin from another
 * building, a Super Admin), then the behaviour the design turns on: the event is
 * changed by the building admin's trigger before anyone is emailed, the admin
 * chooses who hears about it, and the emails list only what changed.
 */

const { createEnv } = require('./apps_script_env');
const { USERS, sheets } = require('./fixtures');

const OMS_CAL = 'oms-tst@group.calendar.google.com';
const P1 = 'Period 1 - 8:10 - 8:57';
const P2 = 'Period 2 - 9:01 - 9:48';

function ymd(offset) {
  const d = new Date();
  d.setDate(d.getDate() + (offset || 0));
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

function appConfig(withCalendar) {
  const oms = {
    name: 'Orono Middle School', carryOverMax: 12, scheduleType: 'periods',
    periods: [P1, P2],
    coverageTypes: [{ label: 'Full Period', value: 1 }, { label: 'Half Period', value: 0.5 }]
  };
  if (withCalendar) Object.assign(oms, { calendarId: OMS_CAL, calendarName: 'OMS TST Calendar' });
  const ohs = { name: 'Orono High School', carryOverMax: 12, scheduleType: 'periods', periods: ['Period 1'] };
  return [['Building', 'Config_JSON'], ['OMS', JSON.stringify(oms)], ['OHS', JSON.stringify(ohs)]];
}

/** An environment whose signed-in user can be switched with env.as(email). */
function makeEnv(withCalendar) {
  const opts = {
    activeUser: USERS.omsAdmin,
    sheets: sheets({ 'App Config': appConfig(withCalendar) }),
    calendars: withCalendar ? { [OMS_CAL]: { name: 'OMS TST Calendar' } } : {}
  };
  const env = createEnv(opts);
  env.as = email => { opts.activeUser = email; return env; };
  return env;
}

function payload(extra) {
  return Object.assign({
    teacherEmail: USERS.omsTeacher,
    teacherName: 'Tina Teacher',
    subbedFor: 'Ted Teacher',
    coveredForEmail: USERS.omsTeacher2,
    date: ymd(3),
    period: P1,
    amount: 1,
    amountType: 'Full Period',
    building: 'OMS'
  }, extra || {});
}

/** What the Edit modal sends: the same fields as assigning, plus who to email. */
function edit(extra) {
  return Object.assign(payload(), { notifySub: true, notifyCovered: true }, extra || {});
}

function runTrigger(env) {
  return env.run('processEmailQueue', { authMode: env.AuthMode.FULL, triggerUid: 'trigger-1' });
}

function row(env) {
  return env.sheet('TST Assignments').values[1];
}

function queued(env) {
  const sheet = env.sheet('Email Queue');
  if (!sheet) return [];
  return sheet.values.slice(1).map(r => ({ recipient: r[1], subject: r[2], body: r[3] }));
}

const COL = { date: 9, period: 10, timeType: 11, hours: 12, coveredFor: 7, coveredForEmail: 8,
  status: 16, nudgedTs: 19, calStatus: 21, editedBy: 26, notice: 27 };

/** Assigns, and with a calendar runs the trigger so the event exists and staff were told. */
function assigned(env, extra) {
  const id = env.run('assignCoverage', payload(extra)).id;
  runTrigger(env);
  env.sentEmails.length = 0;
  return id;
}

exports.name = 'Editing an assignment';

exports.run = function ({ test, assert }) {

  // ---- Who may edit ----------------------------------------------------------

  test('a teacher cannot edit an assignment', () => {
    const env = makeEnv(false);
    const id = assigned(env);
    env.as(USERS.omsTeacher);
    assert.rejected(env.attempt('editAssignment', id, edit({ period: P2 })), /admin access required/i);
    assert.equal(row(env)[COL.period], P1, 'nothing changes');
  });

  test("an admin from another building cannot edit OMS's assignments", () => {
    const env = makeEnv(false);
    const id = assigned(env);
    env.as(USERS.ohsAdmin);
    assert.rejected(env.attempt('editAssignment', id, edit({ period: P2 })), /your own building/i);
    assert.equal(row(env)[COL.period], P1);
  });

  test('a Super Admin can edit any building', () => {
    const env = makeEnv(false);
    const id = assigned(env);
    env.as(USERS.superAdmin);
    assert.allowed(env.attempt('editAssignment', id, edit({ period: P2 })));
    assert.equal(row(env)[COL.period], P2);
  });

  test('recorded and cancelled assignments cannot be edited', () => {
    const env = makeEnv(false);
    const id = assigned(env);
    env.run('cancelAssignment', id);
    assert.rejected(env.attempt('editAssignment', id, edit({ period: P2 })), /already cancelled/i);
  });

  // ---- What it changes -------------------------------------------------------

  test('saves the new details and who made the change', () => {
    const env = makeEnv(false);
    const id = assigned(env);
    const date = ymd(5);
    env.run('editAssignment', id, edit({ date, period: P2, amount: 0.5, amountType: 'Half Period' }));
    const r = row(env);
    assert.equal(r[COL.date], date);
    assert.equal(r[COL.period], P2);
    assert.equal(r[COL.hours], 0.5);
    assert.equal(r[COL.timeType], 'Half Period');
    assert.equal(r[COL.editedBy], 'Amy Admin');
    assert.equal(r[COL.status], 'Assigned', 'still open');
  });

  test('saving with nothing changed does nothing and emails nobody', () => {
    const env = makeEnv(false);
    const id = assigned(env);
    const before = queued(env).length;
    const res = env.run('editAssignment', id, edit());
    assert.ok(res.unchanged);
    assert.equal(queued(env).length, before);
  });

  test('cannot move someone onto a period they already cover that day', () => {
    const env = makeEnv(false);
    const id = assigned(env);
    env.run('assignCoverage', payload({ period: P2, force: true }));
    assert.rejected(env.attempt('editAssignment', id, edit({ period: P2 })), /already assigned/i);
  });

  test('cannot make someone cover for themselves', () => {
    const env = makeEnv(false);
    const id = assigned(env);
    assert.rejected(env.attempt('editAssignment', id,
      edit({ coveredForEmail: USERS.omsTeacher, subbedFor: 'Tina Teacher' })), /themselves/i);
  });

  test('cannot point coverage at staff the admin does not manage', () => {
    const env = makeEnv(false);
    const id = assigned(env);
    assert.rejected(env.attempt('editAssignment', id,
      edit({ coveredForEmail: USERS.ohsTeacher, subbedFor: 'Hank Teacher' })));
  });

  test('moving a past assignment to a future date re-arms the automatic reminder', () => {
    const env = makeEnv(false);
    const id = assigned(env, { date: ymd(-2) });
    env.callInternal('nudgeOutstandingAssignments_');
    assert.ok(row(env)[COL.nudgedTs], 'nudged for the old date');
    env.run('editAssignment', id, edit({ date: ymd(2) }));
    assert.equal(row(env)[COL.nudgedTs], '', 'the reminder belongs after the new date');
  });

  // ---- Emails, no calendar ---------------------------------------------------

  test('with no calendar, the ticked staff are emailed straight away with what changed', () => {
    const env = makeEnv(false);
    const before = queued(env).length;
    const id = assigned(env);
    const start = queued(env).length;
    env.run('editAssignment', id, edit({ period: P2 }));
    const mail = queued(env).slice(start);
    assert.equal(mail.length, 2, 'the person covering and the person covered');
    assert.deepEqual(mail.map(m => m.recipient).sort(), [USERS.omsTeacher2, USERS.omsTeacher].sort());
    mail.forEach(m => {
      assert.ok(/TST Coverage Updated/.test(m.subject));
      assert.ok(/What changed/.test(m.body));
      assert.ok(/Period 1/.test(m.body) && /Period 2/.test(m.body), 'old and new period both shown');
    });
    assert.ok(before >= 0);
  });

  test('unticked staff are not emailed', () => {
    const env = makeEnv(false);
    const id = assigned(env);
    const start = queued(env).length;
    const res = env.run('editAssignment', id, edit({ period: P2, notifySub: false, notifyCovered: false }));
    assert.equal(queued(env).length, start);
    assert.equal(res.emailed, 0);
  });

  test('changing who is covered tells the new person it is arranged and the old one it is off', () => {
    const env = makeEnv(false);
    const id = assigned(env, { coveredForEmail: USERS.omsTeacher2, subbedFor: 'Ted Teacher' });
    const start = queued(env).length;
    env.run('editAssignment', id, edit({
      coveredForEmail: USERS.multiTeacher, subbedFor: 'Mia Multi', notifySub: false
    }));
    const mail = queued(env).slice(start);
    const toNew = mail.find(m => m.recipient === USERS.multiTeacher);
    const toOld = mail.find(m => m.recipient === USERS.omsTeacher2);
    assert.ok(toNew && /Coverage Arranged/.test(toNew.subject));
    assert.ok(toOld && /no longer covering/.test(toOld.body));
  });

  test('a note change only reaches the people the note is for', () => {
    const env = makeEnv(false);
    const id = assigned(env);
    const start = queued(env).length;
    env.run('editAssignment', id, edit({ note: 'Plans on the desk', noteToSub: true, noteToCovered: false }));
    const mail = queued(env).slice(start);
    assert.equal(mail.length, 1, 'the covered teacher sees nothing different, so is not emailed');
    assert.equal(mail[0].recipient, USERS.omsTeacher);
    assert.ok(/Plans on the desk/.test(mail[0].body));
  });

  // ---- With a calendar -------------------------------------------------------

  test('with a calendar, nothing is emailed until the trigger has updated the event', () => {
    const env = makeEnv(true);
    const id = assigned(env);
    assert.equal(env.calendarEvents.length, 1);
    const res = env.run('editAssignment', id, edit({ period: P2 }));
    assert.ok(res.pendingCalendar);
    assert.equal(row(env)[COL.calStatus], 'Pending Update');
    assert.equal(env.sentEmails.length, 0, 'the web app must not touch the event or email first');

    runTrigger(env);
    const event = env.calendarEvents[0];
    const pad = n => String(n).padStart(2, '0');
    assert.equal(pad(event.start.getHours()) + ':' + pad(event.start.getMinutes()), '09:01', 'moved to Period 2');
    assert.equal(event.title, 'TST: Tina Teacher covering Ted Teacher — Period 2');
    assert.equal(env.calendarEvents.length, 1, 'the same event, not a second one');
    assert.equal(row(env)[COL.calStatus], 'Created');
    assert.equal(row(env)[COL.notice], '', 'the queued notice is used up');
    assert.equal(env.sentEmails.length, 2);
    assert.ok(/OMS TST Calendar has been updated/.test(env.sentEmails[0].htmlBody));
  });

  test("the event's guests follow a change of who is covered", () => {
    const env = makeEnv(true);
    const id = assigned(env);
    env.run('editAssignment', id, edit({ coveredForEmail: USERS.multiTeacher, subbedFor: 'Mia Multi' }));
    runTrigger(env);
    assert.deepEqual(env.calendarEvents[0].guests.slice().sort(), [USERS.omsTeacher, USERS.multiTeacher].sort());
  });

  test('two edits before the trigger runs send one email describing both', () => {
    const env = makeEnv(true);
    const id = assigned(env);
    env.run('editAssignment', id, edit({ period: P2, notifyCovered: false }));
    env.run('editAssignment', id, edit({ period: P2, date: ymd(6), notifySub: false, notifyCovered: true }));
    runTrigger(env);
    assert.equal(env.sentEmails.length, 2, 'either edit asking for someone is enough');
    const sub = env.sentEmails.find(m => m.to === USERS.omsTeacher);
    assert.ok(/Period 1/.test(sub.htmlBody), 'compared against what they were first told');
  });

  test('an edit to an assignment nobody has been told about yet sends no change email', () => {
    const env = makeEnv(true);
    const id = env.run('assignCoverage', payload()).id; // event still queued
    env.run('editAssignment', id, edit({ period: P2 }));
    assert.equal(row(env)[COL.calStatus], 'Pending', 'still waiting to be created');
    runTrigger(env);
    assert.equal(env.sentEmails.length, 3, 'just the original three, with the new period');
    assert.ok(env.sentEmails.every(m => !/TST Coverage Updated/.test(m.subject)));
    assert.ok(/Period 2/.test(env.sentEmails[0].htmlBody));
  });

  test('cancelling while an update is queued removes the event and drops the change emails', () => {
    const env = makeEnv(true);
    const id = assigned(env);
    env.run('editAssignment', id, edit({ period: P2 }));
    env.run('cancelAssignment', id);
    assert.equal(row(env)[COL.calStatus], 'Pending Delete');
    runTrigger(env);
    assert.equal(env.calendarEvents.length, 0);
    assert.ok(env.sentEmails.every(m => !/Updated/.test(m.subject)), 'only the cancellation goes out');
  });
};
