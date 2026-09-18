/**
 * The per-building TST calendar.
 *
 * The event is created by the building's own admin through their trigger, not by
 * the web app — the app runs as whoever deployed it, so an event it made would be
 * owned by the deployer rather than by the admin whose calendar it is.
 *
 * That ordering is the point of most of what follows: the event is built first and
 * the emails go afterwards, so the "added to the ... TST Calendar" line only ever
 * appears when there is really something to look at, and a calendar that fails
 * still lets the assignment through.
 */

const { createEnv } = require('./apps_script_env');
const { USERS, sheets } = require('./fixtures');

const OMS_CAL = 'oms-tst@group.calendar.google.com';
const OHS_CAL = 'ohs-tst@group.calendar.google.com';

const OMS_PERIOD = 'Period 1 - 8:10 - 8:57';

/** yyyy-MM-dd `offset` days from today. */
function ymd(offset) {
  const d = new Date();
  d.setDate(d.getDate() + (offset || 0));
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

/** The next date falling on `weekday` (0 = Sunday). */
function nextWeekday(weekday) {
  const d = new Date();
  while (d.getDay() !== weekday) d.setDate(d.getDate() + 1);
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

/** An App Config sheet, so a test can say exactly what each building is set up with. */
function appConfig(overrides) {
  const base = {
    OMS: {
      name: 'Orono Middle School', carryOverMax: 12, scheduleType: 'periods',
      periods: [OMS_PERIOD, 'Period 2 - 9:01 - 9:48'],
      coverageTypes: [{ label: 'Full Period', value: 1 }, { label: 'Half Period', value: 0.5 }]
    },
    OHS: {
      name: 'Orono High School', carryOverMax: 12, scheduleType: 'periods',
      periods: ['Period 1'],
      coverageTypes: [{ label: 'Full Period', value: 1 }]
    }
  };
  Object.keys(overrides || {}).forEach(code => {
    base[code] = Object.assign({}, base[code], overrides[code]);
  });
  return [['Building', 'Config_JSON']]
    .concat(Object.keys(base).map(code => [code, JSON.stringify(base[code])]));
}

function envWith(user, configOverrides, calendars) {
  return createEnv({
    activeUser: user,
    sheets: sheets({ 'App Config': appConfig(configOverrides) }),
    calendars: calendars || {}
  });
}

/** An OMS environment with a working calendar, as the OMS admin. */
function omsEnv(extraConfig) {
  return envWith(
    USERS.omsAdmin,
    { OMS: Object.assign({ calendarId: OMS_CAL, calendarName: 'OMS TST Calendar' }, extraConfig || {}) },
    { [OMS_CAL]: { name: 'OMS TST Calendar' } }
  );
}

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

/** Runs the building admin's trigger: calendar work, then the mail queue. */
function runTrigger(env) {
  return env.run('processEmailQueue', { authMode: env.AuthMode.FULL, triggerUid: 'trigger-1' });
}

function assignmentRow(env) {
  return env.sheet('TST Assignments').values[1];
}

/** How many emails are waiting. The sheet only exists once something is queued. */
function queuedCount(env) {
  const sheet = env.sheet('Email Queue');
  return sheet ? sheet.values.length - 1 : 0;
}

const CAL_STATUS = 21; // Calendar Status column, 0-based
const NOTIFIED = 22;   // Notified TS

exports.name = 'Assignment calendar';

exports.run = function ({ test, assert }) {

  // ---- A building with no calendar is untouched ------------------------------

  test('with no calendar configured, the emails go out immediately', () => {
    const env = envWith(USERS.omsAdmin, {}, {});
    env.run('assignCoverage', payload());

    assert.equal(queuedCount(env), 3, 'all three queued at once');
    assert.equal(assignmentRow(env)[CAL_STATUS], '', 'nothing calendar-related is pending');
    assert.ok(assignmentRow(env)[NOTIFIED], 'and they are marked as notified');
  });

  test('a building with no calendar never mentions one', () => {
    const env = envWith(USERS.omsAdmin, {}, {});
    env.run('assignCoverage', payload());
    env.sheet('Email Queue').values.slice(1).forEach(r => {
      assert.ok(!/TST Calendar/.test(r[3]), 'no calendar sentence when there is no calendar');
    });
  });

  // ---- With a calendar, the event comes first --------------------------------

  test('assigning waits for the event rather than emailing first', () => {
    const env = omsEnv();
    const res = env.run('assignCoverage', payload());

    assert.ok(res.pendingCalendar);
    assert.equal(assignmentRow(env)[CAL_STATUS], 'Pending');
    assert.equal(queuedCount(env), 0,
      'nothing is emailed until the event exists — otherwise it would claim one that does not');
    assert.equal(env.calendarEvents.length, 0, 'the web app must not create it; the admin does');
  });

  test("the building admin's trigger creates the event, then sends the emails", () => {
    const env = omsEnv();
    env.run('assignCoverage', payload());
    runTrigger(env);

    assert.equal(env.calendarEvents.length, 1);
    const event = env.calendarEvents[0];
    assert.equal(event.calendarId, OMS_CAL);
    assert.equal(event.title, 'TST: Tina Teacher covering Ted Teacher — Period 1');

    assert.equal(assignmentRow(env)[CAL_STATUS], 'Created');
    assert.equal(env.sentEmails.length, 3, 'and now everyone is told');
  });

  test('the event runs at the period times, not all day', () => {
    const env = omsEnv();
    const date = ymd(3);
    env.run('assignCoverage', payload({ date: date }));
    runTrigger(env);

    const event = env.calendarEvents[0];
    const pad = n => String(n).padStart(2, '0');
    const hhmm = d => pad(d.getHours()) + ':' + pad(d.getMinutes());
    assert.equal(hhmm(event.start), '08:10');
    assert.equal(hhmm(event.end), '08:57');
    assert.equal(event.start.getDate(), Number(date.split('-')[2]), 'and on the right day');
  });

  test('both teachers are guests, with no Google invite', () => {
    const env = omsEnv();
    env.run('assignCoverage', payload());
    runTrigger(env);

    const options = env.calendarEvents[0].options;
    assert.equal(options.guests, USERS.omsTeacher + ',' + USERS.omsTeacher2);
    assert.equal(options.sendInvites, false,
      "Google's own invite carries a Yes/No/Maybe prompt — a decline button in a flow that has none");
  });

  test('free-text coverage puts only the person covering on the event', () => {
    const env = omsEnv();
    env.run('assignCoverage', payload({ subbedFor: 'Activity Bus', coveredForEmail: '' }));
    runTrigger(env);

    assert.equal(env.calendarEvents[0].options.guests, USERS.omsTeacher);
  });

  test('the emails name the calendar the admin typed', () => {
    const env = omsEnv();
    env.run('assignCoverage', payload());
    runTrigger(env);

    const subEmail = env.sentEmails[0];
    assert.ok(/added to the OMS TST Calendar/.test(subEmail.htmlBody));
  });

  test('a day-specific schedule decides the time', () => {
    const env = envWith(
      USERS.ohsAdmin,
      {
        OHS: {
          calendarId: OHS_CAL,
          calendarName: 'OHS TST Calendar',
          periodTimes: { 'Period 1': { start: '08:00', end: '08:47' } },
          dayGroups: [{ name: 'TTh', days: ['Tue', 'Thu'], times: { 'Period 1': { start: '08:00', end: '09:30' } } }]
        }
      },
      { [OHS_CAL]: { name: 'OHS TST Calendar' } }
    );

    env.run('assignCoverage', {
      teacherEmail: USERS.ohsTeacher, teacherName: 'Hank Teacher', subbedFor: 'Otto Admin',
      coveredForEmail: USERS.ohsAdmin, date: nextWeekday(2), period: 'Period 1',
      amount: 1, amountType: 'Full Period', building: 'OHS'
    });
    runTrigger(env);

    const pad = n => String(n).padStart(2, '0');
    const event = env.calendarEvents[0];
    assert.equal(pad(event.end.getHours()) + ':' + pad(event.end.getMinutes()), '09:30',
      'Tuesday runs the block schedule');
  });

  // ---- Only the building's own admin does its calendar work ------------------

  test("a Super Admin's trigger leaves another building's event alone", () => {
    const env = envWith(
      USERS.superAdmin, // assigned to OMS
      { OHS: { calendarId: OHS_CAL, periodTimes: { 'Period 1': { start: '08:00', end: '08:47' } } } },
      { [OHS_CAL]: { name: 'OHS TST Calendar' } }
    );
    env.run('assignCoverage', {
      teacherEmail: USERS.ohsTeacher, teacherName: 'Hank Teacher', subbedFor: 'Otto Admin',
      coveredForEmail: USERS.ohsAdmin, date: ymd(3), period: 'Period 1',
      amount: 1, amountType: 'Full Period', building: 'OHS'
    });

    runTrigger(env);
    assert.equal(env.calendarEvents.length, 0, 'OHS events belong to the OHS admin');
    assert.equal(assignmentRow(env)[CAL_STATUS], 'Pending', 'it waits for them');
  });

  // ---- Failure never blocks the assignment -----------------------------------

  test('an unreachable calendar still lets the assignment through', () => {
    const env = envWith(
      USERS.omsAdmin,
      { OMS: { calendarId: 'wrong-id@group.calendar.google.com' } },
      {} // this account cannot open that calendar
    );
    env.run('assignCoverage', payload());
    runTrigger(env);

    assert.ok(/^Failed/.test(assignmentRow(env)[CAL_STATUS]));
    assert.equal(env.calendarEvents.length, 0);

    const to = env.sentEmails.map(m => m.to);
    assert.ok(to.indexOf(USERS.omsTeacher) > -1, 'the person covering is still told');
    assert.ok(to.indexOf(USERS.omsTeacher2) > -1, 'so is the person being covered');
    assert.ok(env.sentEmails.some(m => /calendar entry failed/i.test(m.subject)),
      'and the admin hears about the calendar');
  });

  test('a failed event never claims to be on the calendar', () => {
    const env = envWith(USERS.omsAdmin, { OMS: { calendarId: 'wrong-id@x' } }, {});
    env.run('assignCoverage', payload());
    runTrigger(env);

    const subEmail = env.sentEmails.find(m => m.to === USERS.omsTeacher);
    assert.ok(!/added to the/.test(subEmail.htmlBody));
  });

  test('a period with no configured time reports that, rather than guessing one', () => {
    const env = envWith(
      USERS.ohsAdmin,
      { OHS: { calendarId: OHS_CAL } }, // periods, but no times for them
      { [OHS_CAL]: { name: 'OHS TST Calendar' } }
    );
    env.run('assignCoverage', {
      teacherEmail: USERS.ohsTeacher, teacherName: 'Hank Teacher', subbedFor: 'Otto Admin',
      coveredForEmail: USERS.ohsAdmin, date: ymd(3), period: 'Period 1',
      amount: 1, amountType: 'Full Period', building: 'OHS'
    });
    runTrigger(env);

    assert.ok(/No start and end time is set/.test(assignmentRow(env)[CAL_STATUS]));
    assert.equal(env.sentEmails.length, 4, 'three assignment emails plus the failure alert');
  });

  // ---- Cancelling ------------------------------------------------------------

  test('cancelling removes the event, then says so', () => {
    const env = omsEnv();
    const id = env.run('assignCoverage', payload()).id;
    runTrigger(env);
    assert.equal(env.calendarEvents.length, 1);
    env.sentEmails.length = 0;

    const res = env.run('cancelAssignment', id);
    assert.ok(res.pendingCalendar);
    assert.equal(env.calendarEvents.length, 1, 'the web app cannot remove the admin\'s event');
    assert.equal(env.sentEmails.length, 0, 'and must not announce a removal that has not happened');

    runTrigger(env);
    assert.equal(env.calendarEvents.length, 0, 'now it is gone');
    assert.equal(assignmentRow(env)[CAL_STATUS], 'Deleted');
    assert.equal(env.sentEmails.length, 2, 'and both staff are told');
    assert.ok(/removed from the OMS TST Calendar/.test(env.sentEmails[0].htmlBody));
  });

  test('cancelling before anyone was told emails nobody', () => {
    const env = omsEnv();
    const id = env.run('assignCoverage', payload()).id;
    // The trigger has not run, so the assignment emails never went out.
    env.run('cancelAssignment', id);
    runTrigger(env);

    assert.equal(env.sentEmails.length, 0,
      'a cancellation would be the first they ever heard of it');
    assert.equal(env.calendarEvents.length, 0);
  });

  // ---- The test-event button -------------------------------------------------

  test('a teacher cannot fire the calendar test', () => {
    const env = envWith(USERS.omsTeacher, { OMS: { calendarId: OMS_CAL } }, { [OMS_CAL]: { name: 'x' } });
    assert.rejected(env.attempt('sendTestCalendarEvent', 'OMS'), /admin access required/i);
    assert.rejected(env.attempt('getCalendarTestResult', 'OMS'), /admin access required/i);
  });

  test("an admin cannot test another building's calendar", () => {
    const env = envWith(USERS.omsAdmin, { OHS: { calendarId: OHS_CAL } }, { [OHS_CAL]: { name: 'x' } });
    assert.rejected(env.attempt('sendTestCalendarEvent', 'OHS'), /your own building/i);
  });

  test('testing with no calendar set says so up front', () => {
    const env = envWith(USERS.omsAdmin, {}, {});
    assert.rejected(env.attempt('sendTestCalendarEvent', 'OMS'), /No calendar ID is set/i);
  });

  test('a passing test reports the calendar it actually reached', () => {
    const env = omsEnv();
    env.run('sendTestCalendarEvent', 'OMS');
    assert.equal(env.run('getCalendarTestResult', 'OMS').state, 'pending');

    runTrigger(env);

    const result = env.run('getCalendarTestResult', 'OMS');
    assert.equal(result.state, 'ok');
    assert.equal(result.calendarName, 'OMS TST Calendar',
      'the real name, so a wrong display name in Settings is visible');
    assert.equal(result.by, USERS.omsAdmin, 'proving whose trigger did it');
    assert.equal(env.calendarEvents.length, 0, 'the throwaway event removes itself');
  });

  test('a failing test reports why', () => {
    const env = envWith(USERS.omsAdmin, { OMS: { calendarId: 'nope@group.calendar.google.com' } }, {});
    env.run('sendTestCalendarEvent', 'OMS');
    runTrigger(env);

    const result = env.run('getCalendarTestResult', 'OMS');
    assert.equal(result.state, 'failed');
    assert.ok(/Calendar not found/.test(result.error));
  });

  test('the test runs on the admin\'s trigger, not the web app', () => {
    const env = omsEnv();
    env.run('sendTestCalendarEvent', 'OMS');
    // No trigger run yet: nothing may have happened.
    assert.equal(env.run('getCalendarTestResult', 'OMS').state, 'pending');
    assert.equal(env.calendarEvents.length, 0);
  });

  // ---- Private helpers stay off the client surface ---------------------------

  ['processPendingAssignments_', 'finishAssignmentCreation_', 'finishAssignmentCancellation_',
   'runCalendarTests_', 'calendarIdFor_', 'assignmentEventWindow_'].forEach(fn => {
    test(`${fn} is not reachable from google.script.run`, () => {
      const env = envWith(USERS.omsTeacher, {}, {});
      assert.rejected(env.attempt(fn), /private/i);
      assert.equal(typeof env.context[fn], 'function', 'but server code can still call it');
    });
  });
};
