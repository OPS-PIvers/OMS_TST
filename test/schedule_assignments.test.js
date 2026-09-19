/**
 * Assigned coverage on the Master Schedule.
 *
 * The amber hourglass only knows about a request someone filed. An assignment the
 * teacher never got around to submitting was invisible on the grid, which is how
 * the same person ends up booked twice for one period. The marker is therefore
 * built from the TST Assignments row — the thing written when the admin assigns —
 * and not from the calendar event, so it works in a building with no calendar too.
 *
 * Each entry carries the month, weekday and period it belongs to, because the grid
 * cell is (month × weekday × period) and that is the one cell a second booking
 * would collide in.
 */

const { createEnv } = require('./apps_script_env');
const { USERS, AVAILABILITY_HEADER, sheets } = require('./fixtures');

const OMS_PERIOD = 'Period 1 - 8:10 - 8:57';
const OHS_PERIOD = 'Period 1';
const OMS_CAL = 'oms-tst@group.calendar.google.com';

const WEEKDAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
  'August', 'September', 'October', 'November', 'December'];

/** A coverage date `offset` days out, with the grid cell it belongs in. */
function coverageDate(offset) {
  const d = new Date();
  d.setDate(d.getDate() + offset);
  const pad = n => String(n).padStart(2, '0');
  return {
    key: `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`,
    month: MONTHS[d.getMonth()],
    weekday: WEEKDAYS[d.getDay()],
    short: `${d.getMonth() + 1}/${d.getDate()}`
  };
}

const SOON = coverageDate(7);
const PAST = coverageDate(-7);

/**
 * Availability covering every month, so the person is on the grid whatever the
 * date arithmetic above lands on.
 */
const AVAILABILITY = [AVAILABILITY_HEADER].concat(
  MONTHS.map(m => [m, 'Mon,Tue,Wed,Thu,Fri', OMS_PERIOD, 'Tina Teacher', USERS.omsTeacher])
).concat(
  MONTHS.map(m => [m, 'Mon,Tue,Wed,Thu,Fri', OHS_PERIOD, 'Hank Teacher', USERS.ohsTeacher])
);

function appConfig(calendarId) {
  const OMS = {
    name: 'Orono Middle School', scheduleType: 'periods',
    periods: [OMS_PERIOD, 'Period 2 - 9:01 - 9:48'],
    coverageTypes: [{ label: 'Full Period', value: 1 }]
  };
  if (calendarId) {
    OMS.calendarId = calendarId;
    OMS.calendarName = 'OMS TST Calendar';
  }
  const OHS = {
    name: 'Orono High School', scheduleType: 'periods',
    periods: [OHS_PERIOD],
    coverageTypes: [{ label: 'Full Period', value: 1 }]
  };
  return [['Building', 'Config_JSON'],
    ['OMS', JSON.stringify(OMS)],
    ['OHS', JSON.stringify(OHS)]];
}

function envFor(user, options) {
  const o = options || {};
  return createEnv({
    activeUser: user,
    sheets: sheets(Object.assign({
      'TST Availability': AVAILABILITY.map(r => r.slice()),
      'App Config': appConfig(o.calendarId)
    }, o.assignments ? { 'TST Assignments': o.assignments } : {})),
    calendars: o.calendars || {}
  });
}

function payload(extra) {
  return Object.assign({
    teacherEmail: USERS.omsTeacher,
    teacherName: 'Tina Teacher',
    subbedFor: 'Ted Teacher',
    coveredForEmail: USERS.omsTeacher2,
    date: SOON.key,
    period: OMS_PERIOD,
    amount: 1,
    amountType: 'Full Period',
    building: 'OMS'
  }, extra || {});
}

/** The assignments the grid hands the client for one person, in one month. */
function bookingsFor(env, building, month, email) {
  const grid = env.run('getScheduleData', building);
  const entry = (grid[month] || []).find(r => r.email.toLowerCase() === email);
  if (!entry) throw new Error(email + ' is missing from ' + month + ' in ' + building);
  return entry.assignments || [];
}

/** Assigns as the OMS admin, in an environment the caller can keep reading. */
function assign(env, extra) {
  const res = env.run('assignCoverage', payload(extra));
  if (!res || !res.id) throw new Error('assignCoverage did not create a row: ' + JSON.stringify(res));
  return res.id;
}

exports.name = 'Master Schedule assigned coverage';

exports.run = function ({ test, assert }) {

  // ---- The marker, and the cell it belongs in --------------------------------

  test('an assignment appears on the schedule entry for the person covering', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env);

    const booked = bookingsFor(env, 'OMS', SOON.month, USERS.omsTeacher);
    assert.equal(booked.length, 1);
    assert.equal(booked[0].date, SOON.key);
    assert.equal(booked[0].coveredFor, 'Ted Teacher');
  });

  test('it carries the month, weekday and period of the cell it would collide in', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env);

    const b = bookingsFor(env, 'OMS', SOON.month, USERS.omsTeacher)[0];
    assert.equal(b.month, SOON.month);
    assert.equal(b.weekday, SOON.weekday);
    assert.equal(b.period, OMS_PERIOD, 'the raw period, so it matches the grid row exactly');
  });

  test('the date arrives ready to read, not as a storage key', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env);

    const b = bookingsFor(env, 'OMS', SOON.month, USERS.omsTeacher)[0];
    assert.equal(b.dateShort, SOON.short, 'the date that goes on the card');
    assert.equal(b.dateDisplay, SOON.weekday + ' ' + SOON.short);
    assert.equal(b.periodDisplay, 'Period 1 (8:10 AM – 8:57 AM)',
      'never a stored 24-hour time in front of a person');
  });

  test('coverage that has not happened yet reads as upcoming', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env);
    assert.equal(bookingsFor(env, 'OMS', SOON.month, USERS.omsTeacher)[0].upcoming, true);

    const past = envFor(USERS.omsAdmin);
    assign(past, { date: PAST.key });
    assert.equal(bookingsFor(past, 'OMS', PAST.month, USERS.omsTeacher)[0].upcoming, false);
  });

  test('a second booking on another date is listed alongside the first', () => {
    const env = envFor(USERS.omsAdmin);
    const other = coverageDate(14);
    assign(env);
    assign(env, { date: other.key });

    const booked = bookingsFor(env, 'OMS', SOON.month, USERS.omsTeacher);
    assert.equal(booked.length, 2);
    assert.deepEqual(booked.map(b => b.date), [SOON.key, other.key], 'oldest first');
  });

  // ---- What drops off it -----------------------------------------------------

  test('a cancelled assignment stops marking the person as taken', () => {
    const env = envFor(USERS.omsAdmin);
    const id = assign(env);
    assert.allowed(env.attempt('cancelAssignment', id));

    assert.equal(bookingsFor(env, 'OMS', SOON.month, USERS.omsTeacher).length, 0,
      'freeing the person up is the whole reason to cancel');
  });

  test('a recorded assignment stays, flagged as recorded', () => {
    const env = envFor(USERS.omsAdmin);
    const id = assign(env, { date: PAST.key });
    assert.allowed(env.attempt('recordAssignment', id, USERS.omsTeacher));

    const booked = bookingsFor(env, 'OMS', PAST.month, USERS.omsTeacher);
    assert.equal(booked.length, 1, 'the coverage still happened');
    assert.equal(booked[0].recorded, true);
  });

  // ---- Scope -----------------------------------------------------------------

  test("another building's assignments stay out of this building's grid", () => {
    const env = envFor(USERS.superAdmin);
    assert.allowed(env.attempt('assignCoverage', payload({
      teacherEmail: USERS.ohsTeacher, teacherName: 'Hank Teacher',
      subbedFor: 'Otto Admin', coveredForEmail: USERS.ohsAdmin,
      period: OHS_PERIOD, building: 'OHS'
    })));

    assert.equal(bookingsFor(env, 'OMS', SOON.month, USERS.omsTeacher).length, 0);
    assert.equal(bookingsFor(env, 'OHS', SOON.month, USERS.ohsTeacher).length, 1);
  });

  test('a teacher sees their own bookings and nobody else appears at all', () => {
    const admin = envFor(USERS.omsAdmin);
    assign(admin);
    assign(admin, {
      teacherEmail: USERS.omsTeacher2, teacherName: 'Ted Teacher',
      subbedFor: 'Tina Teacher', coveredForEmail: USERS.omsTeacher
    });

    // The same spreadsheet, read by the teacher.
    const teacher = envFor(USERS.omsTeacher, {
      assignments: admin.sheet('TST Assignments').values.map(r => r.slice())
    });

    const rows = (teacher.run('getScheduleData', 'OMS')[SOON.month] || []);
    assert.equal(rows.length, 1, 'a teacher only ever gets their own rows');
    assert.equal(rows[0].email.toLowerCase(), USERS.omsTeacher);
    assert.equal(rows[0].assignments.length, 1, "and only their own building's bookings");
    assert.equal(rows[0].assignments[0].coveredFor, 'Ted Teacher');
  });

  // ---- The calendar, where a building has one --------------------------------

  test('a building with no calendar simply has no calendar line', () => {
    const env = envFor(USERS.omsAdmin);
    assign(env);
    assert.equal(bookingsFor(env, 'OMS', SOON.month, USERS.omsTeacher)[0].calendar, '',
      'the marker still shows — the assignment is what makes someone taken');
  });

  test('a building with a calendar reports where the event stands', () => {
    const env = envFor(USERS.omsAdmin, {
      calendarId: OMS_CAL,
      calendars: { [OMS_CAL]: { name: 'OMS TST Calendar' } }
    });
    assign(env);

    assert.equal(bookingsFor(env, 'OMS', SOON.month, USERS.omsTeacher)[0].calendar,
      'Calendar event still queued', 'the admin trigger has not run yet');

    assert.allowed(env.attempt('processEmailQueue',
      { authMode: env.AuthMode.FULL, triggerUid: 'trigger-1' }));

    assert.equal(bookingsFor(env, 'OMS', SOON.month, USERS.omsTeacher)[0].calendar,
      'On the OMS TST Calendar');
  });
};
