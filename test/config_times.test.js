/**
 * Building configuration: who may change it, and how period times resolve.
 *
 * Calendar events need a real start and end, and only OMS ever had times — baked
 * into the period label ("Period 8 - 12:37 - 1:08"). OHS runs two bell schedules
 * (MWF and TTh), so a period label alone cannot decide a time; the date does.
 */

const { createEnv } = require('./apps_script_env');
const { USERS, sheets } = require('./fixtures');

const envFor = (email, overrides) => createEnv({ activeUser: email, sheets: sheets(overrides) });

/** A yyyy-MM-dd string for the next date falling on `weekday` (0 = Sunday). */
function nextWeekday(weekday) {
  const d = new Date();
  while (d.getDay() !== weekday) d.setDate(d.getDate() + 1);
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

exports.name = 'Building config and period times';

exports.run = function ({ test, assert }) {

  // ---- Who may edit a building's settings ------------------------------------

  test('a teacher cannot save building config', () => {
    const env = envFor(USERS.omsTeacher);
    assert.rejected(env.attempt('saveBuildingConfig', 'OMS', { name: 'Hacked' }),
      /admin access required/i);
  });

  test("an admin cannot rewrite another building's settings", () => {
    const env = envFor(USERS.omsAdmin);
    // Settings is a building-picker form, so nothing but the server stops an admin
    // from selecting a school they do not run and saving over its periods.
    assert.rejected(env.attempt('saveBuildingConfig', 'OHS', { name: 'Renamed by OMS' }),
      /your own building/i);
    assert.equal(env.run('getConfig').OHS.name, 'Orono High School', 'OHS is untouched');
  });

  test('an admin can save their own building', () => {
    const env = envFor(USERS.omsAdmin);
    assert.allowed(env.attempt('saveBuildingConfig', 'OMS',
      { name: 'Orono Middle School', scheduleType: 'periods', periods: ['Period 1'] }));
    assert.deepEqual(env.run('getConfig').OMS.periods, ['Period 1']);
  });

  test('a multi-building admin can save either of theirs', () => {
    const env = envFor(USERS.dualAdmin);
    assert.allowed(env.attempt('saveBuildingConfig', 'OMS', { name: 'OMS' }));
    assert.allowed(env.attempt('saveBuildingConfig', 'OHS', { name: 'OHS' }));
  });

  test('a Super Admin can save any building', () => {
    const env = envFor(USERS.superAdmin);
    assert.allowed(env.attempt('saveBuildingConfig', 'OHS', { name: 'Orono High School' }));
    assert.allowed(env.attempt('saveBuildingConfig', 'SE', { name: 'Schumann Elementary School' }));
  });

  test('a non-Super-Admin still cannot move the carry-over cap', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('saveBuildingConfig', 'OMS', { name: 'OMS', carryOverMax: 999 });
    assert.equal(env.run('getConfig').OMS.carryOverMax, 12, 'the stored cap is preserved');
  });

  // ---- Reading times out of a period label -----------------------------------

  test('an OMS period label carries its own times', () => {
    const env = envFor(USERS.omsAdmin);
    assert.deepEqual(env.callInternal('parseTimeRange_', 'Period 1 - 8:10 - 8:57'),
      { start: '08:10', end: '08:57' });
  });

  test('an afternoon period does not become a twelve-hour event', () => {
    const env = envFor(USERS.omsAdmin);
    // "12:37 - 1:08" is 12:37pm to 1:08pm. Read literally it runs backwards.
    assert.deepEqual(env.callInternal('parseTimeRange_', 'Period 8 - 12:37 - 1:08'),
      { start: '12:37', end: '13:08' });
    assert.deepEqual(env.callInternal('parseTimeRange_', 'Period 10 - 2:03 - 2:50'),
      { start: '14:03', end: '14:50' });
  });

  test('a time-range building\'s period is already the span', () => {
    const env = envFor(USERS.omsAdmin);
    // OIS and SE store "08:30 - 09:15" as the period itself, in 24-hour form.
    assert.deepEqual(env.callInternal('parseTimeRange_', '08:30 - 09:15'),
      { start: '08:30', end: '09:15' });
  });

  test('a label with no times reads as none, not as a guess', () => {
    const env = envFor(USERS.omsAdmin);
    assert.equal(env.callInternal('parseTimeRange_', 'Period 1'), null);
    assert.equal(env.callInternal('parseTimeRange_', 'Time Range'), null);
    assert.equal(env.callInternal('parseTimeRange_', ''), null);
  });

  // ---- periodTimesFor_ -------------------------------------------------------

  test('OMS needs no configuration — its labels already say the times', () => {
    const env = envFor(USERS.omsAdmin);
    const t = env.callInternal('periodTimesFor_', 'OMS', 'Period 3 - 9:52 - 10:39', nextWeekday(1));
    assert.equal(t.start, '09:52');
    assert.equal(t.end, '10:39');
    assert.equal(t.source, 'label');
  });

  test('OHS reads its bell schedule from config', () => {
    const env = envFor(USERS.ohsAdmin);
    const monday = env.callInternal('periodTimesFor_', 'OHS', 'Period 1', nextWeekday(1));
    assert.equal(monday.start, '08:00');
    assert.equal(monday.end, '08:48', 'Mon/Wed/Fri is the default');

    const tuesday = env.callInternal('periodTimesFor_', 'OHS', 'Period 1', nextWeekday(2));
    assert.equal(tuesday.end, '08:41', 'Tue/Thu is shorter, to make room for Spartan Hour');
    assert.equal(tuesday.source, 'TTh');
  });

  test('a period that only runs some days has no time on the others', () => {
    const env = envFor(USERS.ohsAdmin);
    // Spartan Hour is Tue/Thu only, so it has no Mon/Wed/Fri default. This is the
    // case that must report rather than guess: inventing one would put a teacher
    // on a calendar at a time that does not exist.
    assert.equal(env.callInternal('periodTimesFor_', 'OHS', 'Spartan Hour', nextWeekday(1)), null);

    const thursday = env.callInternal('periodTimesFor_', 'OHS', 'Spartan Hour', nextWeekday(4));
    assert.equal(thursday.start, '08:45');
    assert.equal(thursday.end, '09:25');
  });

  test('an unknown period is reported, never invented', () => {
    const env = envFor(USERS.ohsAdmin);
    assert.equal(env.callInternal('periodTimesFor_', 'OHS', 'Period 9', nextWeekday(1)), null);
  });

  test('configured default times are used once set', () => {
    const env = envFor(USERS.ohsAdmin);
    env.run('saveBuildingConfig', 'OHS', {
      name: 'Orono High School',
      scheduleType: 'periods',
      periods: ['Period 1'],
      periodTimes: { 'Period 1': { start: '08:00', end: '08:47' } }
    });

    const t = env.callInternal('periodTimesFor_', 'OHS', 'Period 1', nextWeekday(1));
    assert.equal(t.start, '08:00');
    assert.equal(t.end, '08:47');
    assert.equal(t.source, 'default');
  });

  test('a day schedule overrides the default on its own days only', () => {
    const env = envFor(USERS.ohsAdmin);
    env.run('saveBuildingConfig', 'OHS', {
      name: 'Orono High School',
      scheduleType: 'periods',
      periods: ['Period 1'],
      periodTimes: { 'Period 1': { start: '08:00', end: '08:47' } },
      dayGroups: [{
        name: 'TTh',
        days: ['Tue', 'Thu'],
        times: { 'Period 1': { start: '08:00', end: '09:30' } }
      }]
    });

    const tuesday = env.callInternal('periodTimesFor_', 'OHS', 'Period 1', nextWeekday(2));
    assert.equal(tuesday.end, '09:30', 'Tuesday runs the block schedule');
    assert.equal(tuesday.source, 'TTh');

    const monday = env.callInternal('periodTimesFor_', 'OHS', 'Period 1', nextWeekday(1));
    assert.equal(monday.end, '08:47', 'Monday is not in the group, so it keeps the default');
    assert.equal(monday.source, 'default');

    const friday = env.callInternal('periodTimesFor_', 'OHS', 'Period 1', nextWeekday(5));
    assert.equal(friday.end, '08:47');
  });

  test('a day schedule that omits a period falls through to the default', () => {
    const env = envFor(USERS.ohsAdmin);
    env.run('saveBuildingConfig', 'OHS', {
      name: 'Orono High School',
      scheduleType: 'periods',
      periods: ['Period 1', 'Period 2'],
      periodTimes: {
        'Period 1': { start: '08:00', end: '08:47' },
        'Period 2': { start: '08:51', end: '09:38' }
      },
      dayGroups: [{ name: 'TTh', days: ['Tue', 'Thu'], times: { 'Period 1': { start: '08:00', end: '09:30' } } }]
    });

    const p2 = env.callInternal('periodTimesFor_', 'OHS', 'Period 2', nextWeekday(2));
    assert.equal(p2.end, '09:38', 'only Period 1 was overridden');
    assert.equal(p2.source, 'default');
  });

  test('configured times win over times written into the label', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('saveBuildingConfig', 'OMS', {
      name: 'Orono Middle School',
      scheduleType: 'periods',
      periods: ['Period 1 - 8:10 - 8:57'],
      periodTimes: { 'Period 1 - 8:10 - 8:57': { start: '08:15', end: '09:00' } }
    });

    const t = env.callInternal('periodTimesFor_', 'OMS', 'Period 1 - 8:10 - 8:57', nextWeekday(1));
    assert.equal(t.start, '08:15', 'the label is the fallback, not the authority');
  });

  // ---- How a period reads to a person ---------------------------------------

  test('times are shown in 12-hour form, never 24-hour', () => {
    const env = envFor(USERS.ohsAdmin);
    assert.equal(env.callInternal('formatTime12_', '13:52'), '1:52 PM');
    assert.equal(env.callInternal('formatTime12_', '08:00'), '8:00 AM');
    assert.equal(env.callInternal('formatTime12_', '12:05'), '12:05 PM', 'noon is 12 PM, not 0');
    assert.equal(env.callInternal('formatTime12_', '00:30'), '12:30 AM', 'and midnight is 12 AM');
  });

  test('a bare OHS period gains the time it runs that day', () => {
    const env = envFor(USERS.ohsAdmin);
    assert.equal(env.callInternal('periodDisplay_', 'OHS', 'Period 3', nextWeekday(1)),
      'Period 3 (9:54 AM – 10:42 AM)');
    assert.equal(env.callInternal('periodDisplay_', 'OHS', 'Period 3', nextWeekday(2)),
      'Period 3 (10:24 AM – 11:04 AM)', 'the date decides, because OHS runs two schedules');
  });

  test('an OMS label is not made to repeat its own times', () => {
    const env = envFor(USERS.omsAdmin);
    assert.equal(env.callInternal('periodDisplay_', 'OMS', 'Period 8 - 12:37 - 1:08', nextWeekday(1)),
      'Period 8 (12:37 PM – 1:08 PM)');
  });

  test('a time-range period is shown once, in 12-hour form', () => {
    const env = envFor(USERS.omsAdmin);
    assert.equal(env.callInternal('periodDisplay_', 'OIS', '08:30 - 09:15', nextWeekday(1)),
      '8:30 AM – 9:15 AM');
  });

  test('a period with no known time still reads as itself', () => {
    const env = envFor(USERS.ohsAdmin);
    assert.equal(env.callInternal('periodDisplay_', 'OHS', 'Spartan Hour', nextWeekday(1)),
      'Spartan Hour');
  });

  // ---- installBellSchedule ---------------------------------------------------

  test('installBellSchedule copies config.js into a live App Config sheet', () => {
    const env = envFor(USERS.superAdmin);
    // Simulate a district already running with the old four bare periods.
    env.run('saveBuildingConfig', 'OHS', {
      name: 'Orono High School', scheduleType: 'periods', carryOverMax: 12,
      periods: ['Period 1', 'Period 2', 'Period 3', 'Period 4'],
      calendarId: 'ohs@group.calendar.google.com'
    });

    env.run('installBellSchedule', 'OHS');

    const ohs = env.run('getConfig').OHS;
    assert.equal(ohs.periods.length, 10);
    assert.ok(ohs.periods.indexOf('Spartan Hour') > -1);
    assert.equal(ohs.periods.indexOf('Break'), -1,
      'the bell chart has a 10-minute Break; it is not something anyone covers');
    assert.equal(ohs.dayGroups.length, 1);
    assert.equal(ohs.periodTimes['Period 7'].end, '14:40');
    assert.equal(ohs.calendarId, 'ohs@group.calendar.google.com',
      'what the building set for itself must survive');
  });

  test('installBellSchedule is scoped like every other config write', () => {
    const env = envFor(USERS.omsAdmin);
    assert.rejected(env.attempt('installBellSchedule', 'OHS'), /your own building/i);
    assert.rejected(envFor(USERS.omsTeacher).attempt('installBellSchedule', 'OMS'),
      /admin access required/i);
  });

  // ---- Private helpers stay off the client surface ---------------------------

  ['parseTimeRange_', 'periodTimesFor_', 'formatTime12_', 'periodDisplay_'].forEach(fn => {
    test(`${fn} is not reachable from google.script.run`, () => {
      const env = envFor(USERS.omsTeacher);
      assert.rejected(env.attempt(fn), /private/i);
      assert.equal(typeof env.context[fn], 'function', 'but server code can still call it');
    });
  });
};
