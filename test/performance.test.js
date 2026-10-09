/**
 * Read cost of the busiest screens.
 *
 * Formatting a period or naming a calendar asks for the App Config, and the
 * Assignments queue, the tab badges and the Master Schedule do that once per
 * assignment. Each of those used to re-read the App Config sheet, so every
 * assignment ever made slowed those screens a little more. The config is now read
 * once per execution — these tests hold that, and that the cache never serves
 * a stale or shared copy.
 *
 * The Master Schedule also used to send each person's whole list of pending
 * requests and assignments on every one of their availability rows; each row now
 * carries only its own month's, which is all the grid draws in it.
 */

const { createEnv, FakeSheet } = require('./apps_script_env');
const { USERS, AVAILABILITY_HEADER, sheets } = require('./fixtures');

const OMS_PERIOD = 'Period 1 - 8:10 - 8:57';
const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
  'August', 'September', 'October', 'November', 'December'];

function appConfig(calendarName) {
  const OMS = {
    name: 'Orono Middle School', scheduleType: 'periods', periods: [OMS_PERIOD],
    coverageTypes: [{ label: 'Full Period', value: 1 }], calendarName: calendarName || ''
  };
  const OHS = { name: 'Orono High School', scheduleType: 'periods', periods: ['Period 1'],
    coverageTypes: [{ label: 'Full Period', value: 1 }] };
  return [['Building', 'Config_JSON'], ['OMS', JSON.stringify(OMS)], ['OHS', JSON.stringify(OHS)]];
}

function dateKey(offset) {
  const d = new Date();
  d.setDate(d.getDate() + offset);
  const pad = n => String(n).padStart(2, '0');
  return { key: `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`, month: MONTH_NAMES[d.getMonth()] };
}

function envWithAssignments(count, availability) {
  const env = createEnv({
    activeUser: USERS.omsAdmin,
    sheets: sheets({ 'App Config': appConfig(), 'TST Availability': availability || [AVAILABILITY_HEADER] })
  });
  for (let i = 0; i < count; i++) {
    env.run('assignCoverage', {
      teacherEmail: USERS.omsTeacher, teacherName: 'Tina Teacher', subbedFor: 'Ted Teacher',
      coveredForEmail: USERS.omsTeacher2, date: dateKey(i * 7 - 70).key, period: OMS_PERIOD,
      amount: 1, amountType: 'Full Period', building: 'OMS', force: true
    });
  }
  return env;
}

/** How many times each sheet is read in full during one call. */
function readsDuring(env, fn) {
  const counts = {};
  const original = FakeSheet.prototype.getDataRange;
  FakeSheet.prototype.getDataRange = function () {
    counts[this.name] = (counts[this.name] || 0) + 1;
    return original.call(this);
  };
  try { fn(); } finally { FakeSheet.prototype.getDataRange = original; }
  return counts;
}

exports.name = 'Read cost';

exports.run = function ({ test, assert }) {

  ['getAssignments', 'getDashboardCounts', 'getScheduleData'].forEach(fn => {
    test(`${fn} reads App Config once, however many assignments there are`, () => {
      const env = envWithAssignments(20);
      const reads = readsDuring(env, () => env.run(fn, 'OMS'));
      assert.equal(reads['App Config'], 1, JSON.stringify(reads));
    });
  });

  test('getScheduleData reads Approvals twice and Usage once (balances worked out once)', () => {
    const env = envWithAssignments(3);
    const reads = readsDuring(env, () => env.run('getScheduleData', 'OMS'));
    assert.equal(reads['TST Approvals (New)'], 2, JSON.stringify(reads));
    assert.equal(reads['TST Usage (New)'], 1, JSON.stringify(reads));
  });

  // ---- The cache never serves stale or shared config ------------------------

  test('a config saved earlier in the same execution is what the rest of it reads', () => {
    const env = envWithAssignments(0);
    const ctx = env.context;
    assert.equal(ctx.getConfig().OMS.calendarName, '');
    const updated = Object.assign({}, ctx.getConfig().OMS, { calendarName: 'Renamed' });
    ctx.saveBuildingConfig('OMS', updated);   // no reset between: one execution
    assert.equal(ctx.getConfig().OMS.calendarName, 'Renamed');
  });

  test('a change to the sheet between requests is picked up by the next request', () => {
    const env = envWithAssignments(0);
    assert.equal(env.run('getInitialData').config.OMS.calendarName, '');
    env.sheet('App Config').values = appConfig('Edited in the sheet');
    assert.equal(env.run('getInitialData').config.OMS.calendarName, 'Edited in the sheet');
  });

  test('changing what getConfig returned does not change the next getConfig', () => {
    const env = envWithAssignments(0);
    const ctx = env.context;
    ctx.getConfig().OMS.periods.push('Bogus');
    ctx.getConfig().OMS.name = 'Bogus';
    assert.deepEqual(ctx.getConfig().OMS.periods, [OMS_PERIOD]);
    assert.equal(ctx.getConfig().OMS.name, 'Orono Middle School');
  });

  // ---- Schedule rows carry only their own month -----------------------------

  test('each schedule row carries only the assignments and pending requests of its month', () => {
    const avail = [AVAILABILITY_HEADER];
    ['September', 'October', 'November', 'December', 'January', 'February', 'March', 'April', 'May', 'June']
      .forEach(m => avail.push([m, 'Mon,Tue,Wed,Thu,Fri', OMS_PERIOD, 'Tina Teacher', USERS.omsTeacher, '']));
    const env = envWithAssignments(15, avail);
    // A pending request whose date can't be read has no month: it must stay on every row.
    env.sheet('TST Approvals (New)').values.push([USERS.omsTeacher, 'Tina Teacher', 'Ted Teacher', '',
      'not a date', OMS_PERIOD, 'Full Period', 1, false, '', false, '', '', 'OMS']);

    const grid = env.run('getScheduleData', 'OMS');
    let seen = 0;
    Object.keys(grid).forEach(month => {
      grid[month].filter(e => e.email === USERS.omsTeacher).forEach(e => {
        e.assignments.forEach(a => assert.equal(a.month, month, 'assignment from another month on ' + month));
        e.pendingRequests.forEach(r => assert.ok(!r.month || r.month === month, 'request from another month on ' + month));
        assert.ok(e.pendingRequests.some(r => !r.month), 'the undated request is missing from ' + month);
        seen += e.assignments.length;
      });
    });
    // Every assignment still to come that falls in a school month appears exactly once.
    const today = env.context.normDateKey_(new Date());
    const all = env.run('getAssignments', 'OMS').filter(a => a.date >= today &&
      grid[MONTH_NAMES[Number(a.date.slice(5, 7)) - 1]]);
    assert.equal(seen, all.length);
    assert.ok(seen > 0, 'no assignments landed in the school year');
  });
};
