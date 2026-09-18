/**
 * The hours shown under each teacher in the Master Schedule.
 *
 * It is a per-month count of APPROVED coverage, not a balance, and the grid sorts
 * teachers by it so the admin is offered whoever has picked up least this month.
 * Two things follow that are easy to get wrong:
 *
 *   - A pending or denied request must not count. Denied especially — inflating
 *     someone's number because you turned their request down steers coverage away
 *     from exactly the person it should go to.
 *   - A month with no coverage in it reads 0.0 for everyone, which is correct and
 *     is why every future month looks empty.
 *
 * Dates are built from the current school year rather than hardcoded, because
 * calculateMonthlyHours_ reads the real clock to decide which year it is in.
 */

const { createEnv } = require('./apps_script_env');
const { USERS, APPROVALS_HEADER, AVAILABILITY_HEADER, sheets } = require('./fixtures');

/** Aug–Jul, the same window calculateMonthlyHours_ uses. */
function schoolYearStartYear() {
  const today = new Date();
  return today.getMonth() >= 7 ? today.getFullYear() : today.getFullYear() - 1;
}

const START_YEAR = schoolYearStartYear();
const pad = n => String(n).padStart(2, '0');

/** A yyyy-MM-dd date inside the current school year. month is 1-12. */
const inYear = (month, day) =>
  `${month >= 8 ? START_YEAR : START_YEAR + 1}-${pad(month)}-${pad(day)}`;

// A Email | B Name | C SubbedFor | D | E Date | F Period | G TimeType | H Hours |
// I Approved | J | K Denied | L | M | N Building
const row = (email, name, date, hours, approved, denied) =>
  [email, name, 'Ted Teacher', '', date, 'Period 1 - 8:10 - 8:57', 'Full Period',
   hours, approved, '', denied, '', '', 'OMS'];

const SEPT = inYear(9, 10);
const OCT = inYear(10, 10);

/** Everyone available in September and October, so both months render. */
const AVAILABILITY = [
  AVAILABILITY_HEADER,
  ['September', 'Mon', 'Period 1 - 8:10 - 8:57', 'Tina Teacher', USERS.omsTeacher, ''],
  ['September', 'Mon', 'Period 1 - 8:10 - 8:57', 'Ted Teacher', USERS.omsTeacher2, ''],
  ['October', 'Mon', 'Period 1 - 8:10 - 8:57', 'Tina Teacher', USERS.omsTeacher, ''],
  ['October', 'Mon', 'Period 1 - 8:10 - 8:57', 'Ted Teacher', USERS.omsTeacher2, '']
];

/** The admin's view of the grid, with the given approval rows in place. */
function gridWith(approvalRows) {
  const env = createEnv({
    activeUser: USERS.omsAdmin,
    sheets: sheets({
      'TST Approvals (New)': [APPROVALS_HEADER, ...approvalRows],
      'TST Availability': AVAILABILITY
    })
  });
  return env.run('getScheduleData', 'OMS');
}

/** The hours the grid shows for one person in one month. */
function hoursFor(grid, month, email) {
  const cell = (grid[month] || []).find(r => r.email.toLowerCase() === email);
  if (!cell) throw new Error(`${email} is missing from ${month}`);
  return cell.hours;
}

exports.name = 'Master Schedule monthly hours';

exports.run = function ({ test, assert }) {

  test('approved coverage counts', () => {
    const grid = gridWith([
      row(USERS.omsTeacher, 'Tina Teacher', SEPT, 1, true, false),
      row(USERS.omsTeacher, 'Tina Teacher', SEPT, 1.5, true, false)
    ]);
    assert.equal(hoursFor(grid, 'September', USERS.omsTeacher), 2.5);
  });

  test('a denied request does not inflate the number', () => {
    const grid = gridWith([
      row(USERS.omsTeacher, 'Tina Teacher', SEPT, 1, true, false),
      row(USERS.omsTeacher, 'Tina Teacher', SEPT, 4, false, true)
    ]);
    assert.equal(hoursFor(grid, 'September', USERS.omsTeacher), 1,
      'denying a request must not make someone look busier than they are');
  });

  test('a teacher whose only request was denied reads 0', () => {
    const grid = gridWith([row(USERS.omsTeacher, 'Tina Teacher', SEPT, 4, false, true)]);
    assert.equal(hoursFor(grid, 'September', USERS.omsTeacher), 0);
  });

  test('a pending request is not counted, but still rides along as pending', () => {
    const grid = gridWith([row(USERS.omsTeacher, 'Tina Teacher', SEPT, 2, false, false)]);
    const cell = (grid['September'] || []).find(r => r.email.toLowerCase() === USERS.omsTeacher);

    assert.equal(cell.hours, 0, 'unapproved hours are not earned hours');
    assert.equal(cell.pendingRequests.length, 1,
      'the hourglass is what surfaces pending coverage in this cell');
  });

  test('a checkbox read back as the string "TRUE" still counts', () => {
    const grid = gridWith([row(USERS.omsTeacher, 'Tina Teacher', SEPT, 3, 'TRUE', '')]);
    assert.equal(hoursFor(grid, 'September', USERS.omsTeacher), 3);
  });

  // ---- Per month, not a running balance --------------------------------------

  test('a month with no coverage reads 0 for everyone', () => {
    const grid = gridWith([row(USERS.omsTeacher, 'Tina Teacher', SEPT, 2, true, false)]);
    (grid['October'] || []).forEach(r =>
      assert.equal(r.hours, 0, `${r.email} should have no October hours`));
  });

  test("September's number does not include October's hours", () => {
    const grid = gridWith([
      row(USERS.omsTeacher, 'Tina Teacher', SEPT, 2, true, false),
      row(USERS.omsTeacher, 'Tina Teacher', OCT, 5, true, false)
    ]);
    assert.equal(hoursFor(grid, 'September', USERS.omsTeacher), 2);
    assert.equal(hoursFor(grid, 'October', USERS.omsTeacher), 5);
  });

  test('last school year does not leak into this one', () => {
    const grid = gridWith([
      row(USERS.omsTeacher, 'Tina Teacher', `${START_YEAR - 1}-09-10`, 6, true, false)
    ]);
    assert.equal(hoursFor(grid, 'September', USERS.omsTeacher), 0);
  });

  // ---- The number is always renderable ---------------------------------------

  test('an unreadable hours cell does not swallow the rest of the month', () => {
    // NaN poisons the running sum, and scheduleData_ then reads NaN as falsy and
    // hands the grid a 0 — so one bad cell silently erases real approved hours
    // rather than showing anything obviously wrong. Order matters: the bad row has
    // to land after a good one for the sum to carry the NaN out.
    const grid = gridWith([
      row(USERS.omsTeacher, 'Tina Teacher', SEPT, 2, true, false),
      row(USERS.omsTeacher, 'Tina Teacher', SEPT, 'half', true, false)
    ]);
    assert.equal(hoursFor(grid, 'September', USERS.omsTeacher), 2);
  });

  // Weaker than the one above on purpose: scheduleData_'s own `|| 0` means this
  // holds either way. It is here to catch a future change that lets NaN through.
  test('everyone in the grid has a numeric hours value', () => {
    const grid = gridWith([row(USERS.omsTeacher, 'Tina Teacher', SEPT, 2, true, false)]);
    Object.keys(grid).forEach(month => {
      grid[month].forEach(r => assert.ok(Number.isFinite(r.hours),
        `${r.email} in ${month} has a non-numeric hours value: ${r.hours}`));
    });
  });

  // ---- Hours are combined across buildings, like every other earned total -----

  test("a multi-building teacher's month includes both buildings", () => {
    const env = createEnv({
      activeUser: USERS.omsAdmin,
      sheets: sheets({
        'TST Approvals (New)': [APPROVALS_HEADER,
          [USERS.multiTeacher, 'Mia Multi', 'Tina Teacher', '', SEPT, 'Period 1 - 8:10 - 8:57',
           'Full Period', 1, true, '', false, '', '', 'OMS'],
          [USERS.multiTeacher, 'Mia Multi', 'Hank Teacher', '', SEPT, 'Period 2',
           'Full Period', 1, true, '', false, '', '', 'OHS']],
        'TST Availability': [AVAILABILITY_HEADER,
          ['September', 'Mon', 'Period 1 - 8:10 - 8:57', 'Mia Multi', USERS.multiTeacher, '']]
      })
    });
    assert.equal(hoursFor(env.run('getScheduleData', 'OMS'), 'September', USERS.multiTeacher), 2);
  });
};
