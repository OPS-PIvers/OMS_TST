/**
 * The hours shown under each teacher in the Master Schedule.
 *
 * It is the approved TST they have earned so far this school year, combined
 * across their buildings — the same number the Directory's Earned column shows.
 * The grid sorts teachers by it so the admin is offered whoever has contributed
 * least. Three things follow that are easy to get wrong:
 *
 *   - A pending or denied request must not count. Denied especially — inflating
 *     someone's number because you turned their request down steers coverage away
 *     from exactly the person it should go to.
 *   - It is a year-to-date total, not a per-month one, so it reads the same in
 *     every month a person appears in.
 *   - It counts hours EARNED, not hours available. Spending TST, being paid out,
 *     or carrying hours over must not move it, or someone who covers constantly
 *     and spends it all would look like the least busy person in the building.
 */

const { createEnv } = require('./apps_script_env');
const {
  USERS, APPROVALS_HEADER, USAGE_HEADER, AVAILABILITY_HEADER, sheets
} = require('./fixtures');

// A Email | B Name | C SubbedFor | D | E Date | F Period | G TimeType | H Hours |
// I Approved | J | K Denied | L | M | N Building
const row = (email, name, date, hours, approved, denied, building) =>
  [email, name, 'Ted Teacher', '', date, 'Period 1 - 8:10 - 8:57', 'Full Period',
   hours, approved, '', denied, '', '', building || 'OMS'];

/** Everyone available in both months, so both render and can be compared. */
const AVAILABILITY = [
  AVAILABILITY_HEADER,
  ['September', 'Mon', 'Period 1 - 8:10 - 8:57', 'Tina Teacher', USERS.omsTeacher, ''],
  ['September', 'Mon', 'Period 1 - 8:10 - 8:57', 'Ted Teacher', USERS.omsTeacher2, ''],
  ['October', 'Mon', 'Period 1 - 8:10 - 8:57', 'Tina Teacher', USERS.omsTeacher, ''],
  ['October', 'Mon', 'Period 1 - 8:10 - 8:57', 'Ted Teacher', USERS.omsTeacher2, '']
];

/** An admin's view of the grid, with the given rows in place. */
function envWith(approvalRows, extra) {
  return createEnv({
    activeUser: USERS.omsAdmin,
    sheets: sheets(Object.assign({
      'TST Approvals (New)': [APPROVALS_HEADER, ...approvalRows],
      'TST Availability': AVAILABILITY
    }, extra || {}))
  });
}

const gridWith = (approvalRows, extra) =>
  envWith(approvalRows, extra).run('getScheduleData', 'OMS');

/** The hours the grid shows for one person in one month. */
function hoursFor(grid, month, email) {
  const cell = (grid[month] || []).find(r => r.email.toLowerCase() === email);
  if (!cell) throw new Error(email + ' is missing from ' + month);
  return cell.hours;
}

exports.name = 'Master Schedule earned hours';

exports.run = function ({ test, assert }) {

  // ---- Approved only ---------------------------------------------------------

  test('approved coverage counts', () => {
    const grid = gridWith([
      row(USERS.omsTeacher, 'Tina Teacher', '2025-09-10', 1, true, false),
      row(USERS.omsTeacher, 'Tina Teacher', '2025-09-12', 1.5, true, false)
    ]);
    assert.equal(hoursFor(grid, 'September', USERS.omsTeacher), 2.5);
  });

  test('a denied request does not inflate the number', () => {
    const grid = gridWith([
      row(USERS.omsTeacher, 'Tina Teacher', '2025-09-10', 1, true, false),
      row(USERS.omsTeacher, 'Tina Teacher', '2025-09-12', 4, false, true)
    ]);
    assert.equal(hoursFor(grid, 'September', USERS.omsTeacher), 1,
      'denying a request must not make someone look busier than they are');
  });

  test('a teacher whose only request was denied reads 0', () => {
    const grid = gridWith([row(USERS.omsTeacher, 'Tina Teacher', '2025-09-10', 4, false, true)]);
    assert.equal(hoursFor(grid, 'September', USERS.omsTeacher), 0);
  });

  test('a pending request is not counted, but still rides along as pending', () => {
    const grid = gridWith([row(USERS.omsTeacher, 'Tina Teacher', '2025-09-10', 2, false, false)]);
    const cell = grid['September'].find(r => r.email.toLowerCase() === USERS.omsTeacher);

    assert.equal(cell.hours, 0, 'unapproved hours are not earned hours');
    assert.equal(cell.pendingRequests.length, 1,
      'the hourglass is what surfaces pending coverage in this cell');
  });

  test('a checkbox read back as the string "TRUE" still counts', () => {
    const grid = gridWith([row(USERS.omsTeacher, 'Tina Teacher', '2025-09-10', 3, 'TRUE', '')]);
    assert.equal(hoursFor(grid, 'September', USERS.omsTeacher), 3);
  });

  // ---- A year-to-date total, not a per-month one ------------------------------

  test('the same total shows in every month the person appears in', () => {
    const grid = gridWith([row(USERS.omsTeacher, 'Tina Teacher', '2025-09-10', 2, true, false)]);
    assert.equal(hoursFor(grid, 'September', USERS.omsTeacher), 2);
    assert.equal(hoursFor(grid, 'October', USERS.omsTeacher), 2,
      'the months organise availability, not the hours');
  });

  test("September's cell includes hours earned in October", () => {
    const grid = gridWith([
      row(USERS.omsTeacher, 'Tina Teacher', '2025-09-10', 2, true, false),
      row(USERS.omsTeacher, 'Tina Teacher', '2025-10-10', 5, true, false)
    ]);
    assert.equal(hoursFor(grid, 'September', USERS.omsTeacher), 7);
  });

  test('an old date still counts until finalize clears it', () => {
    // Deliberate: the total is "earned since the last year-end roll", exactly like
    // the Directory's Earned column, rather than being filtered by a date window.
    // finalizeSchoolYear deletes approved rows from the sheet, and that is what
    // resets this number. The per-month version this replaced filtered by date.
    const grid = gridWith([row(USERS.omsTeacher, 'Tina Teacher', '2019-09-10', 3, true, false)]);
    assert.equal(hoursFor(grid, 'September', USERS.omsTeacher), 3);
  });

  // ---- Earned, not available --------------------------------------------------

  test('using TST does not reduce the number', () => {
    // The distinction that makes the sort work: someone who covers constantly and
    // spends it all must not read as the least busy person in the building.
    const grid = gridWith(
      [row(USERS.omsTeacher, 'Tina Teacher', '2025-09-10', 6, true, false)],
      { 'TST Usage (New)': [USAGE_HEADER,
          [USERS.omsTeacher, 'Tina Teacher', '2025-09-20', 5, true, '2025-09-20', '', 'OMS']] }
    );
    assert.equal(hoursFor(grid, 'September', USERS.omsTeacher), 6,
      'hours earned, not hours left');
  });

  test('carry over and paid out do not move the number', () => {
    // Tina carries 3 and has 1 paid out in the base fixture; neither is coverage
    // she provided, so neither belongs in a number about who has been covering.
    const grid = gridWith([row(USERS.omsTeacher, 'Tina Teacher', '2025-09-10', 2, true, false)]);
    assert.equal(hoursFor(grid, 'September', USERS.omsTeacher), 2);
  });

  // ---- Agreement with the rest of the app -------------------------------------

  test("the grid's number is the Directory's Earned column", () => {
    const env = envWith([
      row(USERS.omsTeacher, 'Tina Teacher', '2025-09-10', 1, true, false),
      row(USERS.omsTeacher, 'Tina Teacher', '2025-10-10', 2.5, true, false),
      row(USERS.omsTeacher, 'Tina Teacher', '2025-10-12', 4, false, true)
    ]);
    const tina = env.run('getStaffDirectoryData', 'OMS')
      .find(s => s.email.toLowerCase() === USERS.omsTeacher);
    const grid = env.run('getScheduleData', 'OMS');

    assert.equal(hoursFor(grid, 'September', USERS.omsTeacher), tina.earned,
      'a teacher must not read as busier on one screen than the other');
    assert.equal(tina.earned, 3.5);
  });

  test("a multi-building teacher's total includes both buildings", () => {
    const env = createEnv({
      activeUser: USERS.omsAdmin,
      sheets: sheets({
        'TST Approvals (New)': [APPROVALS_HEADER,
          row(USERS.multiTeacher, 'Mia Multi', '2025-09-10', 1, true, false, 'OMS'),
          row(USERS.multiTeacher, 'Mia Multi', '2025-09-11', 1, true, false, 'OHS')],
        'TST Availability': [AVAILABILITY_HEADER,
          ['September', 'Mon', 'Period 1 - 8:10 - 8:57', 'Mia Multi', USERS.multiTeacher, '']]
      })
    });
    assert.equal(hoursFor(env.run('getScheduleData', 'OMS'), 'September', USERS.multiTeacher), 2,
      'earned totals are combined across buildings everywhere else too');
  });

  // ---- The number is always renderable ----------------------------------------

  test('an unreadable hours cell does not swallow the rest of the total', () => {
    // The grid renders Number(hours).toFixed(1); one NaN in the running sum would
    // take the whole total with it.
    const grid = gridWith([
      row(USERS.omsTeacher, 'Tina Teacher', '2025-09-10', 2, true, false),
      row(USERS.omsTeacher, 'Tina Teacher', '2025-09-12', 'half', true, false)
    ]);
    assert.equal(hoursFor(grid, 'September', USERS.omsTeacher), 2);
  });

  test('everyone in the grid has a numeric hours value', () => {
    const grid = gridWith([row(USERS.omsTeacher, 'Tina Teacher', '2025-09-10', 2, true, false)]);
    Object.keys(grid).forEach(month => {
      grid[month].forEach(r => assert.ok(Number.isFinite(r.hours),
        r.email + ' in ' + month + ' has a non-numeric hours value: ' + r.hours));
    });
  });

  test('a teacher with no approved coverage reads 0, not blank', () => {
    const grid = gridWith([row(USERS.omsTeacher, 'Tina Teacher', '2025-09-10', 2, true, false)]);
    assert.equal(hoursFor(grid, 'September', USERS.omsTeacher2), 0);
  });
};
