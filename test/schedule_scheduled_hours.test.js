/**
 * "(+N scheduled)" beside the hours on the Master Schedule.
 *
 * The hours number is approved TST only — it must keep equalling the Directory's
 * Earned column. A building admin asked for coverage already assigned to be
 * counted too, so the person booked for next week is not offered first just
 * because those hours have not been approved yet. So it is shown separately, and
 * the grid sorts on the two together.
 *
 * The rule that keeps it from counting anything twice: an assignment counts until
 * its hours are approved — then they are in `hours` instead.
 */

const { createEnv } = require('./apps_script_env');
const { USERS, APPROVALS_HEADER, AVAILABILITY_HEADER, sheets } = require('./fixtures');

const P1 = 'Period 1 - 8:10 - 8:57';
const AVAILABILITY = [
  AVAILABILITY_HEADER,
  ['October', 'Mon', P1, 'Tina Teacher', USERS.omsTeacher, ''],
  ['October', 'Mon', P1, 'Ted Teacher', USERS.omsTeacher2, ''],
  ['October', 'Mon', P1, 'Mia Multi', USERS.multiTeacher, '']
];
const OHS_CONFIG = { name: 'Orono High School', scheduleType: 'periods', periods: ['Period 1'],
  coverageTypes: [{ label: 'Full Period', value: 1 }] };
const OMS_CONFIG = { name: 'Orono Middle School', scheduleType: 'periods', periods: [P1],
  coverageTypes: [{ label: 'Full Period', value: 1 }, { label: 'Half Period', value: 0.5 }] };

function day(offset) {
  const d = new Date();
  d.setDate(d.getDate() + offset);
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

function envFor(user) {
  return createEnv({
    activeUser: user || USERS.omsAdmin,
    sheets: sheets({
      'TST Approvals (New)': [APPROVALS_HEADER],
      'TST Availability': AVAILABILITY,
      'App Config': [['Building', 'Config_JSON'], ['OMS', JSON.stringify(OMS_CONFIG)], ['OHS', JSON.stringify(OHS_CONFIG)]]
    })
  });
}

function assign(env, extra) {
  const res = env.run('assignCoverage', Object.assign({
    teacherEmail: USERS.omsTeacher, teacherName: 'Tina Teacher', subbedFor: 'Ted Teacher',
    coveredForEmail: USERS.omsTeacher2, date: day(7), period: P1, amount: 1,
    amountType: 'Full Period', building: 'OMS', force: true
  }, extra || {}));
  if (!res || !res.id) throw new Error('assignCoverage failed: ' + JSON.stringify(res));
  return res.id;
}

const entry = (env, email) => {
  const row = env.run('getScheduleData', 'OMS').October.find(r => r.email === email);
  if (!row) throw new Error(email + ' is not on the October grid');
  return row;
};
const approvalRowFor = (env, email) => {
  const values = env.sheet('TST Approvals (New)').values;
  for (let i = 1; i < values.length; i++) if (values[i][0] === email) return i + 1;
  return -1;
};

exports.name = 'Master Schedule scheduled hours';

exports.run = function ({ test, assert }) {

  test('an upcoming assignment counts as scheduled and leaves the earned hours alone', () => {
    const env = envFor();
    assign(env);
    const e = entry(env, USERS.omsTeacher);
    assert.equal(e.scheduledHours, 1);
    assert.equal(e.hours, 0, 'hours is still approved hours only');
  });

  test('half-period and several assignments add up', () => {
    const env = envFor();
    assign(env);
    assign(env, { date: day(8), amount: 0.5, amountType: 'Half Period' });
    assert.equal(entry(env, USERS.omsTeacher).scheduledHours, 1.5);
  });

  test('a past assignment nobody has filed yet still counts', () => {
    const env = envFor();
    assign(env, { date: day(-3) });
    assert.equal(entry(env, USERS.omsTeacher).scheduledHours, 1);
  });

  test('a recorded assignment counts while its request is pending, then moves into hours once approved', () => {
    const env = envFor();
    const id = assign(env, { date: day(-3) });
    env.run('recordAssignment', id, USERS.omsTeacher);
    assert.equal(entry(env, USERS.omsTeacher).scheduledHours, 1, 'pending: still scheduled');

    env.run('approveEarnedRow', approvalRowFor(env, USERS.omsTeacher), { send: false });
    const e = entry(env, USERS.omsTeacher);
    assert.equal(e.hours, 1, 'approved: now earned');
    assert.equal(e.scheduledHours, 0, 'and no longer scheduled — never counted twice');
  });

  test('a recorded assignment whose request was denied does not count', () => {
    const env = envFor();
    const id = assign(env, { date: day(-3) });
    env.run('recordAssignment', id, USERS.omsTeacher);
    env.run('denyEarnedRow', approvalRowFor(env, USERS.omsTeacher), { send: false, reason: 'x' });
    const e = entry(env, USERS.omsTeacher);
    assert.equal(e.scheduledHours, 0);
    assert.equal(e.hours, 0);
  });

  test('a cancelled assignment does not count', () => {
    const env = envFor();
    const id = assign(env);
    env.run('cancelAssignment', id);
    assert.equal(entry(env, USERS.omsTeacher).scheduledHours, 0);
  });

  test("another building's assignment counts too, like the earned hours", () => {
    const env = envFor(USERS.dualAdmin);
    assign(env, { teacherEmail: USERS.multiTeacher, teacherName: 'Mia Multi', coveredForEmail: USERS.ohsTeacher,
      subbedFor: 'Hank Teacher', period: 'Period 1', building: 'OHS' });
    assert.equal(entry(env, USERS.multiTeacher).scheduledHours, 1);
  });

  test('someone with nothing scheduled reads 0', () => {
    const env = envFor();
    assign(env);
    assert.equal(entry(env, USERS.omsTeacher2).scheduledHours, 0);
  });
};
