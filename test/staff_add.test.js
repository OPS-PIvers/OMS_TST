/**
 * Add Staff when the email is already in the directory.
 *
 * Before: a building admin re-adding someone archived from their own building got
 * "Staff member added." while nothing changed, so the person never appeared. Now
 * the first call writes nothing and describes the person; the one-click fix is a
 * second call with resolveExisting.
 */

const { createEnv } = require('./apps_script_env');
const { USERS, sheets } = require('./fixtures');

const envFor = email => createEnv({ activeUser: email, sheets: sheets() });

const staffRow = (env, email) =>
  env.sheet('Staff Directory').values.find(r => (r[1] || '').toString().toLowerCase() === email);
const BUILDING = 8, ARCHIVED = 9;

const addData = (email, extra) => Object.assign({ name: 'Someone', email: email, role: 'Teacher' }, extra);

exports.name = 'Add Staff: existing people';

exports.run = function ({ test, assert }) {

  test('a new email is added and reported as added', () => {
    const env = envFor(USERS.omsAdmin);
    const res = env.run('addStaffMember', addData('new.person@orono.k12.mn.us', { name: 'New Person' }));
    assert.equal(res.added, true);
    assert.equal(staffRow(env, 'new.person@orono.k12.mn.us')[BUILDING], 'OMS');
  });

  test('re-adding someone archived here writes nothing and says they are archived', () => {
    const env = envFor(USERS.omsAdmin);
    const before = JSON.stringify(env.sheet('Staff Directory').values);
    const res = env.run('addStaffMember', addData(USERS.archivedTeacher.toUpperCase(), { name: 'Arnie Archived' }));
    assert.equal(res.exists, true);
    assert.deepEqual(res.archived, ['OMS']);
    assert.deepEqual(res.missing, []);
    assert.deepEqual(res.active, []);
    assert.equal(res.name, 'Arnie Archived');
    assert.equal(res.nameDiffers, false);
    assert.equal(JSON.stringify(env.sheet('Staff Directory').values), before, 'nothing may change without the fix');
  });

  test('the fix reactivates them here and keeps their balances', () => {
    const env = envFor(USERS.omsAdmin);
    const res = env.run('addStaffMember', addData(USERS.archivedTeacher, { resolveExisting: true, carryOver: 0 }));
    assert.equal(res.resolved, true);
    assert.deepEqual(res.reactivated, ['OMS']);
    assert.deepEqual(res.added, []);
    const row = staffRow(env, USERS.archivedTeacher);
    assert.equal(row[ARCHIVED], '');
    assert.equal(row[BUILDING], 'OMS');
    assert.equal(row[5], 9, 'Carry Over is untouched');
    const dir = env.run('getStaffDirectoryData', 'OMS').map(s => s.email.toLowerCase());
    assert.ok(dir.includes(USERS.archivedTeacher), 'they show in the directory again');
  });

  test('someone from another building is described, then added on the fix', () => {
    const env = envFor(USERS.omsAdmin);
    const res = env.run('addStaffMember', addData(USERS.ohsTeacher, { name: 'Henry Teacher' }));
    assert.equal(res.exists, true);
    assert.deepEqual(res.missing, ['OMS']);
    assert.deepEqual(res.buildings, ['OHS']);
    assert.equal(res.nameDiffers, true);
    assert.equal(staffRow(env, USERS.ohsTeacher)[BUILDING], 'OHS', 'not merged until the admin confirms');

    const fixed = env.run('addStaffMember', addData(USERS.ohsTeacher, { resolveExisting: true }));
    assert.deepEqual(fixed.added, ['OMS']);
    const row = staffRow(env, USERS.ohsTeacher);
    assert.equal(row[BUILDING], 'OHS, OMS', 'appended, so OHS stays primary');
    assert.equal(row[0], 'Hank Teacher', 'existing name kept');
  });

  test('archived elsewhere only: adding here leaves the other archive alone', () => {
    const env = createEnv({ activeUser: USERS.omsAdmin, sheets: sheets() });
    const row = staffRow(env, USERS.ohsTeacher);
    row[ARCHIVED] = 'OHS';
    const res = env.run('addStaffMember', addData(USERS.ohsTeacher));
    assert.deepEqual(res.missing, ['OMS']);
    assert.deepEqual(res.archivedFrom, ['OHS']);
    env.run('addStaffMember', addData(USERS.ohsTeacher, { resolveExisting: true }));
    assert.equal(staffRow(env, USERS.ohsTeacher)[ARCHIVED], 'OHS');
    assert.equal(staffRow(env, USERS.ohsTeacher)[BUILDING], 'OHS, OMS');
  });

  test('a legacy TRUE archive does not swallow the building being added', () => {
    const env = createEnv({ activeUser: USERS.omsAdmin, sheets: sheets() });
    staffRow(env, USERS.ohsTeacher)[ARCHIVED] = 'TRUE';
    env.run('addStaffMember', addData(USERS.ohsTeacher, { resolveExisting: true }));
    const row = staffRow(env, USERS.ohsTeacher);
    assert.equal(row[BUILDING], 'OHS, OMS');
    assert.equal(row[ARCHIVED], 'OHS', 'still archived from OHS, active at OMS');
  });

  test('someone already active here gets a notice and no fix', () => {
    const env = envFor(USERS.omsAdmin);
    const res = env.run('addStaffMember', addData(USERS.omsTeacher));
    assert.equal(res.exists, true);
    assert.deepEqual(res.active, ['OMS']);
    assert.deepEqual(res.missing.concat(res.archived), []);
  });

  test('a multi-building admin adds to the building they are working in', () => {
    const env = envFor(USERS.dualAdmin);
    env.run('addStaffMember', addData('new.ohs@orono.k12.mn.us', { building: 'OHS' }));
    assert.equal(staffRow(env, 'new.ohs@orono.k12.mn.us')[BUILDING], 'OHS');
  });

  test("a building admin cannot name someone else's building", () => {
    const env = envFor(USERS.omsAdmin);
    env.run('addStaffMember', addData('sneaky@orono.k12.mn.us', { building: 'OHS' }));
    assert.equal(staffRow(env, 'sneaky@orono.k12.mn.us')[BUILDING], 'OMS', 'falls back to their own building');
    const res = env.run('addStaffMember', addData(USERS.ohsTeacher, { building: 'OHS', resolveExisting: true }));
    assert.deepEqual(res.added, ['OMS']);
  });

  test('a Super Admin gets the notice too instead of an error, and can fix several buildings', () => {
    const env = envFor(USERS.superAdmin);
    const res = env.run('addStaffMember', addData(USERS.archivedTeacher, { building: 'OMS, OHS' }));
    assert.equal(res.exists, true);
    assert.deepEqual(res.archived, ['OMS']);
    assert.deepEqual(res.missing, ['OHS']);
    const fixed = env.run('addStaffMember', addData(USERS.archivedTeacher, { building: 'OMS, OHS', resolveExisting: true }));
    assert.deepEqual(fixed.added, ['OHS']);
    assert.deepEqual(fixed.reactivated, ['OMS']);
  });

  test('a teacher is refused', () => {
    const env = envFor(USERS.omsTeacher);
    assert.rejected(env.attempt('addStaffMember', addData(USERS.archivedTeacher, { resolveExisting: true })), /admin/i);
    assert.equal(staffRow(env, USERS.archivedTeacher)[ARCHIVED], 'OMS');
  });
};
