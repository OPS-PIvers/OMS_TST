/**
 * Per-person display preferences (name order). They are the signed-in person's
 * own and nobody else's, so the rules are: anyone in the directory may set their
 * own, nobody can set anyone else's, and getInitialData hands them back.
 */

const { createEnv } = require('./apps_script_env');
const { USERS, sheets } = require('./fixtures');

const envFor = (email, properties) => createEnv({ activeUser: email, sheets: sheets(), properties });

exports.name = 'User preferences';

exports.run = function ({ test, assert }) {

  test('everyone starts on First Last', () => {
    const env = envFor(USERS.omsAdmin);
    assert.deepEqual(env.run('getInitialData').preferences, { nameOrder: 'first' });
  });

  test('an admin can switch to Last, First and it comes back on the next load', () => {
    const env = envFor(USERS.omsAdmin);
    assert.deepEqual(env.run('saveMyPreferences', { nameOrder: 'last' }), { nameOrder: 'last' });
    assert.equal(env.run('getInitialData').preferences.nameOrder, 'last');
  });

  test('a teacher can set their own', () => {
    const env = envFor(USERS.omsTeacher);
    assert.allowed(env.attempt('saveMyPreferences', { nameOrder: 'last' }));
    assert.equal(env.run('getInitialData').preferences.nameOrder, 'last');
  });

  test("one person's preference does not change anyone else's", () => {
    const admin = envFor(USERS.omsAdmin);
    admin.run('saveMyPreferences', { nameOrder: 'last' });

    // Same Script Properties, different signed-in people.
    [USERS.ohsAdmin, USERS.superAdmin, USERS.omsTeacher].forEach(email => {
      const other = envFor(email, admin.properties);
      assert.equal(other.run('getInitialData').preferences.nameOrder, 'first', email);
    });
  });

  test('the key is the email, whatever its case', () => {
    const env = envFor(USERS.superAdmin);
    env.run('saveMyPreferences', { nameOrder: 'last' });
    assert.ok(('PREFS_' + USERS.superAdmin.toLowerCase()) in env.properties);
  });

  test('someone not in the directory cannot save preferences', () => {
    const env = envFor(USERS.stranger);
    assert.rejected(env.attempt('saveMyPreferences', { nameOrder: 'last' }), /not listed/i);
    assert.equal(Object.keys(env.properties).filter(k => /^PREFS_/.test(k)).length, 0);
  });

  test('an unknown name order is refused rather than stored', () => {
    const env = envFor(USERS.omsAdmin);
    assert.rejected(env.attempt('saveMyPreferences', { nameOrder: '<script>' }), /unknown name order/i);
    assert.equal(env.run('getInitialData').preferences.nameOrder, 'first');
  });

  test('the name list covers every building and archived staff, and nothing but names', () => {
    const env = envFor(USERS.omsAdmin);
    const names = env.run('getInitialData').staffNames;
    assert.ok(names.includes('Arnie Archived'), 'archived staff are still named in old requests');
    assert.ok(names.some(n => /hank/i.test(n)), 'coverage can be for someone at another building');
    assert.ok(names.every(n => typeof n === 'string' && !n.includes('@')), 'names only');
  });

  test('someone not in the directory gets no names', () => {
    const env = envFor(USERS.stranger);
    assert.deepEqual(env.run('getInitialData').staffNames, []);
  });

  test('a damaged stored value falls back to the default', () => {
    const env = envFor(USERS.omsAdmin, { ['PREFS_' + USERS.omsAdmin]: '{not json' });
    assert.equal(env.run('getInitialData').preferences.nameOrder, 'first');
  });
};
