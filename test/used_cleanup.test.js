/**
 * archiveUsedBefore / restoreArchivedUsed: the one-off repair for a building that
 * skipped Finalize and set Carry Over by hand, leaving last year's approved Used
 * rows counted against this year. It must only move what finalize would have
 * moved, do nothing without apply === true, and be fully undoable.
 */

const { createEnv } = require('./apps_script_env');
const { USERS, USAGE_HEADER, sheets } = require('./fixtures');

const LABEL = '2025-26 OHS Used cleanup';

// Hank: OHS primary. Mia: OMS primary, also at OHS. Tina: OMS only.
const usage = () => [USAGE_HEADER,
  [USERS.ohsTeacher, 'Hank Teacher', '2026-03-10', 2, true, '2026-03-10', '', 'OHS'],   // moves
  [USERS.ohsTeacher, 'Hank Teacher', '2026-05-20', 1.5, true, '2026-05-20', '', 'OMS'], // moves: primary OHS, any tag
  [USERS.ohsTeacher, 'Hank Teacher', '2026-05-21', 1, false, '', '', 'OHS'],            // pending: reported, stays
  [USERS.ohsTeacher, 'Hank Teacher', '2026-09-15', 1, true, '2026-09-15', '', 'OHS'],   // this year: stays
  [USERS.ohsTeacher, 'Hank Teacher', 'sometime', 1, true, '', '', 'OHS'],               // unreadable: reported, stays
  [USERS.multiTeacher, 'Mia Multi', '2026-04-01', 3, true, '2026-04-01', '', 'OHS'],    // primary OMS: stays
  [USERS.omsTeacher, 'Tina Teacher', '2026-04-02', 1, true, '2026-04-02', '', 'OMS']    // OMS: stays
];

const envFor = email => createEnv({ activeUser: email, sheets: sheets({ 'TST Usage (New)': usage() }) });
// What the Directory shows — the number the OHS admin asked about.
const usedFor = (env, email) => env.run('getStaffDirectoryData', email === USERS.omsTeacher ? 'OMS' : 'OHS')
  .find(s => s.email === email).used;

exports.name = 'Used cleanup (archiveUsedBefore)';

exports.run = function ({ test, assert }) {

  test('only a Super Admin can run it or undo it', () => {
    [USERS.ohsAdmin, USERS.omsAdmin, USERS.ohsTeacher].forEach(email => {
      const env = envFor(email);
      assert.rejected(env.attempt('archiveUsedBefore', 'OHS', '2026-07-01', LABEL, true), /super admin/i, email);
      assert.rejected(env.attempt('restoreArchivedUsed', LABEL, true), /super admin/i, email);
    });
  });

  test('a dry run reports the rows and changes nothing', () => {
    const env = envFor(USERS.superAdmin);
    const before = JSON.stringify(env.sheet('TST Usage (New)').values);
    const r = env.run('archiveUsedBefore', 'OHS', '2026-07-01', LABEL);
    assert.equal(r.applied, false);
    assert.equal(r.rowCount, 2);
    assert.equal(r.totalHours, 3.5);
    assert.deepEqual(r.byPerson.map(p => [p.email, p.rows, p.hours]), [[USERS.ohsTeacher, 2, 3.5]]);
    assert.equal(r.skippedPending.length, 1);
    assert.equal(r.unreadableDates.length, 1);
    assert.equal(JSON.stringify(env.sheet('TST Usage (New)').values), before);
    assert.equal(env.sheet('TST Usage Archive'), null);
  });

  test('applying moves only last year\'s approved rows for primary-building staff', () => {
    const env = envFor(USERS.superAdmin);
    assert.equal(usedFor(env, USERS.ohsTeacher), 5.5);
    env.run('archiveUsedBefore', 'OHS', '2026-07-01', LABEL, true);

    assert.equal(usedFor(env, USERS.ohsTeacher), 2, 'this year plus the unreadable row');
    assert.equal(usedFor(env, USERS.multiTeacher), 3, 'OMS-primary staff are untouched');
    assert.equal(usedFor(env, USERS.omsTeacher), 1);
    const arch = env.sheet('TST Usage Archive').values;
    assert.equal(arch.length, 3);
    assert.ok(arch.slice(1).every(r => r[USAGE_HEADER.length] === LABEL));
  });

  test('it does not touch Carry Over, Paid Out or Earned', () => {
    const env = envFor(USERS.superAdmin);
    const staff = JSON.stringify(env.sheet('Staff Directory').values);
    const approvals = JSON.stringify(env.sheet('TST Approvals (New)').values);
    env.run('archiveUsedBefore', 'OHS', '2026-07-01', LABEL, true);
    assert.equal(JSON.stringify(env.sheet('Staff Directory').values), staff);
    assert.equal(JSON.stringify(env.sheet('TST Approvals (New)').values), approvals);
  });

  test('restoring the label puts every row back', () => {
    const env = envFor(USERS.superAdmin);
    env.run('archiveUsedBefore', 'OHS', '2026-07-01', LABEL, true);
    assert.equal(env.run('restoreArchivedUsed', LABEL).rowCount, 2, 'dry run counts them');
    assert.equal(usedFor(env, USERS.ohsTeacher), 2, 'dry run changes nothing');

    env.run('restoreArchivedUsed', LABEL, true);
    assert.equal(usedFor(env, USERS.ohsTeacher), 5.5);
    assert.equal(env.sheet('TST Usage Archive').values.length, 1, 'header only');
    const sortRows = rows => rows.map(r => JSON.stringify(r)).sort();
    assert.deepEqual(sortRows(env.sheet('TST Usage (New)').values), sortRows(usage()));
  });

  test('restore leaves rows archived under other labels (e.g. a real finalize) alone', () => {
    const env = envFor(USERS.superAdmin);
    env.run('archiveUsedBefore', 'OHS', '2026-07-01', LABEL, true);
    env.sheet('TST Usage Archive').values.push(
      [USERS.omsTeacher, 'Tina Teacher', '2025-04-01', 4, true, '', '', 'OMS', '2024-25']);
    env.run('restoreArchivedUsed', LABEL, true);
    assert.equal(env.sheet('TST Usage Archive').values.length, 2);
  });

  test('a label already in the archive is refused, so an undo can never mix two runs', () => {
    const env = envFor(USERS.superAdmin);
    env.run('archiveUsedBefore', 'OHS', '2026-07-01', LABEL, true);
    assert.rejected(env.attempt('archiveUsedBefore', 'OHS', '2026-09-01', LABEL, true), /already has rows/i);
  });

  test('building, cutoff and label are all required', () => {
    const env = envFor(USERS.superAdmin);
    assert.rejected(env.attempt('archiveUsedBefore', '', '2026-07-01', LABEL, true), /name the building/i);
    assert.rejected(env.attempt('archiveUsedBefore', 'OHS', 'not a date', LABEL, true), /cutoff/i);
    assert.rejected(env.attempt('archiveUsedBefore', 'OHS', '2026-07-01', '', true), /label/i);
  });
};
