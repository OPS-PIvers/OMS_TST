/**
 * A small two-building district used by the authorization tests.
 *
 * Staff Directory columns follow the documented schema:
 *   A Name | B Email | C Role | D Earned | E Used | F Carry Over | G Paid Out |
 *   H Running Total (ARRAYFORMULA) | I Building | J Archived | K Last Finalized |
 *   L Pending Finalize
 */

const USERS = {
  superAdmin: 'sam.super@orono.k12.mn.us',
  omsAdmin: 'amy.admin@orono.k12.mn.us',
  ohsAdmin: 'otto.admin@orono.k12.mn.us',
  dualAdmin: 'dana.dual@orono.k12.mn.us',   // Admin at OMS + OHS
  omsTeacher: 'tina.teacher@orono.k12.mn.us',
  omsTeacher2: 'ted.teacher@orono.k12.mn.us',
  ohsTeacher: 'hank.teacher@orono.k12.mn.us',
  multiTeacher: 'mia.multi@orono.k12.mn.us', // OMS primary, also at OHS
  archivedTeacher: 'arnie.archived@orono.k12.mn.us',
  stranger: 'nobody@orono.k12.mn.us'         // signed in, not in the directory
};

const STAFF_HEADER = ['Name', 'Email', 'Role', 'Earned', 'Used', 'Carry Over', 'Paid Out',
  'Running Total', 'Building', 'Archived', 'Last Finalized', 'Pending Finalize'];

// Name, Email, Role, Earned, Used, CarryOver, PaidOut, RunningTotal, Building, Archived, LastFinalized, PendingFinalize
const STAFF_ROWS = [
  ['Sam Super', USERS.superAdmin, 'Super Admin', '', '', 0, 0, '', 'OMS', '', '', ''],
  ['Amy Admin', USERS.omsAdmin, 'Admin', '', '', 1, 0, '', 'OMS', '', '', ''],
  ['Otto Admin', USERS.ohsAdmin, 'Admin', '', '', 2, 0, '', 'OHS', '', '', ''],
  ['Dana Dual', USERS.dualAdmin, 'Admin', '', '', 0, 0, '', 'OMS, OHS', '', '', ''],
  ['Tina Teacher', USERS.omsTeacher, 'Teacher', '', '', 3, 1, '', 'OMS', '', '', ''],
  ['Ted Teacher', USERS.omsTeacher2, 'Teacher', '', '', 7, 2, '', 'OMS', '', '', ''],
  ['Hank Teacher', USERS.ohsTeacher, 'Teacher', '', '', 5, 0, '', 'OHS', '', '', ''],
  ['Mia Multi', USERS.multiTeacher, 'Teacher', '', '', 4, 0, '', 'OMS, OHS', '', '', ''],
  ['Arnie Archived', USERS.archivedTeacher, 'Teacher', '', '', 9, 0, '', 'OMS', 'OMS', '', '']
];

// TST Approvals (New): A Email | B Name | C SubbedFor | D (unused) | E Date | F Period |
// G TimeType | H Hours | I Approved | J ApprovedTS | K Denied | L DeniedTS | M Reason | N Building
const APPROVALS_HEADER = ['Email', 'Name', 'Subbed For', 'Other', 'Date', 'Period',
  'Time Type', 'Hours', 'Approved', 'Approved TS', 'Denied', 'Denied TS', 'Denial Reason', 'Building'];

const approvalRow = (email, name, subbedFor, date, period, hours, approved, denied, building) =>
  [email, name, subbedFor, '', date, period, 'Full Period', hours, approved, '', denied, '', '', building];

const APPROVALS_ROWS = [
  // Approved OMS earnings for Tina (2.5) — combined-balance material
  approvalRow(USERS.omsTeacher, 'Tina Teacher', 'Ted Teacher', '2025-09-10', 'Period 1 - 8:10 - 8:57', 1, true, false, 'OMS'),
  approvalRow(USERS.omsTeacher, 'Tina Teacher', 'Ted Teacher', '2025-09-12', 'Period 2 - 9:01 - 9:48', 1.5, true, false, 'OMS'),
  // Pending at OMS
  approvalRow(USERS.omsTeacher2, 'Ted Teacher', 'Tina Teacher', '2025-09-15', 'Period 3 - 9:52 - 10:39', 1, false, false, 'OMS'),
  // Pending at OHS
  approvalRow(USERS.ohsTeacher, 'Hank Teacher', 'Otto Admin', '2025-09-16', 'Period 1', 1, false, false, 'OHS'),
  // Mia earns at both buildings; both approved (combined = 2)
  approvalRow(USERS.multiTeacher, 'Mia Multi', 'Tina Teacher', '2025-09-17', 'Period 4 - 10:43 - 11:09', 1, true, false, 'OMS'),
  approvalRow(USERS.multiTeacher, 'Mia Multi', 'Hank Teacher', '2025-09-18', 'Period 2', 1, true, false, 'OHS'),
  // Blank building = OMS, matching the queues' default
  approvalRow(USERS.omsTeacher, 'Tina Teacher', 'Ted Teacher', '2025-09-19', 'Period 5 - 11:11 - 11:37', 1, false, false, '')
];

// TST Usage (New): A Email | B Name | C Date | D Used | E Status | F Timestamp | G Notes | H Building
const USAGE_HEADER = ['Email', 'Name', 'Date', 'Used', 'Status', 'Timestamp', 'Notes', 'Building'];

const USAGE_ROWS = [
  [USERS.omsTeacher, 'Tina Teacher', '2025-10-01', 1, true, '2025-10-01', '', 'OMS'],
  [USERS.omsTeacher2, 'Ted Teacher', '2025-10-02', 2, false, '2025-10-02', '', 'OMS'],
  [USERS.ohsTeacher, 'Hank Teacher', '2025-10-03', 1, false, '2025-10-03', '', 'OHS']
];

// submitEarned_ mirrors every submission into Form Responses 1 before processing it,
// so anything that records an earned request needs this sheet present.
const FORM_RESPONSES_HEADER = ['Timestamp', 'Email', 'Subbed For', 'Other', 'Date',
  'Period', 'Time Type', 'Hours'];

const AVAILABILITY_HEADER = ['Month', 'Day(s) Available', 'Period', 'Name', 'Email', 'Hours Earned This Month'];

const AVAILABILITY_ROWS = [
  ['September', 'Mon,Tue', 'Period 1 - 8:10 - 8:57', 'Tina Teacher', USERS.omsTeacher, ''],
  ['September', 'Wed', 'Period 2 - 9:01 - 9:48', 'Ted Teacher', USERS.omsTeacher2, ''],
  ['September', 'Thu', 'Period 1', 'Hank Teacher', USERS.ohsTeacher, ''],
  ['October', 'Fri', 'Period 1 - 8:10 - 8:57', 'Mia Multi', USERS.multiTeacher, '']
];

/** Fresh copies every call so a test that writes can't affect the next one. */
function sheets(overrides) {
  const base = {
    'Staff Directory': [STAFF_HEADER, ...STAFF_ROWS],
    'TST Approvals (New)': [APPROVALS_HEADER, ...APPROVALS_ROWS],
    'TST Usage (New)': [USAGE_HEADER, ...USAGE_ROWS],
    'TST Availability': [AVAILABILITY_HEADER, ...AVAILABILITY_ROWS],
    'Form Responses 1': [FORM_RESPONSES_HEADER]
  };
  const out = {};
  Object.keys(base).forEach(k => { out[k] = base[k].map(r => r.slice()); });
  Object.keys(overrides || {}).forEach(k => { out[k] = overrides[k]; });
  return out;
}

module.exports = { USERS, STAFF_HEADER, STAFF_ROWS, APPROVALS_HEADER, USAGE_HEADER, sheets };
