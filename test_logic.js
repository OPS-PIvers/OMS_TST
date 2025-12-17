// Mock Environment
const Session = {
  getActiveUser: () => ({ getEmail: () => 'admin@orono.k12.mn.us' })
};

const SpreadsheetApp = {
  getActiveSpreadsheet: () => ({
    getSheetByName: (name) => {
      if (name === 'Staff Directory') {
        return {
          getDataRange: () => ({
            getValues: () => [
              ['Name', 'Email', 'Role', 'Earned', 'Used', 'CarryOver', 'Building'], // Headers
              ['Admin User', 'admin@orono.k12.mn.us', 'Super Admin', 10, 5, 0, 'OMS'],
              ['Teacher 1', 't1@orono.k12.mn.us', 'Teacher', 0, 0, 0, 'OHS'],
              ['Teacher 2', 't2@orono.k12.mn.us', 'Teacher', 0, 0, 0, 'OIS']
            ]
          })
        };
      }
      if (name === 'TST Approvals (New)' || name === 'TST Usage (New)' || name === 'TST Availability') {
          return {
              getDataRange: () => ({ getValues: () => [['Header']] }),
              insertSheet: () => ({ appendRow: () => {} })
          };
      }
      return null;
    }
  })
};

// Mock Utilities
function safeDate(d) { return d; }

// --- Load Code ---
// (We will concatenate this file with config.js and Code.js in the bash command)

// --- Tests ---
function runTests() {
  console.log("Running Tests...");

  try {
    const initData = getInitialData();
    console.log("getInitialData Result:", JSON.stringify(initData, null, 2));

    if (initData.email !== 'admin@orono.k12.mn.us') throw new Error("Email mismatch");
    if (!initData.isSuperAdmin) throw new Error("Should be Super Admin");
    if (initData.building !== 'OMS') throw new Error("Default building mismatch");
    if (initData.config['OHS'].name !== 'Orono High School') throw new Error("Config mismatch");

    console.log("SUCCESS: getInitialData passed.");

    // Test Data Fetching with Filter (Simulated)
    // Note: Since mocks return empty data for other sheets, we just check function execution
    const pending = getPendingEarned('OHS');
    console.log("getPendingEarned('OHS') executed. Count:", pending.length);

  } catch (e) {
    console.error("TEST FAILED:", e);
    process.exit(1);
  }
}

// execute
runTests();
