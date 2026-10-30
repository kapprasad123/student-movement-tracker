/**
 * Student Movement Tracker — Separate Test Spreadsheet Creator & Connector
 * (scripts/create-test-sheet.js)
 *
 * Responsibilities:
 * 1. Creates a standalone, isolated test spreadsheet runtime completely detached
 *    from real school data.
 * 2. Initializes all 8 required sheets with authoritative schemas and headers:
 *    - Students, Classes, Movement_Log, Settings, Teachers, Teaching_Assignments,
 *      Leading_Teachers, Leading_Teacher_Assignments.
 * 3. Populates realistic test data across all roles (Classes, Teachers, Students,
 *    Leading Teachers, Movements, Passcodes).
 * 4. Exports offline CSV copies to test-spreadsheet/ for inspection or manual Google Sheets import.
 * 5. Connects the local application to this test environment via the local API endpoint.
 */

const fs = require('fs');
const path = require('path');
const { createGasEnvironment } = require('../tests/gas-mock.js');
const Code = require('../google-apps-script/Code.gs');
const { seedEnvironment, exportCsvFiles, TEST_CREDENTIALS } = require('./seed-test-spreadsheet.js');

const ROOT_DIR = path.resolve(__dirname, '..');
const TEST_DIR = path.join(ROOT_DIR, 'test-spreadsheet');
const CONFIG_FILE = path.join(ROOT_DIR, 'config.json');

function createAndConnectTestSpreadsheet(options = {}) {
  console.log('============================================================');
  console.log('Student Movement Tracker — Creating Separate Test Spreadsheet');
  console.log('============================================================');
  console.log('Safety Guarantee: Zero real school data will be touched.');

  // 1. Instantiate completely isolated virtual Google Sheets environment
  const gasEnv = createGasEnvironment();
  
  // 2. Seed all 8 sheets with realistic test data & salted SHA-256 test credentials
  seedEnvironment(Code, gasEnv);
  console.log('-> Initialized and seeded all 8 required sheets in isolated test environment.');

  // 3. Export CSV representations to test-spreadsheet/
  const exportedFiles = exportCsvFiles(gasEnv, TEST_DIR);
  console.log(`-> Exported ${exportedFiles.length} sheet CSV files to ${TEST_DIR}:`);
  exportedFiles.forEach((f) => console.log(`   - ${path.basename(f)}`));

  // 4. Update config.json to connect to the test spreadsheet environment
  try {
    let configData = {};
    if (fs.existsSync(CONFIG_FILE)) {
      configData = JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8'));
    }
    configData.activeEnvironment = 'test';
    configData.testSpreadsheetPath = TEST_DIR;
    configData.testCredentials = {
      adminPasscode: TEST_CREDENTIALS.adminPasscode,
      principalPasscode: TEST_CREDENTIALS.principalPasscode,
      leadingTeacherPasscode: TEST_CREDENTIALS.leadingTeacherPasscode,
      teacherPin: TEST_CREDENTIALS.teacherPin
    };
    configData.connectedEndpoint = 'http://localhost:3000/api';

    fs.writeFileSync(CONFIG_FILE, JSON.stringify(configData, null, 2), 'utf8');
    console.log('-> Updated config.json: application is now connected to the test spreadsheet.');
  } catch (err) {
    console.warn('-> Warning updating config.json:', err.message);
  }

  console.log('============================================================');
  console.log('Test Spreadsheet Successfully Created & Connected!');
  console.log('Test Credentials:');
  console.log(`  Admin Passcode:           ${TEST_CREDENTIALS.adminPasscode}`);
  console.log(`  Principal Passcode:       ${TEST_CREDENTIALS.principalPasscode}`);
  console.log(`  Leading Teacher Passcode: ${TEST_CREDENTIALS.leadingTeacherPasscode}`);
  console.log(`  Teacher App PIN:          ${TEST_CREDENTIALS.teacherPin}`);
  console.log('============================================================');

  return { gasEnv, exportedFiles, testCredentials: TEST_CREDENTIALS };
}

if (require.main === module) {
  createAndConnectTestSpreadsheet();
}

module.exports = { createAndConnectTestSpreadsheet };

