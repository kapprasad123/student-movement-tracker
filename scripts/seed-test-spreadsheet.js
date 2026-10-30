/**
 * Student Movement Tracker — Test Spreadsheet Seeder & Exporter (scripts/seed-test-spreadsheet.js)
 *
 * Creates a standalone, isolated test spreadsheet environment with:
 * - Pre-configured roles & passcodes
 * - Realistic classes, teachers, leading teachers, and students
 * - Sample movement records
 * - Complete CSV exports for all 8 required sheets in test-spreadsheet/
 *
 * Guarantees zero interaction with real school data.
 */

const fs = require('fs');
const path = require('path');

const TEST_CREDENTIALS = {
  adminPasscode: 'AdminSecret#2026',
  principalPasscode: 'PrincipalPass#2026',
  leadingTeacherPasscode: 'LeadTeacher#2026',
  teacherPin: '4829'
};

const SAMPLE_CLASSES = [
  { ClassID: 'CLS-10A', ClassName: 'Grade 10A', Status: 'ACTIVE' },
  { ClassID: 'CLS-10B', ClassName: 'Grade 10B', Status: 'ACTIVE' },
  { ClassID: 'CLS-11A', ClassName: 'Grade 11 Science', Status: 'ACTIVE' }
];

const SAMPLE_TEACHERS = [
  { TeacherID: 'TCH-01', TeacherName: 'Mr. Ibrahim Didi', Status: 'ACTIVE' },
  { TeacherID: 'TCH-02', TeacherName: 'Ms. Mariyam Waheed', Status: 'ACTIVE' },
  { TeacherID: 'TCH-03', TeacherName: 'Mr. Ali Rasheed', Status: 'ACTIVE' }
];

const SAMPLE_TEACHING_ASSIGNMENTS = [
  { AssignmentID: 'ASN-001', TeacherID: 'TCH-01', ClassID: 'CLS-10A', Status: 'ACTIVE' },
  { AssignmentID: 'ASN-002', TeacherID: 'TCH-02', ClassID: 'CLS-10B', Status: 'ACTIVE' },
  { AssignmentID: 'ASN-003', TeacherID: 'TCH-03', ClassID: 'CLS-11A', Status: 'ACTIVE' }
];

const SAMPLE_LEADING_TEACHERS = [
  { LeadingTeacherID: 'LT-01', LeadingTeacherName: 'Mrs. Hawwa Zahir', Status: 'ACTIVE' },
  { LeadingTeacherID: 'LT-02', LeadingTeacherName: 'Mr. Hassan Manik', Status: 'ACTIVE' }
];

const SAMPLE_LEADING_TEACHER_ASSIGNMENTS = [
  { AssignmentID: 'LTASN-001', LeadingTeacherID: 'LT-01', ClassID: 'CLS-10A', Period: '2026 Term 1', Status: 'ACTIVE' },
  { AssignmentID: 'LTASN-002', LeadingTeacherID: 'LT-01', ClassID: 'CLS-10B', Period: '2026 Term 1', Status: 'ACTIVE' },
  { AssignmentID: 'LTASN-003', LeadingTeacherID: 'LT-02', ClassID: 'CLS-11A', Period: '2026 Term 1', Status: 'ACTIVE' }
];

const SAMPLE_STUDENTS = [
  { StudentID: 'STU-101', StudentName: 'Ahmed Ali', ClassID: 'CLS-10A', Status: 'ACTIVE' },
  { StudentID: 'STU-102', StudentName: 'Fathimath Shifa', ClassID: 'CLS-10A', Status: 'ACTIVE' },
  { StudentID: 'STU-103', StudentName: 'Hussain Zayan', ClassID: 'CLS-10B', Status: 'ACTIVE' },
  { StudentID: 'STU-104', StudentName: 'Mariyam Raudha', ClassID: 'CLS-10B', Status: 'ACTIVE' },
  { StudentID: 'STU-105', StudentName: 'Mohamed Shuaib', ClassID: 'CLS-11A', Status: 'ACTIVE' },
  { StudentID: 'STU-106', StudentName: 'Aishath Leen', ClassID: 'CLS-11A', Status: 'ACTIVE' }
];

function seedEnvironment(Code, gasEnv) {
  Code.setupInitialCredentials(TEST_CREDENTIALS);

  // Initialize all sheets
  Code.doGet({ parameter: { action: 'getSettings' } });

  const now = new Date();
  const nowIso = now.toISOString();
  const eightMinsAgo = new Date(now.getTime() - 8 * 60 * 1000).toISOString();

  // Seed Classes
  const classesSheet = gasEnv.spreadsheet.getSheetByName('Classes');
  for (const c of SAMPLE_CLASSES) {
    const nextRow = classesSheet.getLastRow() + 1;
    classesSheet.getRange(nextRow, 1, 1, 5).setValues([[c.ClassID, c.ClassName, c.Status, nowIso, nowIso]]);
  }

  // Seed Teachers
  const teachersSheet = gasEnv.spreadsheet.getSheetByName('Teachers');
  for (const t of SAMPLE_TEACHERS) {
    const nextRow = teachersSheet.getLastRow() + 1;
    teachersSheet.getRange(nextRow, 1, 1, 5).setValues([[t.TeacherID, t.TeacherName, t.Status, nowIso, nowIso]]);
  }

  // Seed Teaching Assignments
  const teachAsnSheet = gasEnv.spreadsheet.getSheetByName('Teaching_Assignments');
  for (const ta of SAMPLE_TEACHING_ASSIGNMENTS) {
    const nextRow = teachAsnSheet.getLastRow() + 1;
    teachAsnSheet.getRange(nextRow, 1, 1, 6).setValues([[ta.AssignmentID, ta.TeacherID, ta.ClassID, ta.Status, nowIso, nowIso]]);
  }

  // Seed Leading Teachers
  const ltSheet = gasEnv.spreadsheet.getSheetByName('Leading_Teachers');
  for (const lt of SAMPLE_LEADING_TEACHERS) {
    const nextRow = ltSheet.getLastRow() + 1;
    ltSheet.getRange(nextRow, 1, 1, 5).setValues([[lt.LeadingTeacherID, lt.LeadingTeacherName, lt.Status, nowIso, nowIso]]);
  }

  // Seed Leading Teacher Assignments
  const ltaSheet = gasEnv.spreadsheet.getSheetByName('Leading_Teacher_Assignments');
  for (const lta of SAMPLE_LEADING_TEACHER_ASSIGNMENTS) {
    const nextRow = ltaSheet.getLastRow() + 1;
    ltaSheet.getRange(nextRow, 1, 1, 7).setValues([[lta.AssignmentID, lta.LeadingTeacherID, lta.ClassID, lta.Period, lta.Status, nowIso, nowIso]]);
  }

  // Seed Students
  const stuSheet = gasEnv.spreadsheet.getSheetByName('Students');
  for (const s of SAMPLE_STUDENTS) {
    const nextRow = stuSheet.getLastRow() + 1;
    stuSheet.getRange(nextRow, 1, 1, 6).setValues([[s.StudentID, s.StudentName, s.ClassID, s.Status, nowIso, nowIso]]);
  }

  // Seed 1 active Movement Log entry (Ahmed Ali out to Library for 8 minutes)
  const movSheet = gasEnv.spreadsheet.getSheetByName('Movement_Log');
  const mRow = movSheet.getLastRow() + 1;
  movSheet.getRange(mRow, 1, 1, 12).setValues([[
    'MOV-' + Date.now() + '-0001',
    'STU-101',
    'Ahmed Ali',
    'CLS-10A',
    'Grade 10A',
    'TCH-01',
    'Mr. Ibrahim Didi',
    'Library',
    eightMinsAgo,
    '',
    'OUT',
    'Reference book collection'
  ]]);

  return gasEnv;
}

function exportCsvFiles(gasEnv, outDir) {
  if (!fs.existsSync(outDir)) {
    fs.mkdirSync(outDir, { recursive: true });
  }

  const sheets = gasEnv.spreadsheet.getSheets();
  const exported = [];

  for (const sheet of sheets) {
    const name = sheet.getName();
    const rows = sheet.getLastRow();
    const cols = sheet.getLastColumn();
    if (rows === 0 || cols === 0) continue;

    const values = sheet.getRange(1, 1, rows, cols).getValues();
    const csvContent = values.map((row) =>
      row.map((cell) => {
        const str = String(cell !== undefined && cell !== null ? cell : '');
        if (str.includes(',') || str.includes('"') || str.includes('\n')) {
          return '"' + str.replace(/"/g, '""') + '"';
        }
        return str;
      }).join(',')
    ).join('\n');

    const filePath = path.join(outDir, `${name}.csv`);
    fs.writeFileSync(filePath, csvContent, 'utf8');
    exported.push(filePath);
  }

  return exported;
}

module.exports = {
  TEST_CREDENTIALS,
  seedEnvironment,
  exportCsvFiles
};
