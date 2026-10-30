/**
 * Student Movement Tracker — Comprehensive Automated Test Suite (tests/run-tests.js)
 */

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const http = require('http');
const { createGasEnvironment } = require('./gas-mock.js');

const gasEnv = createGasEnvironment();
const Code = require('../google-apps-script/Code.gs');
const SMT_UTILS = require('../js/utils.js');
const SMT_CONFIG = require('../js/config.js');
const SMT_API = require('../js/api.js');
const { createServer } = require('../server.js');

function invokeGet(params) {
  const out = Code.doGet({ parameter: params });
  return JSON.parse(out.getContent());
}

function invokePost(bodyObj) {
  const out = Code.doPost({
    parameter: {},
    postData: { contents: JSON.stringify(bodyObj) }
  });
  return JSON.parse(out.getContent());
}

let totalTests = 0;
let passedTests = 0;
let failedTests = 0;

async function runTest(name, fn) {
  totalTests++;
  try {
    await fn();
    passedTests++;
    console.log(`  [PASS] ${name}`);
  } catch (err) {
    failedTests++;
    console.error(`  [FAIL] ${name}`);
    console.error(`         -> ${err.message}`);
  }
}

async function main() {
  console.log('============================================================');
  console.log('Student Movement Tracker — Automated Test Suite');
  console.log('============================================================');

  Code.setupInitialCredentials({
    adminPasscode: 'AdminSecret#2026',
    principalPasscode: 'PrincipalPass#2026',
    leadingTeacherPasscode: 'LeadTeacher#2026',
    teacherPin: '4829'
  });

  // 1. Workbook Initialization & Header Mapping
  console.log('\n1. Sheet Initialization & Header Mapping');

  await runTest('Initializes all 8 required sheets with exact headers without duplicating headers on startup', () => {
    const res = invokeGet({ action: 'getSettings' });
    assert.strictEqual(res.success, true);
    const sheetNames = Object.values(Code.SMT_SHEETS).map((s) => s.name);
    for (const sName of sheetNames) {
      const sheet = gasEnv.spreadsheet.getSheetByName(sName);
      assert.ok(sheet, `Sheet ${sName} should exist`);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const expected = Code.SMT_SHEETS[Object.keys(Code.SMT_SHEETS).find((k) => Code.SMT_SHEETS[k].name === sName)].headers;
      assert.deepStrictEqual(headers, expected);
    }
    const res2 = invokeGet({ action: 'getSettings' });
    assert.strictEqual(res2.success, true);
  });

  // 2. Role Logins & Invalid Credentials
  console.log('\n2. Authentication & Role Logins');

  let adminToken = '';
  let principalToken = '';
  let leadingTeacherToken = '';
  let teacher1Token = '';
  let teacher2Token = '';

  await runTest('Admin login succeeds with valid passcode and fails with invalid passcode', () => {
    const badRes = invokePost({ action: 'verifyPasscode', role: 'admin', passcode: 'WrongPass' });
    assert.strictEqual(badRes.success, false);
    assert.strictEqual(badRes.code, 'INVALID_CREDENTIALS');

    const okRes = invokePost({ action: 'verifyPasscode', role: 'admin', passcode: 'AdminSecret#2026' });
    assert.strictEqual(okRes.success, true);
    assert.ok(okRes.data.token);
    assert.strictEqual(okRes.data.role, 'admin');
    adminToken = okRes.data.token;
  });

  await runTest('Principal login succeeds with valid passcode and rejects wrong passcode', () => {
    const bad = invokePost({ action: 'verifyPasscode', role: 'principal', passcode: 'AdminSecret#2026' });
    assert.strictEqual(bad.success, false);

    const ok = invokePost({ action: 'verifyPasscode', role: 'principal', passcode: 'PrincipalPass#2026' });
    assert.strictEqual(ok.success, true);
    assert.strictEqual(ok.data.role, 'principal');
    principalToken = ok.data.token;
  });

  await runTest('Admin creates Classes, Teachers, and Leading Teachers for login & assignment tests', () => {
    const c1 = invokePost({ action: 'addClass', token: adminToken, classId: 'CLS-10A', className: 'Grade 10A' });
    const c2 = invokePost({ action: 'addClass', token: adminToken, classId: 'CLS-10B', className: 'Grade 10B' });
    assert.strictEqual(c1.success, true);
    assert.strictEqual(c2.success, true);

    const t1 = invokePost({ action: 'addTeacher', token: adminToken, teacherId: 'TCH-01', teacherName: 'Mr. Ibrahim Didi' });
    const t2 = invokePost({ action: 'addTeacher', token: adminToken, teacherId: 'TCH-02', teacherName: 'Ms. Mariyam Waheed' });
    assert.strictEqual(t1.success, true);
    assert.strictEqual(t2.success, true);

    const lt1 = invokePost({ action: 'addLeadingTeacher', token: adminToken, leadingTeacherId: 'LT-01', leadingTeacherName: 'Mrs. Hawwa Zahir' });
    assert.strictEqual(lt1.success, true);

    const lta = invokePost({
      action: 'addLeadingTeacherAssignment',
      token: adminToken,
      leadingTeacherId: 'LT-01',
      classId: 'CLS-10A',
      period: '2026 Term 1'
    });
    assert.strictEqual(lta.success, true);
  });

  await runTest('Leading Teacher login succeeds when profile and passcode match, and binds assignedClassIds', () => {
    const noProfile = invokePost({ action: 'verifyPasscode', role: 'leading_teacher', passcode: 'LeadTeacher#2026' });
    assert.strictEqual(noProfile.success, false);
    assert.strictEqual(noProfile.code, 'MISSING_LEADING_TEACHER_ID');

    const ok = invokePost({
      action: 'verifyPasscode',
      role: 'leading_teacher',
      leadingTeacherId: 'LT-01',
      passcode: 'LeadTeacher#2026'
    });
    assert.strictEqual(ok.success, true);
    assert.strictEqual(ok.data.role, 'leading_teacher');
    assert.strictEqual(ok.data.leadingTeacherId, 'LT-01');
    assert.deepStrictEqual(ok.data.assignedClassIds, ['CLS-10A']);
    leadingTeacherToken = ok.data.token;
  });

  await runTest('Teacher login succeeds with active TeacherID + Teacher App PIN and issues bound token', () => {
    const badPin = invokePost({ action: 'verifyTeacherPasscode', teacherId: 'TCH-01', pin: '0000' });
    assert.strictEqual(badPin.success, false);
    assert.strictEqual(badPin.code, 'INVALID_CREDENTIALS');

    const ok1 = invokePost({ action: 'verifyTeacherPasscode', teacherId: 'TCH-01', pin: '4829' });
    assert.strictEqual(ok1.success, true);
    assert.strictEqual(ok1.data.teacherId, 'TCH-01');
    assert.strictEqual(ok1.data.teacherName, 'Mr. Ibrahim Didi');
    teacher1Token = ok1.data.token;

    const ok2 = invokePost({ action: 'verifyTeacherPasscode', teacherId: 'TCH-02', pin: '4829' });
    assert.strictEqual(ok2.success, true);
    teacher2Token = ok2.data.token;

    const verifyRes = invokeGet({ action: 'verifyTeacherToken', token: teacher1Token });
    assert.strictEqual(verifyRes.success, true);
    assert.strictEqual(verifyRes.data.teacherId, 'TCH-01');
  });

  // 3. Expired & Tampered Tokens + Role-Based Access Restrictions
  console.log('\n3. Token Integrity & Role-Based Access Control');

  await runTest('Rejects missing, tampered, and expired tokens across protected endpoints', () => {
    const missing = invokeGet({ action: 'getStudents' });
    assert.strictEqual(missing.success, false);
    assert.strictEqual(missing.code, 'AUTH_REQUIRED');

    const tampered = teacher1Token.slice(0, -4) + 'ffff';
    const tamperedRes = invokeGet({ action: 'verifyTeacherToken', token: tampered });
    assert.strictEqual(tamperedRes.success, false);
    assert.strictEqual(tamperedRes.code, 'INVALID_TOKEN');

    const secret = gasEnv.scriptProps.get('TEACHER_TOKEN_SECRET');
    const expiredClaims = {
      typ: 'SMT_TEACHER',
      role: 'teacher',
      teacherId: 'TCH-01',
      teacherName: 'Mr. Ibrahim Didi',
      iat: Date.now() - 7200000,
      exp: Date.now() - 3600000,
      jti: 'expired1'
    };
    const encoded = gasEnv.Utilities.base64EncodeWebSafe(JSON.stringify(expiredClaims)).replace(/=+$/, '');
    const sigBytes = gasEnv.Utilities.computeHmacSha256Signature(encoded, secret);
    const sigHex = sigBytes.map((b) => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
    const expiredToken = `${encoded}.${sigHex}`;

    const expRes = invokeGet({ action: 'verifyTeacherToken', token: expiredToken });
    assert.strictEqual(expRes.success, false);
    assert.strictEqual(expRes.code, 'TOKEN_EXPIRED');
  });

  await runTest('Enforces least-privilege role restrictions (Principal, Leading Teacher, and Teacher cannot perform Admin mutations)', () => {
    const pAddStu = invokePost({
      action: 'addStudent',
      token: principalToken,
      studentId: 'STU-999',
      studentName: 'Unauthorized',
      classId: 'CLS-10A'
    });
    assert.strictEqual(pAddStu.success, false);
    assert.strictEqual(pAddStu.code, 'FORBIDDEN');

    const ltChangePass = invokePost({
      action: 'changePasscode',
      token: leadingTeacherToken,
      targetRole: 'admin',
      newPasscode: 'HackedPass123',
      currentAdminPasscode: 'AdminSecret#2026'
    });
    assert.strictEqual(ltChangePass.success, false);
    assert.strictEqual(ltChangePass.code, 'FORBIDDEN');

    const pChangePass = invokePost({
      action: 'changePasscode',
      token: principalToken,
      targetRole: 'admin',
      newPasscode: 'HackedPass123',
      currentAdminPasscode: 'AdminSecret#2026'
    });
    assert.strictEqual(pChangePass.success, false);
    assert.strictEqual(pChangePass.code, 'FORBIDDEN');

    const tManageDest = invokePost({
      action: 'manageDestinations',
      token: teacher1Token,
      operation: 'add',
      destination: 'Unauthorized Room'
    });
    assert.strictEqual(tManageDest.success, false);
    assert.strictEqual(tManageDest.code, 'FORBIDDEN');
  });

  // 4. Student, Class, Teacher, Leading Teacher & Assignment CRUD
  console.log('\n4. Student, Class, Teacher & Leading Teacher CRUD');

  await runTest('Student CRUD: adds students, prevents duplicate StudentIDs, prevents changing StudentID', () => {
    const s1 = invokePost({
      action: 'addStudent',
      token: adminToken,
      studentId: 'STU-101',
      studentName: 'Ahmed Riyaz',
      classId: 'CLS-10A'
    });
    assert.strictEqual(s1.success, true);

    const s2 = invokePost({
      action: 'addStudent',
      token: adminToken,
      studentId: 'STU-102',
      studentName: 'Aminath Shifana',
      classId: 'CLS-10B'
    });
    assert.strictEqual(s2.success, true);

    const dup = invokePost({
      action: 'addStudent',
      token: adminToken,
      studentId: 'stu-101',
      studentName: 'Duplicate Student',
      classId: 'CLS-10A'
    });
    assert.strictEqual(dup.success, false);
    assert.strictEqual(dup.code, 'DUPLICATE_STUDENT_ID');

    const immutableAttempt = invokePost({
      action: 'updateStudent',
      token: adminToken,
      studentId: 'STU-101',
      newStudentId: 'STU-999',
      studentName: 'Ahmed Riyaz Updated'
    });
    assert.strictEqual(immutableAttempt.success, false);
    assert.strictEqual(immutableAttempt.code, 'IMMUTABLE_STUDENT_ID');

    const upd = invokePost({
      action: 'updateStudent',
      token: adminToken,
      studentId: 'STU-101',
      studentName: 'Ahmed Riyaz Hussain'
    });
    assert.strictEqual(upd.success, true);
    assert.strictEqual(upd.data.student.StudentName, 'Ahmed Riyaz Hussain');
  });

  await runTest('Class CRUD: prevents case-insensitive duplicate class names and preserves student links on rename', () => {
    const dupClass = invokePost({
      action: 'addClass',
      token: adminToken,
      className: 'grade 10a'
    });
    assert.strictEqual(dupClass.success, false);
    assert.strictEqual(dupClass.code, 'DUPLICATE_CLASS_NAME');

    const renameRes = invokePost({
      action: 'updateClass',
      token: adminToken,
      classId: 'CLS-10A',
      className: 'Grade 10-A (Science)'
    });
    assert.strictEqual(renameRes.success, true);

    const stuList = invokeGet({ action: 'getStudents', token: adminToken, classId: 'CLS-10A' });
    assert.strictEqual(stuList.success, true);
    assert.strictEqual(stuList.data.students.length, 1);
    assert.strictEqual(stuList.data.students[0].ClassName, 'Grade 10-A (Science)');
  });

  await runTest('Teaching Assignments & Leading Teacher Assignments prevent duplicate/conflicting active records', () => {
    const ta1 = invokePost({
      action: 'addTeachingAssignment',
      token: adminToken,
      teacherId: 'TCH-01',
      classId: 'CLS-10A'
    });
    assert.strictEqual(ta1.success, true);

    const taDup = invokePost({
      action: 'addTeachingAssignment',
      token: adminToken,
      teacherId: 'TCH-01',
      classId: 'CLS-10A'
    });
    assert.strictEqual(taDup.success, false);
    assert.strictEqual(taDup.code, 'DUPLICATE_ASSIGNMENT');

    const lt2 = invokePost({
      action: 'addLeadingTeacher',
      token: adminToken,
      leadingTeacherId: 'LT-02',
      leadingTeacherName: 'Mr. Hassan Manik'
    });
    assert.strictEqual(lt2.success, true);

    const ltaConflict = invokePost({
      action: 'addLeadingTeacherAssignment',
      token: adminToken,
      leadingTeacherId: 'LT-02',
      classId: 'CLS-10A',
      period: '2026 term 1'
    });
    assert.strictEqual(ltaConflict.success, false);
    assert.strictEqual(ltaConflict.code, 'DUPLICATE_ASSIGNMENT');
  });

  // 5. Bulk Student Import Validation & Execution
  console.log('\n5. Bulk CSV / Tabular Student Import');

  await runTest('Client & Server Bulk Import: supports column aliases, detects sheet & batch duplicates, leaves Movement_Log untouched', () => {
    const csvText = [
      'admission_no,full_name,grade',
      'STU-101,Existing Student,Grade 10-A (Science)',
      'STU-201,Fathimath Nahuza,Grade 10-A (Science)',
      'STU-202,Hussain Shareef,CLS-10B',
      'STU-201,Batch Duplicate ID,CLS-10B',
      'STU-203,Invalid Class Student,NonExistentClass'
    ].join('\n');

    const classesRes = invokeGet({ action: 'getClasses', token: adminToken });
    const studentsRes = invokeGet({ action: 'getStudents', token: adminToken, includeInactive: true });

    const preview = SMT_UTILS.parseAndValidateStudentImport(
      csvText,
      classesRes.data.classes,
      studentsRes.data.students
    );

    assert.strictEqual(preview.valid.length, 2, 'Should have 2 valid rows (STU-201, STU-202)');
    assert.strictEqual(preview.skipped.length, 1, 'Should skip STU-101 already in sheet');
    assert.strictEqual(preview.invalid.length, 2, 'Should flag duplicate STU-201 in batch and unknown class');

    const bulkRes = invokePost({
      action: 'bulkAddStudents',
      token: adminToken,
      students: [
        { studentId: 'STU-101', studentName: 'Should Skip', classId: 'CLS-10A' },
        { studentId: 'STU-201', studentName: 'Fathimath Nahuza', classId: 'CLS-10A' },
        { studentId: 'STU-202', studentName: 'Hussain Shareef', classId: 'CLS-10B' },
        { studentId: 'STU-201', studentName: 'Duplicate In Batch', classId: 'CLS-10B' },
        { studentId: 'STU-203', studentName: 'Bad Class', classId: 'CLS-NOPE' }
      ]
    });
    assert.strictEqual(bulkRes.success, true);
    assert.strictEqual(bulkRes.data.addedCount, 2);
    assert.strictEqual(bulkRes.data.skippedCount, 1);
    assert.strictEqual(bulkRes.data.invalidCount, 2);
  });

  // 6. Movement OUT & RETURN Operations, Identity Binding, and Concurrency
  console.log('\n6. Movement Operations (OUT & RETURN), Identity Binding & Locks');

  let activeMovement1Id = '';

  await runTest('Teacher Token Identity Binding: recordOut rejects spoofed teacherId and records authenticated teacher', () => {
    const spoofAttempt = invokePost({
      action: 'recordOut',
      token: teacher1Token,
      teacherId: 'TCH-02',
      studentId: 'STU-101',
      classId: 'CLS-10A',
      destination: 'Washroom'
    });
    assert.strictEqual(spoofAttempt.success, false);
    assert.strictEqual(spoofAttempt.code, 'IDENTITY_MISMATCH');

    const locksBefore = gasEnv.LockService._getStats().lockAcquireCount;
    const out1 = invokePost({
      action: 'recordOut',
      token: teacher1Token,
      studentId: 'STU-101',
      classId: 'CLS-10A',
      destination: 'Washroom',
      notes: 'Quick pass'
    });
    const locksAfter = gasEnv.LockService._getStats().lockAcquireCount;
    assert.strictEqual(out1.success, true);
    assert.ok(locksAfter > locksBefore, 'LockService must be acquired during recordOut');
    assert.strictEqual(out1.data.movement.TeacherID, 'TCH-01');
    assert.strictEqual(out1.data.movement.TeacherName, 'Mr. Ibrahim Didi');
    assert.ok(out1.data.movement.MovementID.startsWith('MOV-'));
    activeMovement1Id = out1.data.movement.MovementID;
  });

  await runTest('Prevents duplicate active OUT movements for the same student', () => {
    const dupOut = invokePost({
      action: 'recordOut',
      token: teacher1Token,
      studentId: 'STU-101',
      classId: 'CLS-10A',
      destination: 'Library'
    });
    assert.strictEqual(dupOut.success, false);
    assert.strictEqual(dupOut.code, 'DUPLICATE_ACTIVE_MOVEMENT');
  });

  await runTest('Blocks deactivating a student, class, or teacher with an unresolved OUT movement', () => {
    const deactStu = invokePost({
      action: 'deactivateStudent',
      token: adminToken,
      studentId: 'STU-101',
      status: 'Inactive'
    });
    assert.strictEqual(deactStu.success, false);
    assert.strictEqual(deactStu.code, 'UNRESOLVED_MOVEMENT');

    const deactTch = invokePost({
      action: 'deactivateTeacher',
      token: adminToken,
      teacherId: 'TCH-01',
      status: 'Inactive'
    });
    assert.strictEqual(deactTch.success, false);
    assert.strictEqual(deactTch.code, 'UNRESOLVED_MOVEMENT');

    const deactCls = invokePost({
      action: 'deactivateClass',
      token: adminToken,
      classId: 'CLS-10A',
      status: 'Inactive'
    });
    assert.strictEqual(deactCls.success, false);
    assert.strictEqual(deactCls.code, 'UNRESOLVED_MOVEMENT');
  });

  await runTest('Leading Teacher scope restriction: Leading Teacher only sees movements for assigned classes (CLS-10A, not CLS-10B)', () => {
    const out2 = invokePost({
      action: 'recordOut',
      token: teacher2Token,
      studentId: 'STU-102',
      classId: 'CLS-10B',
      destination: 'Science Laboratory'
    });
    assert.strictEqual(out2.success, true);

    const pOut = invokeGet({ action: 'getCurrentOut', token: principalToken });
    assert.strictEqual(pOut.success, true);
    assert.strictEqual(pOut.data.records.length, 2);

    const ltOut = invokeGet({ action: 'getCurrentOut', token: leadingTeacherToken });
    assert.strictEqual(ltOut.success, true);
    assert.strictEqual(ltOut.data.records.length, 1);
    assert.strictEqual(ltOut.data.records[0].ClassID, 'CLS-10A');
    assert.strictEqual(ltOut.data.records[0].StudentID, 'STU-101');
  });

  await runTest('recordReturn prevents returning the wrong student movement and preserves TimeOut when recording TimeIn', () => {
    const wrongStudent = invokePost({
      action: 'recordReturn',
      token: teacher1Token,
      movementId: activeMovement1Id,
      studentId: 'STU-102'
    });
    assert.strictEqual(wrongStudent.success, false);
    assert.strictEqual(wrongStudent.code, 'WRONG_STUDENT_MOVEMENT');

    const validReturn = invokePost({
      action: 'recordReturn',
      token: teacher1Token,
      movementId: activeMovement1Id,
      studentId: 'STU-101'
    });
    assert.strictEqual(validReturn.success, true);
    assert.strictEqual(validReturn.data.movement.Status, 'RETURNED');
    assert.ok(validReturn.data.movement.TimeOut);
    assert.ok(validReturn.data.movement.TimeIn);

    const repeatReturn = invokePost({
      action: 'recordReturn',
      token: teacher1Token,
      movementId: activeMovement1Id
    });
    assert.strictEqual(repeatReturn.success, false);
    assert.strictEqual(repeatReturn.code, 'ALREADY_RETURNED');
  });

  // 7. Timestamp Parsing (1899 Protection) & Attention Thresholds
  console.log('\n7. Timestamp Parsing (1899 Protection) & Attention Thresholds');

  await runTest('Never calculates duration from 1899 when Google Sheets returns a time-only Date object', () => {
    const nowMs = Date.now();
    const twelveMinsAgoMs = nowMs - (12 * 60 * 1000 + 15 * 1000);
    const movId = Code.generateMovementId_(twelveMinsAgoMs);
    const corrupted1899Date = new Date(Date.UTC(1899, 11, 30, 10, 15, 0));

    const serverMetrics = Code.computeMovementMetrics_(corrupted1899Date, null, movId, 10, 15, nowMs);
    assert.strictEqual(serverMetrics.elapsedSeconds, 12 * 60 + 15);
    assert.strictEqual(serverMetrics.formattedDuration, '12m 15s');
    assert.strictEqual(serverMetrics.attentionLevel, 'Warning');

    const clientMetrics = SMT_UTILS.calculateMovementStatus(corrupted1899Date, null, movId, 10, 15, nowMs);
    assert.strictEqual(clientMetrics.elapsedSeconds, 12 * 60 + 15);
    assert.strictEqual(clientMetrics.formattedDuration, '12m 15s');
    assert.strictEqual(clientMetrics.attentionLevel, 'Warning');
  });

  await runTest('Correctly classifies Normal (<10m), Warning (10m to <15m), and Critical (>=15m) attention levels', () => {
    const refNow = 1793337666000;
    const mNormal = SMT_UTILS.calculateMovementStatus(
      new Date(refNow - 599 * 1000).toISOString(),
      null,
      'MOV-' + (refNow - 599 * 1000) + '-AA11BB',
      10,
      15,
      refNow
    );
    assert.strictEqual(mNormal.attentionLevel, 'Normal');
    assert.strictEqual(mNormal.formattedDuration, '9m 59s');

    const mWarn = SMT_UTILS.calculateMovementStatus(
      new Date(refNow - 600 * 1000).toISOString(),
      null,
      'MOV-' + (refNow - 600 * 1000) + '-CC22DD',
      10,
      15,
      refNow
    );
    assert.strictEqual(mWarn.attentionLevel, 'Warning');
    assert.strictEqual(mWarn.formattedDuration, '10m 00s');

    const mWarn2 = SMT_UTILS.calculateMovementStatus(
      new Date(refNow - 899 * 1000).toISOString(),
      null,
      'MOV-' + (refNow - 899 * 1000) + '-EE33FF',
      10,
      15,
      refNow
    );
    assert.strictEqual(mWarn2.attentionLevel, 'Warning');

    const mCrit = SMT_UTILS.calculateMovementStatus(
      new Date(refNow - 900 * 1000).toISOString(),
      null,
      'MOV-' + (refNow - 900 * 1000) + '-112233',
      10,
      15,
      refNow
    );
    assert.strictEqual(mCrit.attentionLevel, 'Critical');
    assert.strictEqual(mCrit.formattedDuration, '15m 00s');
  });

  // 8. Reports, Repeated Student Analytics & CSV Export
  console.log('\n8. Movement Reports, Analytics & CSV Export');

  await runTest('Generates accurate movement analytics, repeated student counts, and formula-safe CSV export', () => {
    const secondOut = invokePost({
      action: 'recordOut',
      token: teacher1Token,
      studentId: 'STU-101',
      classId: 'CLS-10A',
      destination: 'Library'
    });
    assert.strictEqual(secondOut.success, true);

    const reportRes = invokeGet({
      action: 'getMovementReports',
      token: adminToken,
      dateRange: 'today'
    });
    assert.strictEqual(reportRes.success, true);
    assert.strictEqual(reportRes.data.summary.totalMovements, 3);
    assert.strictEqual(reportRes.data.summary.currentOutCount, 2);
    assert.strictEqual(reportRes.data.summary.returnedCount, 1);
    assert.strictEqual(reportRes.data.repeatedStudents.length, 1);
    assert.strictEqual(reportRes.data.repeatedStudents[0].StudentID, 'STU-101');
    assert.strictEqual(reportRes.data.repeatedStudents[0].movementCount, 2);

    const csv = SMT_UTILS.movementRecordsToCsv(reportRes.data.records);
    assert.ok(csv.includes('MovementID,StudentID,StudentName,ClassID,ClassName'));
    assert.ok(csv.includes('STU-101'));
    assert.ok(csv.includes('STU-102'));
  });

  // 9. Passcode Changes, Session Invalidation & Secret Sanitization
  console.log('\n9. Passcode Management, Secret Rotation & Sanitization');

  await runTest('Public settings and security metadata never leak hashes, PINs, salts, or signing secrets', () => {
    const pubSettings = invokeGet({ action: 'getSettings' });
    assert.strictEqual(pubSettings.success, true);
    const jsonStr = JSON.stringify(pubSettings);
    assert.strictEqual(/sha256\$|AUTH_HASH|SECRET|4829|AdminSecret/i.test(jsonStr), false);

    const secMeta = invokeGet({ action: 'getSecuritySettings', token: adminToken });
    assert.strictEqual(secMeta.success, true);
    const secStr = JSON.stringify(secMeta);
    assert.strictEqual(/sha256\$|SECRET|4829|AdminSecret/i.test(secStr), false);
    assert.strictEqual(secMeta.data.credentialsConfigured.admin, true);
    assert.strictEqual(secMeta.data.credentialsConfigured.teacher_pin, true);
  });

  await runTest('Changing Teacher App PIN requires current Admin passcode and invalidates active Teacher tokens immediately', () => {
    const badAuth = invokePost({
      action: 'changePasscode',
      token: adminToken,
      targetRole: 'teacher_pin',
      newPasscode: '9512',
      currentAdminPasscode: 'WrongAdminPass'
    });
    assert.strictEqual(badAuth.success, false);
    assert.strictEqual(badAuth.code, 'INVALID_ADMIN_PASSCODE');

    const rotateRes = invokePost({
      action: 'changePasscode',
      token: adminToken,
      targetRole: 'teacher_pin',
      newPasscode: '9512',
      currentAdminPasscode: 'AdminSecret#2026'
    });
    assert.strictEqual(rotateRes.success, true);
    assert.strictEqual(rotateRes.data.sessionsInvalidated, true);

    const oldTokenCheck = invokeGet({ action: 'verifyTeacherToken', token: teacher1Token });
    assert.strictEqual(oldTokenCheck.success, false);
    assert.strictEqual(oldTokenCheck.code, 'INVALID_TOKEN');

    const oldPinLogin = invokePost({ action: 'verifyTeacherPasscode', teacherId: 'TCH-01', pin: '4829' });
    assert.strictEqual(oldPinLogin.success, false);

    const newPinLogin = invokePost({ action: 'verifyTeacherPasscode', teacherId: 'TCH-01', pin: '9512' });
    assert.strictEqual(newPinLogin.success, true);
    teacher1Token = newPinLogin.data.token;
  });

  await runTest('Auto-migrates any plaintext credential row in Settings sheet into a salted SHA-256 hash', () => {
    const settingsSheet = gasEnv.spreadsheet.getSheetByName('Settings');
    const nextRow = settingsSheet.getLastRow() + 1;
    settingsSheet.getRange(nextRow, 1, 1, 2).setValues([['principalPasscode', 'PlaintextPrincipal#99']]);

    invokeGet({ action: 'getSettings' });

    const allSettings = settingsSheet.getRange(1, 1, settingsSheet.getLastRow(), 2).getValues();
    for (const [k, v] of allSettings) {
      assert.notStrictEqual(v, 'PlaintextPrincipal#99');
      if (k === 'AUTH_HASH_PRINCIPAL') {
        assert.ok(String(v).startsWith('sha256$'));
      }
    }

    const pLogin = invokePost({ action: 'verifyPasscode', role: 'principal', passcode: 'PlaintextPrincipal#99' });
    assert.strictEqual(pLogin.success, true);
  });

  // 10. Malformed Request Handling, Local Server & Frontend Regression Verification
  console.log('\n10. Error Handling, Local Server & Frontend Regression Checks');

  await runTest('Handles malformed JSON requests and unknown actions gracefully without exposing stack traces', () => {
    const badJsonOut = Code.doPost({
      parameter: {},
      postData: { contents: '{ malformed json ' }
    });
    const badJson = JSON.parse(badJsonOut.getContent());
    assert.strictEqual(badJson.success, false);
    assert.strictEqual(badJson.code, 'MALFORMED_JSON');

    const unknownAct = invokeGet({ action: 'nonExistentAction' });
    assert.strictEqual(unknownAct.success, false);
    assert.strictEqual(unknownAct.code, 'UNKNOWN_ACTION');
  });

  await runTest('Central configuration defaults to production Web App URL and supports intentional local-dev override', () => {
    SMT_CONFIG.resetToDefault();
    const prodUrl = 'https://script.google.com/macros/s/AKfycbygZM_HFzRnom0e3AFk4pSyvHufSLyF0hCpPba9oSr9wkyGXQrT-8SsMTU3MoLI-vtB/exec';
    assert.strictEqual(SMT_CONFIG.getWebAppUrl(), prodUrl);
    assert.strictEqual(SMT_CONFIG.isUsingProduction(), true);

    // Intentional local development override
    SMT_CONFIG.setWebAppUrl('http://localhost:3000/api');
    assert.strictEqual(SMT_CONFIG.getWebAppUrl(), 'http://localhost:3000/api');
    assert.strictEqual(SMT_CONFIG.isUsingLocalOverride(), true);
    assert.strictEqual(SMT_CONFIG.isUsingProduction(), false);

    // Reset back to default production
    SMT_CONFIG.resetToDefault();
    assert.strictEqual(SMT_CONFIG.getWebAppUrl(), prodUrl);
    assert.strictEqual(SMT_CONFIG.isUsingProduction(), true);
  });

  await runTest('Shared API client (js/api.js) never reports Connected to Google Sheets before a successful backend response', async () => {
    SMT_CONFIG.setWebAppUrl('');
    assert.notStrictEqual(SMT_API.getConnectionState().status, 'connected');
    await assert.rejects(
      async () => {
        await SMT_API.get('getSettings');
      },
      (err) => err.code === 'UNCONFIGURED_URL'
    );
    assert.strictEqual(SMT_API.getConnectionState().status, 'unconfigured');
  });

  await runTest('Local development HTTP server serves all 4 dashboards and assets with HTTP 200', async () => {
    const srv = createServer();
    await new Promise((resolve) => srv.listen(0, '127.0.0.1', resolve));
    const port = srv.address().port;

    const fetchStatus = (urlPath) => new Promise((resolve, reject) => {
      http.get(`http://127.0.0.1:${port}${urlPath}`, (res) => {
        let body = '';
        res.on('data', (chunk) => { body += chunk; });
        res.on('end', () => resolve({ statusCode: res.statusCode, body }));
      }).on('error', reject);
    });

    try {
      for (const page of ['/index.html', '/principal-dashboard.html', '/leading-teacher-dashboard.html', '/admin-dashboard.html', '/css/styles.css', '/js/api.js']) {
        const res = await fetchStatus(page);
        assert.strictEqual(res.statusCode, 200, `Expected HTTP 200 for ${page}`);
        assert.ok(res.body.length > 100);
      }
    } finally {
      await new Promise((resolve) => srv.close(resolve));
    }
  });

  console.log('\n============================================================');
  console.log(`Test Suite Summary: ${passedTests}/${totalTests} passed, ${failedTests} failed`);
  console.log('============================================================');

  if (failedTests > 0) {
    process.exit(1);
  }
}

main();
