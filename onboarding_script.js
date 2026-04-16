// ============================================================
// ONBOARDING PROCESSOR — Google Apps Script
// Spreadsheet: 1B0JjqH3R0f2VEYk9EZchjvuPtdhd7SSpHuu3iSANdfg
// Naming conventions: MVP Naming Conventions & ID Rules v2
// ============================================================

var SPREADSHEET_ID = '1B0JjqH3R0f2VEYk9EZchjvuPtdhd7SSpHuu3iSANdfg';
var DATA_START_ROW = 3; // Row 1 = title, Row 2 = headers, Row 3+ = data
var PREPOSITIONS = ['de', 'da', 'do', 'dos', 'das'];

// Course code mapping (platform name variants → standard code)
var COURSE_MAP = {
  'coding i': 'COD1', 'coding 1': 'COD1', 'coding1': 'COD1',
  'int. tec. e coding i': 'COD1', 'inteligencia tecnologica e coding i': 'COD1',
  'inteligência tecnológica e coding i': 'COD1',
  'coding ii': 'COD2', 'coding 2': 'COD2', 'coding2': 'COD2',
  'int. tec. e coding ii': 'COD2', 'inteligencia tecnologica e coding ii': 'COD2',
  'inteligência tecnológica e coding ii': 'COD2',
  'financeira': 'FIN3', 'inteligencia financeira': 'FIN3',
  'inteligência financeira': 'FIN3', 'inteligencia financeira e speech': 'FIN3',
  'inteligência financeira e comunicação': 'FIN3',
  'empreendedorismo': 'EMP4', 'soft skills': 'EMP4',
  'soft skills carreira e empreendedorismo': 'EMP4',
  'inteligência carreira e empreendedorismo': 'EMP4',
  'kindergarten': 'KG', 'kids': 'KID', 'pre teens': 'PT',
  'teens': 'TN', 'young adults': 'YA'
};

// -------------------- MENU --------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('\uD83D\uDE80 Onboarding')
    .addItem('Process new entries', 'processOnboarding')
    .addToUi();
}

// -------------------- DOPOST --------------------

function doPost(e) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('onboarding');
  var data = JSON.parse(e.postData.contents);
  var rows = data.rows;
  var timestamp = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
  rows.forEach(function (row) {
    sheet.appendRow([timestamp].concat(row));
  });
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// -------------------- MAIN PROCESSOR --------------------

function processOnboarding() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var ui = SpreadsheetApp.getUi();

  var onbSheet = ss.getSheetByName('onboarding');
  var schoolsSheet = ss.getSheetByName('schools');
  var teachersSheet = ss.getSheetByName('teachers');
  var groupsSheet = ss.getSheetByName('groups');
  var studentsSheet = ss.getSheetByName('students');

  // Ensure column P header exists
  var headerP = onbSheet.getRange(2, 16).getValue();
  if (headerP !== 'Processed') {
    onbSheet.getRange(2, 16).setValue('Processed');
  }

  var lastRow = onbSheet.getLastRow();
  if (lastRow < 3) {
    ui.alert('No data to process.');
    return;
  }
  var onbData = onbSheet.getRange(3, 1, lastRow - 2, 16).getValues();

  // Filter unprocessed rows
  var unprocessedIndices = [];
  var unprocessedRows = [];
  for (var i = 0; i < onbData.length; i++) {
    if (!onbData[i][15] || onbData[i][15] === '') {
      unprocessedIndices.push(i);
      unprocessedRows.push(onbData[i]);
    }
  }

  if (unprocessedRows.length === 0) {
    ui.alert('No new entries to process.');
    return;
  }

  // Load existing data
  var schoolsData = loadSheetData(schoolsSheet, 10);
  var teachersData = loadSheetData(teachersSheet, 8);
  var groupsData = loadSheetData(groupsSheet, 9);
  var studentsData = loadSheetData(studentsSheet, 8);

  // Build lookup maps
  var schoolMap = buildSchoolMap(schoolsData);
  var teacherMap = buildTeacherMap(teachersData);
  var groupMap = buildGroupMap(groupsData);
  var studentMap = buildStudentMap(studentsData);

  // Collect existing IDs
  var existingSchoolIDs = collectColumn(schoolsData, 0);
  var existingTeacherIDs = collectColumn(teachersData, 0);
  var existingGroupIDs = collectColumn(groupsData, 0);

  var stats = { schools: 0, teachers: 0, groups: 0, students: 0 };
  var newSchoolRows = [];
  var newTeacherRows = [];
  var newGroupRows = [];
  var newStudentRows = [];

  // Track relationships
  var schoolGroupIDs = {};
  var schoolTeacherIDs = {};
  var teacherGroupIDs = {};

  for (var s = 0; s < schoolsData.length; s++) {
    var sid = schoolsData[s][0];
    if (sid) {
      schoolGroupIDs[sid] = parseCSVSet(schoolsData[s][7]);
      schoolTeacherIDs[sid] = parseCSVSet(schoolsData[s][8]);
    }
  }
  for (var t = 0; t < teachersData.length; t++) {
    var tid = teachersData[t][0];
    if (tid) {
      teacherGroupIDs[tid] = parseCSVSet(teachersData[t][6]);
    }
  }

  var groupStudentCount = {};
  var now = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });

  // ---------- PASS 1: Schools & Teachers ----------
  for (var r = 0; r < unprocessedRows.length; r++) {
    var row = unprocessedRows[r];
    var schoolName = trim(row[1]);
    var coordName = trim(row[2]);
    var coordWA = trim(row[3]);
    var coordEmail = trim(row[4]);
    var teacherName = trim(row[5]);
    var teacherWA = trim(row[6]);
    var teacherEmail = trim(row[7]);

    if (schoolName && !schoolMap[schoolName]) {
      var schoolID = generateSchoolID(schoolName, existingSchoolIDs);
      existingSchoolIDs.push(schoolID);
      var schoolRow = [schoolID, schoolName, coordName, coordEmail, coordWA, '', '', '', '', now];
      newSchoolRows.push(schoolRow);
      schoolMap[schoolName] = schoolRow;
      schoolGroupIDs[schoolID] = {};
      schoolTeacherIDs[schoolID] = {};
      stats.schools++;
    }

    if (teacherName && !teacherMap[teacherName]) {
      var currentSchoolID = getSchoolID(schoolName, schoolMap);
      var teacherID = generateTeacherID(teacherName, existingTeacherIDs);
      existingTeacherIDs.push(teacherID);
      var teacherRow = [teacherID, teacherName, teacherWA, '', '', currentSchoolID, '', now];
      newTeacherRows.push(teacherRow);
      teacherMap[teacherName] = teacherRow;
      teacherGroupIDs[teacherID] = {};
      stats.teachers++;
    }
  }

  // ---------- PASS 2: Groups ----------
  for (var r = 0; r < unprocessedRows.length; r++) {
    var row = unprocessedRows[r];
    var schoolName = trim(row[1]);
    var teacherName = trim(row[5]);
    var groupName = trim(row[8]);
    var program = trim(row[9]);
    var timetable = trim(row[10]);
    var project = trim(row[11]);

    if (!groupName || !schoolName) continue;

    var schoolID = getSchoolID(schoolName, schoolMap);
    // Extract just the school code (without the 3-digit number) for group ID prefix
    var schoolCode = getSchoolCode(schoolID);
    var groupKey = schoolID + '|' + groupName;

    if (!groupMap[groupKey]) {
      var groupID = generateGroupID(groupName, schoolCode, existingGroupIDs);
      existingGroupIDs.push(groupID);
      var teacherID = getTeacherID(teacherName, teacherMap);
      var groupRow = [groupID, groupName, 0, program, project, timetable, schoolID, teacherID, now];
      newGroupRows.push(groupRow);
      groupMap[groupKey] = groupRow;
      groupStudentCount[groupID] = 0;
      stats.groups++;

      if (schoolID && schoolGroupIDs[schoolID]) schoolGroupIDs[schoolID][groupID] = true;
      if (schoolID && schoolTeacherIDs[schoolID] && teacherID) schoolTeacherIDs[schoolID][teacherID] = true;
      if (teacherID && teacherGroupIDs[teacherID]) teacherGroupIDs[teacherID][groupID] = true;
    } else {
      var existingGroupID = groupMap[groupKey][0];
      var teacherID = getTeacherID(teacherName, teacherMap);
      if (schoolID && schoolGroupIDs[schoolID]) schoolGroupIDs[schoolID][existingGroupID] = true;
      if (schoolID && schoolTeacherIDs[schoolID] && teacherID) schoolTeacherIDs[schoolID][teacherID] = true;
      if (teacherID && teacherGroupIDs[teacherID]) teacherGroupIDs[teacherID][existingGroupID] = true;
    }
  }

  // ---------- PASS 3: Students ----------
  var groupStudentIDs = {};
  for (var st = 0; st < studentsData.length; st++) {
    var stGid = studentsData[st][5];
    var stId = studentsData[st][0];
    if (stGid) {
      if (!groupStudentIDs[stGid]) groupStudentIDs[stGid] = [];
      groupStudentIDs[stGid].push(stId);
    }
  }

  for (var r = 0; r < unprocessedRows.length; r++) {
    var row = unprocessedRows[r];
    var schoolName = trim(row[1]);
    var teacherName = trim(row[5]);
    var groupName = trim(row[8]);
    var studentName = trim(row[12]);
    var studentAge = trim(row[13]);
    var consent = trim(row[14]);

    if (!studentName || !groupName || !schoolName) continue;

    var schoolID = getSchoolID(schoolName, schoolMap);
    var groupKey = schoolID + '|' + groupName;
    var groupID = groupMap[groupKey] ? groupMap[groupKey][0] : '';
    var studentKey = groupID + '|' + studentName;

    if (!studentMap[studentKey]) {
      if (!groupStudentIDs[groupID]) groupStudentIDs[groupID] = [];
      var studentID = generateStudentID(studentName, groupStudentIDs[groupID]);
      groupStudentIDs[groupID].push(studentID);
      var teacherID = getTeacherID(teacherName, teacherMap);
      var studentRow = [studentID, studentName, studentAge, consent, schoolID, groupID, teacherID, now];
      newStudentRows.push(studentRow);
      studentMap[studentKey] = studentRow;

      if (!groupStudentCount[groupID] && groupStudentCount[groupID] !== 0) {
        groupStudentCount[groupID] = (groupMap[groupKey] ? (parseInt(groupMap[groupKey][2], 10) || 0) : 0);
      }
      groupStudentCount[groupID]++;
      stats.students++;
    }
  }

  // ---------- WRITE NEW ROWS ----------
  if (newSchoolRows.length > 0) {
    var sLastRow = Math.max(schoolsSheet.getLastRow(), 2);
    schoolsSheet.getRange(sLastRow + 1, 1, newSchoolRows.length, 10).setValues(newSchoolRows);
  }
  if (newTeacherRows.length > 0) {
    var tLastRow = Math.max(teachersSheet.getLastRow(), 2);
    teachersSheet.getRange(tLastRow + 1, 1, newTeacherRows.length, 8).setValues(newTeacherRows);
  }
  if (newGroupRows.length > 0) {
    var gLastRow = Math.max(groupsSheet.getLastRow(), 2);
    groupsSheet.getRange(gLastRow + 1, 1, newGroupRows.length, 9).setValues(newGroupRows);
  }
  if (newStudentRows.length > 0) {
    var stLastRow = Math.max(studentsSheet.getLastRow(), 2);
    studentsSheet.getRange(stLastRow + 1, 1, newStudentRows.length, 8).setValues(newStudentRows);
  }

  // ---------- UPDATE RELATIONSHIPS ----------
  updateSchoolRelationships(schoolsSheet, schoolGroupIDs, schoolTeacherIDs);
  updateTeacherRelationships(teachersSheet, teacherGroupIDs);
  updateGroupStudentCounts(groupsSheet, groupStudentCount);

  // ---------- MARK PROCESSED ----------
  var processedTimestamp = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
  for (var p = 0; p < unprocessedIndices.length; p++) {
    var actualRow = unprocessedIndices[p] + DATA_START_ROW;
    onbSheet.getRange(actualRow, 16).setValue(processedTimestamp);
  }

  // ---------- SUMMARY ----------
  var msg = 'Processing complete!\n\n' +
    'New schools: ' + stats.schools + '\n' +
    'New teachers: ' + stats.teachers + '\n' +
    'New groups: ' + stats.groups + '\n' +
    'New students: ' + stats.students + '\n' +
    'Rows processed: ' + unprocessedRows.length;
  ui.alert(msg);
}

// ============================================================
// HELPER FUNCTIONS
// ============================================================

function trim(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim();
}

function stripAccents(str) {
  var accents = 'ÀÁÂÃÄÅàáâãäåÈÉÊËèéêëÌÍÎÏìíîïÒÓÔÕÖØòóôõöøÙÚÛÜùúûüÇçÑñÝýÿ';
  var plain   = 'AAAAAAaaaaaaEEEEeeeeIIIIiiiiOOOOOOooooooUUUUuuuuCcNnYyy';
  var result = '';
  for (var i = 0; i < str.length; i++) {
    var idx = accents.indexOf(str[i]);
    result += (idx >= 0) ? plain[idx] : str[i];
  }
  return result;
}

function isPreposition(word) {
  return PREPOSITIONS.indexOf(word.toLowerCase()) >= 0;
}

function getInitials(text) {
  var words = text.split(/\s+/);
  var initials = '';
  for (var i = 0; i < words.length; i++) {
    var w = words[i];
    if (w && !isPreposition(w)) {
      initials += stripAccents(w.charAt(0)).toUpperCase();
    }
  }
  return initials;
}

function titleCase(str) {
  if (!str) return '';
  return str.charAt(0).toUpperCase() + str.substring(1).toLowerCase();
}

function padNumber(num, digits) {
  var s = String(num);
  while (s.length < digits) s = '0' + s;
  return s;
}

function getNextSequence(base, existingIDs) {
  var maxSeq = 0;
  for (var i = 0; i < existingIDs.length; i++) {
    var id = String(existingIDs[i]);
    if (id.indexOf(base) === 0) {
      var numPart = id.substring(base.length);
      var num = parseInt(numPart, 10);
      if (!isNaN(num) && num > maxSeq) maxSeq = num;
    }
  }
  return maxSeq + 1;
}

// ============================================================
// ID GENERATORS — follows MVP Naming Conventions v2
// ============================================================

// --- SCHOOLS ---
// Format: FRANCHISE_CODE + UNIT_CODE + 3-digit number (e.g. PEMAN001)
// Collision: extend unit code by one char at a time until unique base
function generateSchoolID(schoolName, existingIDs) {
  var parts = schoolName.split(' - ');
  var baseCode;

  if (parts.length >= 2) {
    var franchise = trim(parts[0]);
    var unit = trim(parts.slice(1).join(' - '));
    baseCode = getInitials(franchise) + getUnitCode(unit);
  } else {
    baseCode = getInitials(schoolName);
  }
  baseCode = stripAccents(baseCode).toUpperCase();

  // Check for collision with existing base codes (strip trailing digits)
  var existingBases = existingIDs.map(function(id) {
    return String(id).replace(/\d+$/, '');
  });

  // If base collides, extend unit code
  var unitText = parts.length >= 2 ? trim(parts.slice(1).join(' - ')) : schoolName;
  var unitWords = unitText.split(/\s+/).filter(function(w) { return !isPreposition(w); });
  var extendIdx = 0;

  while (hasBaseCollision(baseCode, existingBases, existingIDs)) {
    // Extend by adding more chars from unit words
    baseCode = extendCode(baseCode, unitWords, extendIdx);
    extendIdx++;
    if (extendIdx > 10) break; // safety
  }

  var seq = getNextSequence(baseCode, existingIDs);
  return baseCode + padNumber(seq, 3);
}

// Get unit code: initials of significant words, skip prepositions
function getUnitCode(unitText) {
  var words = unitText.split(/\s+/);
  var code = '';
  for (var i = 0; i < words.length; i++) {
    var w = words[i];
    if (!w) continue;
    // Check if it's a number (like "Carpina 1" → keep the number)
    if (/^\d+$/.test(w)) {
      code += w;
    } else if (!isPreposition(w)) {
      code += stripAccents(w.charAt(0)).toUpperCase();
    }
  }
  return code;
}

// Check if a base code collides (different school using same base)
function hasBaseCollision(baseCode, existingBases, existingIDs) {
  // Count how many different full IDs use this base
  var count = 0;
  for (var i = 0; i < existingBases.length; i++) {
    if (existingBases[i] === baseCode) count++;
  }
  return count > 0;
}

// Extend code by adding chars from unit words
function extendCode(currentCode, unitWords, extendIdx) {
  // Try adding more characters from the last significant word
  for (var i = unitWords.length - 1; i >= 0; i--) {
    var w = stripAccents(unitWords[i]).toUpperCase();
    if (w.length > 1) {
      var extraChar = w.charAt(Math.min(1 + extendIdx, w.length - 1));
      return currentCode + extraChar;
    }
  }
  return currentCode + 'X';
}

// Extract school code (base without trailing digits) for group ID prefix
function getSchoolCode(schoolID) {
  return String(schoolID).replace(/\d+$/, '');
}

// --- TEACHERS ---
// Format: FirLas + 3-digit number (e.g. AnaMon001)
// Title case. First 3 of first name + first 3 of last name.
// Collision: try 4+3, then 3+4, then extend further
function generateTeacherID(fullName, existingIDs) {
  var parts = fullName.split(/\s+/);
  var firstName = stripAccents(parts[0]);
  var lastName = stripAccents(parts.length > 1 ? parts[parts.length - 1] : 'X');

  // Standard: 3+3
  var base = titleCase(firstName.substring(0, 3)) + titleCase(lastName.substring(0, 3));
  var seq = getNextSequence(base, existingIDs);
  var candidate = base + padNumber(seq, 3);

  // Check if this base already exists for a DIFFERENT person
  // If seq === 1 and no existing entries, it's clean
  if (!hasTeacherBaseConflict(base, fullName, existingIDs)) {
    return base + padNumber(seq, 3);
  }

  // Collision — try 4+3
  base = titleCase(firstName.substring(0, 4)) + titleCase(lastName.substring(0, 3));
  if (!hasTeacherBaseConflict(base, fullName, existingIDs)) {
    seq = getNextSequence(base, existingIDs);
    return base + padNumber(seq, 3);
  }

  // Try 3+4
  base = titleCase(firstName.substring(0, 3)) + titleCase(lastName.substring(0, 4));
  if (!hasTeacherBaseConflict(base, fullName, existingIDs)) {
    seq = getNextSequence(base, existingIDs);
    return base + padNumber(seq, 3);
  }

  // Fallback: use 3+3 with next sequence number
  base = titleCase(firstName.substring(0, 3)) + titleCase(lastName.substring(0, 3));
  seq = getNextSequence(base, existingIDs);
  return base + padNumber(seq, 3);
}

function hasTeacherBaseConflict(base, fullName, existingIDs) {
  // A conflict means the same base code is used by a different person
  // For simplicity, we just check if the base is already in use
  for (var i = 0; i < existingIDs.length; i++) {
    if (String(existingIDs[i]).indexOf(base) === 0) return true;
  }
  return false;
}

// --- GROUPS ---
// Format: SCHOOLCODE_COURSECODE-SECTIONNUMBER (e.g. PEP_COD1-2)
// Course code mapped from platform name. Section number from group name.
function generateGroupID(groupName, schoolCode, existingGroupIDs) {
  var courseCode = mapCourseCode(groupName);
  var sectionNum = extractSectionNumber(groupName);

  var groupID = schoolCode + '_' + courseCode + '-' + sectionNum;

  // Check for collision — increment section number if needed
  while (existingGroupIDs.indexOf(groupID) >= 0) {
    sectionNum++;
    groupID = schoolCode + '_' + courseCode + '-' + sectionNum;
  }

  return groupID;
}

// Map a platform group name to a standard course code
function mapCourseCode(groupName) {
  var normalized = stripAccents(groupName).toLowerCase()
    .replace(/\s*-\s*\d+.*$/, '')  // Remove " - 01", " - 03" etc.
    .replace(/\s+\d+$/, '')         // Remove trailing numbers like "Coding 1"... wait, "Coding 1" IS the course
    .trim();

  // First try exact match with the full cleaned name
  var fullNorm = stripAccents(groupName).toLowerCase()
    .replace(/\s*-\s*\d+.*$/, '')  // Remove section number part " - 01"
    .trim();

  // Try matching against COURSE_MAP
  if (COURSE_MAP[fullNorm]) return COURSE_MAP[fullNorm];

  // Try without trailing section indicators
  var withoutSection = fullNorm.replace(/\s*[-–]\s*\d+$/, '').trim();
  if (COURSE_MAP[withoutSection]) return COURSE_MAP[withoutSection];

  // Try progressively shorter matches
  var words = fullNorm.split(/\s+/);
  for (var len = words.length; len >= 1; len--) {
    var attempt = words.slice(0, len).join(' ');
    if (COURSE_MAP[attempt]) return COURSE_MAP[attempt];
  }

  // Fallback: first 3 chars uppercase
  var fallback = stripAccents(groupName).replace(/[^a-zA-Z]/g, '').substring(0, 3).toUpperCase();
  return fallback || 'GRP';
}

// Extract section number from group name (e.g. "Coding 1 - 03" → 3)
function extractSectionNumber(groupName) {
  // Try "Name - XX" pattern first
  var dashMatch = groupName.match(/[-–]\s*(\d+)\s*$/);
  if (dashMatch) return parseInt(dashMatch[1], 10);

  // Try trailing number after course name
  var trailingMatch = groupName.match(/\d+\s*$/);
  if (trailingMatch) return parseInt(trailingMatch[0], 10);

  // Default to 1
  return 1;
}

// --- STUDENTS ---
// Format: FirL + 3-digit number (e.g. CarB001)
// First 3 of first name + initial of father's surname (last word). Title case.
// Collision: extend surname by one letter at a time, then increment number.
// Unique within group only.
function generateStudentID(fullName, groupExistingIDs) {
  var parts = fullName.split(/\s+/);
  var firstName = stripAccents(parts[0]);

  if (parts.length <= 1) {
    var base = titleCase(firstName.substring(0, 3)) + 'X';
    var seq = getNextSequence(base, groupExistingIDs);
    return base + padNumber(seq, 3);
  }

  // Father's surname = last word
  var lastName = stripAccents(parts[parts.length - 1]);

  // Standard: First3 + Last initial
  var base = titleCase(firstName.substring(0, 3)) + lastName.charAt(0).toUpperCase();
  var seq = getNextSequence(base, groupExistingIDs);
  var candidate = base + padNumber(seq, 3);

  // If seq is 1 and no collision, use it
  if (groupExistingIDs.indexOf(candidate) < 0) return candidate;

  // Collision — extend surname letter by letter
  var surnameLen = 2;
  while (surnameLen <= lastName.length) {
    base = titleCase(firstName.substring(0, 3)) + titleCase(lastName.substring(0, surnameLen));
    seq = getNextSequence(base, groupExistingIDs);
    candidate = base + padNumber(seq, 3);
    if (groupExistingIDs.indexOf(candidate) < 0) return candidate;
    surnameLen++;
  }

  // Final fallback: use first3 + full last + increment
  base = titleCase(firstName.substring(0, 3)) + titleCase(lastName);
  seq = getNextSequence(base, groupExistingIDs);
  return base + padNumber(seq, 3);
}

// ============================================================
// DATA LOADING & LOOKUP
// ============================================================

function loadSheetData(sheet, numCols) {
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return [];
  return sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, numCols).getValues();
}

function collectColumn(data, colIndex) {
  var result = [];
  for (var i = 0; i < data.length; i++) {
    var val = data[i][colIndex];
    if (val) result.push(String(val));
  }
  return result;
}

function parseCSVSet(val) {
  var set = {};
  if (!val) return set;
  var items = String(val).split(',');
  for (var i = 0; i < items.length; i++) {
    var item = trim(items[i]);
    if (item) set[item] = true;
  }
  return set;
}

function setToCSV(setObj) {
  var keys = [];
  for (var k in setObj) {
    if (setObj.hasOwnProperty(k) && k) keys.push(k);
  }
  keys.sort();
  return keys.join(', ');
}

// -------------------- LOOKUP MAPS --------------------

function buildSchoolMap(data) {
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var name = trim(data[i][1]);
    if (name) map[name] = data[i];
  }
  return map;
}

function buildTeacherMap(data) {
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var name = trim(data[i][1]);
    if (name) map[name] = data[i];
  }
  return map;
}

function buildGroupMap(data) {
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var schoolID = trim(data[i][6]);
    var groupName = trim(data[i][1]);
    if (schoolID && groupName) map[schoolID + '|' + groupName] = data[i];
  }
  return map;
}

function buildStudentMap(data) {
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var groupID = trim(data[i][5]);
    var name = trim(data[i][1]);
    if (groupID && name) map[groupID + '|' + name] = data[i];
  }
  return map;
}

function getSchoolID(schoolName, schoolMap) {
  var entry = schoolMap[schoolName];
  return entry ? String(entry[0]) : '';
}

function getTeacherID(teacherName, teacherMap) {
  var entry = teacherMap[teacherName];
  return entry ? String(entry[0]) : '';
}

// -------------------- RELATIONSHIP UPDATERS --------------------

function updateSchoolRelationships(sheet, schoolGroupIDs, schoolTeacherIDs) {
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;
  var data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 10).getValues();
  for (var i = 0; i < data.length; i++) {
    var schoolID = trim(data[i][0]);
    if (!schoolID) continue;
    var rowNum = i + DATA_START_ROW;
    if (schoolGroupIDs[schoolID]) sheet.getRange(rowNum, 8).setValue(setToCSV(schoolGroupIDs[schoolID]));
    if (schoolTeacherIDs[schoolID]) sheet.getRange(rowNum, 9).setValue(setToCSV(schoolTeacherIDs[schoolID]));
  }
}

function updateTeacherRelationships(sheet, teacherGroupIDs) {
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;
  var data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 8).getValues();
  for (var i = 0; i < data.length; i++) {
    var teacherID = trim(data[i][0]);
    if (!teacherID) continue;
    var rowNum = i + DATA_START_ROW;
    if (teacherGroupIDs[teacherID]) sheet.getRange(rowNum, 7).setValue(setToCSV(teacherGroupIDs[teacherID]));
  }
}

function updateGroupStudentCounts(sheet, groupStudentCount) {
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;
  var data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 9).getValues();
  for (var i = 0; i < data.length; i++) {
    var groupID = trim(data[i][0]);
    if (!groupID) continue;
    if (groupStudentCount.hasOwnProperty(groupID)) {
      var rowNum = i + DATA_START_ROW;
      sheet.getRange(rowNum, 3).setValue(groupStudentCount[groupID]);
    }
  }
}
