/**
 * Student Movement Tracker — Shared Utilities (js/utils.js)
 *
 * Includes:
 * - XSS-safe HTML escaping
 * - Authoritative timestamp parsing (protecting against Google Sheets 1899 time-only values)
 * - Elapsed duration and Normal / Warning / Critical attention level calculation
 * - Bulk student & teacher CSV/TSV/Excel parser and validator with column alias support
 * - Zero-dependency browser .xlsx reader (parses Excel .xlsx files into CSV/tabular text)
 * - CSV export generator and safe DOM modal/toast helpers
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = factory();
  } else {
    root.SMT_UTILS = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {

  function escapeHtml(value) {
    if (value === null || value === undefined) return '';
    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function extractEpochFromMovementId(movementId) {
    var m = /^MOV-(\d{12,14})-[A-Z0-9]+$/i.exec(String(movementId || '').trim());
    if (m && m[1]) {
      var ms = parseInt(m[1], 10);
      if (!isNaN(ms) && ms > 1577836800000) {
        return ms;
      }
    }
    return null;
  }

  function parseAuthoritativeTimestamp(rawValue, movementId, referenceMs) {
    var embeddedEpoch = extractEpochFromMovementId(movementId);

    if (rawValue === null || rawValue === undefined || rawValue === '') {
      return embeddedEpoch ? new Date(embeddedEpoch) : null;
    }

    if (Object.prototype.toString.call(rawValue) === '[object Date]') {
      if (isNaN(rawValue.getTime())) {
        return embeddedEpoch ? new Date(embeddedEpoch) : null;
      }
      if (rawValue.getUTCFullYear() <= 1905) {
        if (embeddedEpoch) return new Date(embeddedEpoch);
        var ref = new Date(referenceMs || Date.now());
        ref.setHours(rawValue.getHours(), rawValue.getMinutes(), rawValue.getSeconds(), 0);
        return ref;
      }
      return rawValue;
    }

    var str = String(rawValue).trim();
    if (!str) {
      return embeddedEpoch ? new Date(embeddedEpoch) : null;
    }

    if (/^\d{12,14}$/.test(str)) {
      return new Date(parseInt(str, 10));
    }

    var timeOnlyMatch = /^(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*(AM|PM))?$/i.exec(str);
    if (timeOnlyMatch) {
      if (embeddedEpoch && !referenceMs) {
        return new Date(embeddedEpoch);
      }
      var baseDate = new Date(referenceMs || embeddedEpoch || Date.now());
      var hrs = parseInt(timeOnlyMatch[1], 10);
      var mins = parseInt(timeOnlyMatch[2], 10);
      var secs = timeOnlyMatch[3] ? parseInt(timeOnlyMatch[3], 10) : 0;
      var ampm = timeOnlyMatch[4] ? timeOnlyMatch[4].toUpperCase() : null;
      if (ampm === 'PM' && hrs < 12) hrs += 12;
      if (ampm === 'AM' && hrs === 12) hrs = 0;
      baseDate.setHours(hrs, mins, secs, 0);
      return baseDate;
    }

    var parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      if (parsed.getUTCFullYear() <= 1905) {
        if (embeddedEpoch) return new Date(embeddedEpoch);
        var fallback = new Date(referenceMs || Date.now());
        fallback.setHours(parsed.getHours(), parsed.getMinutes(), parsed.getSeconds(), 0);
        return fallback;
      }
      return parsed;
    }

    return embeddedEpoch ? new Date(embeddedEpoch) : null;
  }

  function calculateMovementStatus(timeOutRaw, timeInRaw, movementId, warningMinutes, criticalMinutes, nowMsOverride) {
    var warnThreshold = Number(warningMinutes || 10);
    var critThreshold = Number(criticalMinutes || 15);
    var outDate = parseAuthoritativeTimestamp(timeOutRaw, movementId, null);
    if (!outDate) {
      return {
        valid: false,
        checkoutDate: null,
        returnDate: null,
        elapsedSeconds: 0,
        elapsedMinutes: 0,
        formattedDuration: '0m 00s',
        attentionLevel: 'Normal'
      };
    }

    var outMs = outDate.getTime();
    var inDate = (timeInRaw !== null && timeInRaw !== undefined && String(timeInRaw).trim() !== '')
      ? parseAuthoritativeTimestamp(timeInRaw, null, outMs)
      : null;

    var endMs = inDate ? inDate.getTime() : (nowMsOverride || Date.now());
    var elapsedSeconds = Math.max(0, Math.floor((endMs - outMs) / 1000));
    var elapsedMinutes = elapsedSeconds / 60;
    var mins = Math.floor(elapsedSeconds / 60);
    var secs = elapsedSeconds % 60;
    var formattedDuration = mins + 'm ' + (secs < 10 ? '0' + secs : secs) + 's';

    var attentionLevel = 'Normal';
    if (elapsedMinutes >= critThreshold) {
      attentionLevel = 'Critical';
    } else if (elapsedMinutes >= warnThreshold) {
      attentionLevel = 'Warning';
    }

    return {
      valid: true,
      checkoutDate: outDate,
      returnDate: inDate,
      elapsedSeconds: elapsedSeconds,
      elapsedMinutes: Math.round(elapsedMinutes * 100) / 100,
      formattedDuration: formattedDuration,
      attentionLevel: attentionLevel
    };
  }

  function formatTimeShort(isoOrRaw, movementId) {
    var d = parseAuthoritativeTimestamp(isoOrRaw, movementId, null);
    if (!d) return '—';
    return d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  }

  function formatDateTime(isoOrRaw, movementId) {
    var d = parseAuthoritativeTimestamp(isoOrRaw, movementId, null);
    if (!d) return '—';
    return d.toLocaleDateString() + ' ' + d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  }

  function attentionBadgeHtml(level) {
    var norm = String(level || 'Normal').trim();
    var cls = 'badge-normal';
    if (norm === 'Critical') cls = 'badge-critical';
    else if (norm === 'Warning') cls = 'badge-warning';
    return '<span class="badge ' + cls + '">' + escapeHtml(norm) + '</span>';
  }

  function splitDelimitedLine(line, delimiter) {
    var result = [];
    var current = '';
    var inQuotes = false;
    for (var i = 0; i < line.length; i++) {
      var ch = line[i];
      if (ch === '"') {
        if (inQuotes && line[i + 1] === '"') {
          current += '"';
          i++;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (ch === delimiter && !inQuotes) {
        result.push(current.trim());
        current = '';
      } else {
        current += ch;
      }
    }
    result.push(current.trim());
    return result;
  }

  function detectDelimiter(headerLine) {
    if (headerLine.indexOf('\t') !== -1) return '\t';
    if (headerLine.indexOf(',') !== -1) return ',';
    if (headerLine.indexOf(';') !== -1) return ';';
    return ',';
  }

  /**
   * Parses and validates CSV or pasted Excel/tabular student data for Admin Bulk Import preview.
   */
  function parseAndValidateStudentImport(rawText, classesList, existingStudentsList) {
    var text = String(rawText || '').trim();
    if (!text) {
      return { valid: [], skipped: [], invalid: [], totalRows: 0 };
    }

    var lines = text.split(/\r?\n/).map(function (l) { return l.trim(); }).filter(Boolean);
    if (lines.length === 0) {
      return { valid: [], skipped: [], invalid: [], totalRows: 0 };
    }

    var delimiter = detectDelimiter(lines[0]);
    var firstCols = splitDelimitedLine(lines[0], delimiter);

    var idAliases = ['studentid', 'student_id', 'id', 'studentno', 'student_no', 'admissionno', 'admission_no', 'index', 'indexno'];
    var nameAliases = ['studentname', 'student_name', 'name', 'fullname', 'full_name', 'student'];
    var classAliases = ['classid', 'class_id', 'classname', 'class_name', 'class', 'grade', 'form', 'section'];
    var statusAliases = ['status', 'state'];

    var colMap = { id: 0, name: 1, classRef: 2, status: 3 };
    var hasHeader = false;

    var normalizedFirst = firstCols.map(function (c) {
      return c.toLowerCase().replace(/[^a-z0-9_]/g, '');
    });

    var matchedHeaderCount = 0;
    var detectedMap = { id: -1, name: -1, classRef: -1, status: -1 };
    for (var cIdx = 0; cIdx < normalizedFirst.length; cIdx++) {
      var h = normalizedFirst[cIdx];
      if (idAliases.indexOf(h) !== -1 && detectedMap.id === -1) {
        detectedMap.id = cIdx;
        matchedHeaderCount++;
      } else if (nameAliases.indexOf(h) !== -1 && detectedMap.name === -1) {
        detectedMap.name = cIdx;
        matchedHeaderCount++;
      } else if (classAliases.indexOf(h) !== -1 && detectedMap.classRef === -1) {
        detectedMap.classRef = cIdx;
        matchedHeaderCount++;
      } else if (statusAliases.indexOf(h) !== -1 && detectedMap.status === -1) {
        detectedMap.status = cIdx;
      }
    }

    if (matchedHeaderCount >= 2) {
      hasHeader = true;
      colMap.id = detectedMap.id !== -1 ? detectedMap.id : 0;
      colMap.name = detectedMap.name !== -1 ? detectedMap.name : 1;
      colMap.classRef = detectedMap.classRef !== -1 ? detectedMap.classRef : 2;
      colMap.status = detectedMap.status;
    }

    var classById = {};
    var classByName = {};
    var classesArr = Array.isArray(classesList) ? classesList : [];
    for (var k = 0; k < classesArr.length; k++) {
      var cls = classesArr[k];
      var cid = String(cls.ClassID || '').trim();
      var cname = String(cls.ClassName || '').trim();
      if (cid) classById[cid.toLowerCase()] = { ClassID: cid, ClassName: cname || cid };
      if (cname) classByName[cname.toLowerCase()] = { ClassID: cid, ClassName: cname };
    }

    var existingIdSet = {};
    var existingArr = Array.isArray(existingStudentsList) ? existingStudentsList : [];
    for (var e = 0; e < existingArr.length; e++) {
      var sid = String(existingArr[e].StudentID || '').trim().toLowerCase();
      if (sid) existingIdSet[sid] = true;
    }

    var startIdx = hasHeader ? 1 : 0;
    var seenInBatch = {};
    var valid = [];
    var skipped = [];
    var invalid = [];

    for (var r = startIdx; r < lines.length; r++) {
      var cols = splitDelimitedLine(lines[r], delimiter);
      var rowNumber = r + 1;
      var studentId = String(cols[colMap.id] || '').trim();
      var studentName = String(cols[colMap.name] || '').trim();
      var classRef = String(cols[colMap.classRef] || '').trim();
      var status = (colMap.status !== -1 && cols[colMap.status]) ? String(cols[colMap.status]).trim() : 'Active';
      if (status.toLowerCase() === 'inactive') status = 'Inactive';
      else status = 'Active';

      if (!studentId || !studentName || !classRef) {
        invalid.push({
          rowNumber: rowNumber,
          studentId: studentId,
          studentName: studentName,
          classRef: classRef,
          reason: 'Missing StudentID, StudentName, or Class'
        });
        continue;
      }

      var matchedClass = classById[classRef.toLowerCase()] || classByName[classRef.toLowerCase()] || null;
      if (!matchedClass) {
        invalid.push({
          rowNumber: rowNumber,
          studentId: studentId,
          studentName: studentName,
          classRef: classRef,
          reason: 'Unknown Class "' + classRef + '"'
        });
        continue;
      }

      var idLower = studentId.toLowerCase();
      if (existingIdSet[idLower]) {
        skipped.push({
          rowNumber: rowNumber,
          studentId: studentId,
          studentName: studentName,
          classId: matchedClass.ClassID,
          className: matchedClass.ClassName,
          reason: 'Student ID already exists in Google Sheets'
        });
        continue;
      }

      if (seenInBatch[idLower]) {
        invalid.push({
          rowNumber: rowNumber,
          studentId: studentId,
          studentName: studentName,
          classRef: classRef,
          reason: 'Duplicate Student ID within import file'
        });
        continue;
      }

      seenInBatch[idLower] = true;
      valid.push({
        rowNumber: rowNumber,
        studentId: studentId,
        studentName: studentName,
        classId: matchedClass.ClassID,
        className: matchedClass.ClassName,
        status: status
      });
    }

    return {
      valid: valid,
      skipped: skipped,
      invalid: invalid,
      totalRows: lines.length - startIdx
    };
  }

  /**
   * Parses and validates CSV, Excel copy-paste (TSV), or parsed .xlsx rows for Bulk Teacher Import.
   * Supports:
   * - TeacherID (or auto-generated TCH-xxx if omitted/blank)
   * - TeacherName (required)
   * - Optional Classes (e.g., "Grade 10A; Grade 10B" or "CLS-10A")
   * - Optional Status ("Active" / "Inactive")
   */
  function parseAndValidateTeacherImport(rawText, existingTeachersList, classesList) {
    var text = String(rawText || '').trim();
    if (!text) {
      return { valid: [], skipped: [], invalid: [], totalRows: 0 };
    }

    var lines = text.split(/\r?\n/).map(function (l) { return l.trim(); }).filter(Boolean);
    if (lines.length === 0) {
      return { valid: [], skipped: [], invalid: [], totalRows: 0 };
    }

    var delimiter = detectDelimiter(lines[0]);
    var firstCols = splitDelimitedLine(lines[0], delimiter);

    var idAliases = ['teacherid', 'teacher_id', 'id', 'staffid', 'staff_id', 'empid', 'emp_id', 'employeeid', 'employee_id', 'teacherno', 'code'];
    var nameAliases = ['teachername', 'teacher_name', 'name', 'fullname', 'full_name', 'teacher', 'staffname', 'staff_name'];
    var classAliases = ['classes', 'class', 'classid', 'class_id', 'classname', 'class_name', 'assignedclasses', 'assigned_classes', 'grade'];
    var statusAliases = ['status', 'state'];

    var normalizedFirst = firstCols.map(function (c) {
      return c.toLowerCase().replace(/[^a-z0-9_]/g, '');
    });

    var hasHeader = false;
    var detectedMap = { id: -1, name: -1, classes: -1, status: -1 };
    var matchedHeaders = 0;

    for (var cIdx = 0; cIdx < normalizedFirst.length; cIdx++) {
      var h = normalizedFirst[cIdx];
      if (idAliases.indexOf(h) !== -1 && detectedMap.id === -1) {
        detectedMap.id = cIdx;
        matchedHeaders++;
      } else if (nameAliases.indexOf(h) !== -1 && detectedMap.name === -1) {
        detectedMap.name = cIdx;
        matchedHeaders++;
      } else if (classAliases.indexOf(h) !== -1 && detectedMap.classes === -1) {
        detectedMap.classes = cIdx;
        matchedHeaders++;
      } else if (statusAliases.indexOf(h) !== -1 && detectedMap.status === -1) {
        detectedMap.status = cIdx;
        matchedHeaders++;
      }
    }

    var colMap = { id: 0, name: 1, classes: 2, status: 3 };
    if (matchedHeaders >= 1 && (detectedMap.name !== -1 || detectedMap.id !== -1)) {
      hasHeader = true;
      if (detectedMap.id !== -1 && detectedMap.name !== -1) {
        colMap.id = detectedMap.id;
        colMap.name = detectedMap.name;
      } else if (detectedMap.name !== -1 && detectedMap.id === -1) {
        colMap.id = -1; // Single-column or no-ID table where TeacherID is auto-generated
        colMap.name = detectedMap.name;
      } else {
        colMap.id = detectedMap.id;
        colMap.name = detectedMap.id === 0 ? 1 : 0;
      }
      colMap.classes = detectedMap.classes;
      colMap.status = detectedMap.status;
    } else if (firstCols.length === 1) {
      // Single column of teacher names without header
      colMap = { id: -1, name: 0, classes: -1, status: -1 };
    }

    var existingIdSet = {};
    var existingTeachers = Array.isArray(existingTeachersList) ? existingTeachersList : [];
    for (var e = 0; e < existingTeachers.length; e++) {
      var tid = String(existingTeachers[e].TeacherID || '').trim().toLowerCase();
      if (tid) existingIdSet[tid] = true;
    }

    var classById = {};
    var classByName = {};
    var classesArr = Array.isArray(classesList) ? classesList : [];
    for (var k = 0; k < classesArr.length; k++) {
      var cls = classesArr[k];
      var cid = String(cls.ClassID || '').trim();
      var cname = String(cls.ClassName || '').trim();
      if (cid) classById[cid.toLowerCase()] = { ClassID: cid, ClassName: cname || cid };
      if (cname) classByName[cname.toLowerCase()] = { ClassID: cid, ClassName: cname };
    }

    var startIdx = hasHeader ? 1 : 0;
    var seenInBatch = {};
    var autoIdCounter = existingTeachers.length + 101;
    var valid = [];
    var skipped = [];
    var invalid = [];

    for (var r = startIdx; r < lines.length; r++) {
      var cols = splitDelimitedLine(lines[r], delimiter);
      var rowNumber = r + 1;

      var teacherId = colMap.id !== -1 ? String(cols[colMap.id] || '').trim() : '';
      var teacherName = colMap.name !== -1 ? String(cols[colMap.name] || '').trim() : '';
      // If user pasted 1 column (only name) when 2 columns were default
      if (teacherId && !teacherName && cols.length === 1) {
        teacherName = teacherId;
        teacherId = '';
      }
      var rawClasses = (colMap.classes !== -1 && cols[colMap.classes]) ? String(cols[colMap.classes]).trim() : '';
      var status = (colMap.status !== -1 && cols[colMap.status]) ? String(cols[colMap.status]).trim() : 'Active';
      if (status.toLowerCase() === 'inactive') status = 'Inactive';
      else status = 'Active';

      if (!teacherName) {
        invalid.push({
          rowNumber: rowNumber,
          teacherId: teacherId || '—',
          teacherName: '',
          reason: 'Missing Teacher Name'
        });
        continue;
      }

      if (!teacherId) {
        do {
          teacherId = 'TCH-' + (autoIdCounter++);
        } while (existingIdSet[teacherId.toLowerCase()] || seenInBatch[teacherId.toLowerCase()]);
      }

      var idLower = teacherId.toLowerCase();
      if (existingIdSet[idLower]) {
        skipped.push({
          rowNumber: rowNumber,
          teacherId: teacherId,
          teacherName: teacherName,
          reason: 'Teacher ID already exists in Google Sheets'
        });
        continue;
      }

      if (seenInBatch[idLower]) {
        invalid.push({
          rowNumber: rowNumber,
          teacherId: teacherId,
          teacherName: teacherName,
          reason: 'Duplicate Teacher ID within import file'
        });
        continue;
      }

      // Validate optional classes if provided
      var resolvedClassIds = [];
      var resolvedClassNames = [];
      var unknownClass = null;
      if (rawClasses) {
        var parts = rawClasses.split(/[;|,]/).map(function (s) { return s.trim(); }).filter(Boolean);
        var seenC = {};
        for (var p = 0; p < parts.length; p++) {
          var mCls = classById[parts[p].toLowerCase()] || classByName[parts[p].toLowerCase()] || null;
          if (!mCls) {
            unknownClass = parts[p];
            break;
          }
          if (!seenC[mCls.ClassID]) {
            seenC[mCls.ClassID] = true;
            resolvedClassIds.push(mCls.ClassID);
            resolvedClassNames.push(mCls.ClassName);
          }
        }
      }

      if (unknownClass) {
        invalid.push({
          rowNumber: rowNumber,
          teacherId: teacherId,
          teacherName: teacherName,
          reason: 'Unknown Class "' + unknownClass + '"'
        });
        continue;
      }

      seenInBatch[idLower] = true;
      valid.push({
        rowNumber: rowNumber,
        teacherId: teacherId,
        teacherName: teacherName,
        classIds: resolvedClassIds,
        classNames: resolvedClassNames,
        status: status
      });
    }

    return {
      valid: valid,
      skipped: skipped,
      invalid: invalid,
      totalRows: lines.length - startIdx
    };
  }

  /**
   * Lightweight, zero-dependency browser .xlsx reader using DecompressionStream & DOMParser.
   * Extracts the first worksheet from an Excel .xlsx ArrayBuffer and returns tab-delimited text.
   */
  async function parseXlsxArrayBufferToText(arrayBuffer) {
    var bytes = new Uint8Array(arrayBuffer);
    var view = new DataView(arrayBuffer);
    var entries = {};

    // Locate End of Central Directory (EOCD) signature 0x06054b50
    var eocdOffset = -1;
    for (var i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
      if (view.getUint32(i, true) === 0x06054b50) {
        eocdOffset = i;
        break;
      }
    }
    if (eocdOffset === -1) {
      throw new Error('Invalid .xlsx file format (ZIP directory not found). Save as .csv or .xlsx.');
    }

    var totalEntries = view.getUint16(eocdOffset + 10, true);
    var cdOffset = view.getUint32(eocdOffset + 16, true);
    var ptr = cdOffset;
    var decoder = new TextDecoder('utf-8');

    for (var e = 0; e < totalEntries; e++) {
      if (view.getUint32(ptr, true) !== 0x02014b50) break;
      var compMethod = view.getUint16(ptr + 10, true);
      var compSize = view.getUint32(ptr + 20, true);
      var fileNameLen = view.getUint16(ptr + 28, true);
      var extraLen = view.getUint16(ptr + 30, true);
      var commentLen = view.getUint16(ptr + 32, true);
      var localHeaderOffset = view.getUint32(ptr + 42, true);
      var fileName = decoder.decode(bytes.subarray(ptr + 46, ptr + 46 + fileNameLen));

      if (fileName === 'xl/sharedStrings.xml' || /^xl\/worksheets\/sheet\d+\.xml$/i.test(fileName)) {
        var localFileNameLen = view.getUint16(localHeaderOffset + 26, true);
        var localExtraLen = view.getUint16(localHeaderOffset + 28, true);
        var dataStart = localHeaderOffset + 30 + localFileNameLen + localExtraLen;
        var compData = bytes.subarray(dataStart, dataStart + compSize);

        var xmlText = '';
        if (compMethod === 0) {
          xmlText = decoder.decode(compData);
        } else if (compMethod === 8 && typeof DecompressionStream !== 'undefined') {
          var ds = new DecompressionStream('deflate-raw');
          var writer = ds.writable.getWriter();
          writer.write(compData);
          writer.close();
          var resp = new Response(ds.readable);
          var decompressedBuf = await resp.arrayBuffer();
          xmlText = decoder.decode(new Uint8Array(decompressedBuf));
        } else {
          throw new Error('Browser does not support DecompressionStream for compressed .xlsx. Please save as CSV or copy-paste cells directly from Excel.');
        }
        entries[fileName] = xmlText;
      }
      ptr += 46 + fileNameLen + extraLen + commentLen;
    }

    var sheetKey = entries['xl/worksheets/sheet1.xml']
      ? 'xl/worksheets/sheet1.xml'
      : Object.keys(entries).find(function (k) { return /^xl\/worksheets\/sheet/i.test(k); });

    if (!sheetKey) {
      throw new Error('No worksheet found inside .xlsx file.');
    }

    var sharedStrings = [];
    if (entries['xl/sharedStrings.xml']) {
      var ssXml = entries['xl/sharedStrings.xml'];
      var siMatches = ssXml.match(/<si[\s>][\s\S]*?<\/si>/g) || [];
      for (var s = 0; s < siMatches.length; s++) {
        var tMatches = siMatches[s].match(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g) || [];
        var combined = tMatches.map(function (tTag) {
          return tTag.replace(/<[^>]+>/g, '');
        }).join('');
        sharedStrings.push(
          combined
            .replace(/&amp;/g, '&')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&quot;/g, '"')
            .replace(/&#39;/g, "'")
        );
      }
    }

    var colLettersToIndex = function (ref) {
      var m = /^([A-Z]+)/i.exec(ref || '');
      if (!m) return 0;
      var letters = m[1].toUpperCase();
      var idx = 0;
      for (var c = 0; c < letters.length; c++) {
        idx = idx * 26 + (letters.charCodeAt(c) - 64);
      }
      return idx - 1;
    };

    var sheetXml = entries[sheetKey];
    var rowBlocks = sheetXml.match(/<row[\s>][\s\S]*?<\/row>/g) || [];
    var outputLines = [];

    for (var r = 0; r < rowBlocks.length; r++) {
      var cellBlocks = rowBlocks[r].match(/<c[\s>][\s\S]*?(?:<\/c>|\/>)/g) || [];
      var rowCells = [];
      for (var c = 0; c < cellBlocks.length; c++) {
        var cellXml = cellBlocks[c];
        var refMatch = /\br=["']([A-Z]+\d+)["']/i.exec(cellXml);
        var colIdx = refMatch ? colLettersToIndex(refMatch[1]) : c;
        var typeMatch = /\bt=["']([^"']+)["']/i.exec(cellXml);
        var cellType = typeMatch ? typeMatch[1] : '';
        var val = '';

        if (cellType === 'inlineStr') {
          var isMatch = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/i.exec(cellXml);
          val = isMatch ? isMatch[1] : '';
        } else {
          var vMatch = /<v>([\s\S]*?)<\/v>/i.exec(cellXml);
          var rawV = vMatch ? vMatch[1].trim() : '';
          if (cellType === 's') {
            var ssIdx = parseInt(rawV, 10);
            val = !isNaN(ssIdx) && sharedStrings[ssIdx] !== undefined ? sharedStrings[ssIdx] : rawV;
          } else {
            val = rawV;
          }
        }
        val = String(val || '')
          .replace(/&amp;/g, '&')
          .replace(/&lt;/g, '<')
          .replace(/&gt;/g, '>')
          .replace(/&quot;/g, '"')
          .replace(/&#39;/g, "'")
          .trim();
        rowCells[colIdx] = val;
      }
      if (rowCells.some(function (cell) { return cell !== undefined && cell !== ''; })) {
        for (var fill = 0; fill < rowCells.length; fill++) {
          if (rowCells[fill] === undefined) rowCells[fill] = '';
        }
        outputLines.push(rowCells.join('\t'));
      }
    }

    return outputLines.join('\n');
  }

  /**
   * Reads a File object (.csv, .tsv, .txt, or .xlsx) and resolves to tabular text.
   */
  function readSpreadsheetFileAsText(file) {
    return new Promise(function (resolve, reject) {
      if (!file) {
        reject(new Error('No file selected'));
        return;
      }
      var name = String(file.name || '').toLowerCase();
      if (name.endsWith('.xlsx')) {
        var arrayReader = new FileReader();
        arrayReader.onload = function (evt) {
          parseXlsxArrayBufferToText(evt.target.result).then(resolve).catch(reject);
        };
        arrayReader.onerror = function () {
          reject(new Error('Failed to read Excel file.'));
        };
        arrayReader.readAsArrayBuffer(file);
      } else {
        var textReader = new FileReader();
        textReader.onload = function (evt) {
          resolve(String(evt.target.result || ''));
        };
        textReader.onerror = function () {
          reject(new Error('Failed to read file.'));
        };
        textReader.readAsText(file);
      }
    });
  }

  function movementRecordsToCsv(records) {
    var headers = [
      'MovementID',
      'StudentID',
      'StudentName',
      'ClassID',
      'ClassName',
      'TeacherID',
      'TeacherName',
      'Destination',
      'TimeOut',
      'TimeIn',
      'Duration',
      'AttentionLevel',
      'Status',
      'Notes'
    ];
    var escapeCsvCell = function (val) {
      var str = val === null || val === undefined ? '' : String(val);
      if (/^[=+\-@]/.test(str)) {
        str = "'" + str;
      }
      if (/[",\r\n]/.test(str)) {
        return '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    };

    var rows = [headers.join(',')];
    var list = Array.isArray(records) ? records : [];
    for (var i = 0; i < list.length; i++) {
      var r = list[i];
      var line = [
        r.MovementID,
        r.StudentID,
        r.StudentName,
        r.ClassID,
        r.ClassName,
        r.TeacherID,
        r.TeacherName,
        r.Destination,
        r.TimeOut,
        r.TimeIn || '',
        r.formattedDuration || '',
        r.attentionLevel || 'Normal',
        r.Status || '',
        r.Notes || ''
      ].map(escapeCsvCell).join(',');
      rows.push(line);
    }
    return rows.join('\r\n');
  }

  function downloadCsvFile(filename, csvContent) {
    if (typeof document === 'undefined') return;
    var blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename || 'movement-report.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function showToast(message, type) {
    if (typeof document === 'undefined') return;
    var container = document.getElementById('smt-toast-container');
    if (!container) {
      container = document.createElement('div');
      container.id = 'smt-toast-container';
      container.className = 'toast-container';
      container.setAttribute('aria-live', 'polite');
      document.body.appendChild(container);
    }
    var toast = document.createElement('div');
    var kind = type || 'info';
    toast.className = 'toast toast-' + kind;
    toast.textContent = String(message || '');
    container.appendChild(toast);
    setTimeout(function () {
      if (toast.parentNode) {
        toast.parentNode.removeChild(toast);
      }
    }, 4000);
  }

  function confirmDialog(options) {
    return new Promise(function (resolve) {
      if (typeof document === 'undefined') {
        resolve(true);
        return;
      }
      var backdrop = document.getElementById('smt-confirm-modal');
      if (!backdrop) {
        backdrop = document.createElement('div');
        backdrop.id = 'smt-confirm-modal';
        backdrop.className = 'modal-backdrop';
        backdrop.innerHTML = [
          '<div class="modal-dialog" role="dialog" aria-modal="true" aria-labelledby="smt-confirm-title">',
          '  <h3 class="modal-title" id="smt-confirm-title">Confirm Action</h3>',
          '  <p id="smt-confirm-message" style="color: var(--text-secondary); margin: 0;"></p>',
          '  <div class="modal-actions">',
          '    <button type="button" class="btn btn-secondary" id="smt-confirm-cancel">Cancel</button>',
          '    <button type="button" class="btn btn-primary" id="smt-confirm-ok">Confirm</button>',
          '  </div>',
          '</div>'
        ].join('');
        document.body.appendChild(backdrop);
      }
      var titleEl = document.getElementById('smt-confirm-title');
      var msgEl = document.getElementById('smt-confirm-message');
      var okBtn = document.getElementById('smt-confirm-ok');
      var cancelBtn = document.getElementById('smt-confirm-cancel');

      titleEl.textContent = (options && options.title) || 'Confirm Action';
      msgEl.textContent = (options && options.message) || 'Are you sure you want to proceed?';
      okBtn.textContent = (options && options.confirmText) || 'Confirm';
      okBtn.className = (options && options.danger) ? 'btn btn-danger' : 'btn btn-primary';

      var cleanup = function (result) {
        backdrop.classList.remove('open');
        okBtn.onclick = null;
        cancelBtn.onclick = null;
        resolve(result);
      };

      okBtn.onclick = function () { cleanup(true); };
      cancelBtn.onclick = function () { cleanup(false); };
      backdrop.classList.add('open');
    });
  }

  return {
    escapeHtml: escapeHtml,
    extractEpochFromMovementId: extractEpochFromMovementId,
    parseAuthoritativeTimestamp: parseAuthoritativeTimestamp,
    calculateMovementStatus: calculateMovementStatus,
    formatTimeShort: formatTimeShort,
    formatDateTime: formatDateTime,
    attentionBadgeHtml: attentionBadgeHtml,
    parseAndValidateStudentImport: parseAndValidateStudentImport,
    parseAndValidateTeacherImport: parseAndValidateTeacherImport,
    parseXlsxArrayBufferToText: parseXlsxArrayBufferToText,
    readSpreadsheetFileAsText: readSpreadsheetFileAsText,
    movementRecordsToCsv: movementRecordsToCsv,
    downloadCsvFile: downloadCsvFile,
    showToast: showToast,
    confirmDialog: confirmDialog
  };
});
