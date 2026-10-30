/**
 * Student Movement Tracker — Shared Utilities (js/utils.js)
 *
 * Includes:
 * - XSS-safe HTML escaping
 * - Authoritative timestamp parsing (protecting against Google Sheets 1899 time-only values)
 * - Elapsed duration and Normal / Warning / Critical attention level calculation
 * - Bulk student CSV/TSV parser and validator with column alias support
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

  /**
   * Extracts embedded epoch milliseconds from MovementID (MOV-<epoch_ms>-<hex>).
   */
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

  /**
   * Safely parses checkout or return timestamps.
   * Never treats a Google Sheets 1899/1900 time-only date as an actual historical year.
   */
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

  /**
   * Computes elapsed duration and attention level:
   * - Less than 10 minutes: Normal
   * - 10 minutes to less than 15 minutes: Warning
   * - 15 minutes or more: Critical
   */
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

  /**
   * Parses a single CSV/TSV line respecting quoted fields.
   */
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

  /**
   * Parses and validates CSV or pasted tabular student data for Admin Bulk Import preview.
   * Supports column aliases for StudentID, StudentName, and ClassID/ClassName.
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

    var delimiter = lines[0].indexOf('\t') !== -1 ? '\t' : ',';
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
   * Formats movement records into a clean, formula-safe CSV string.
   */
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
        str = "'" + str; // Prevent spreadsheet formula injection
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
    movementRecordsToCsv: movementRecordsToCsv,
    downloadCsvFile: downloadCsvFile,
    showToast: showToast,
    confirmDialog: confirmDialog
  };
});
