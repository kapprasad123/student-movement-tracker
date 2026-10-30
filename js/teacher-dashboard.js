/**
 * Student Movement Tracker — Teacher Dashboard Controller (js/teacher-dashboard.js)
 */
(function () {
  var state = {
    teachers: [],
    classes: [],
    destinations: [],
    studentsInClass: [],
    currentOutRecords: [],
    selectedClassId: '',
    selectedDestination: '',
    selectedStudentId: '',
    warningMinutes: 10,
    criticalMinutes: 15,
    submittingOut: false,
    returningIds: {},
    refreshTimer: null,
    durationTickTimer: null
  };

  var els = {};

  function cacheElements() {
    els.schoolName = document.getElementById('header-school-name');
    els.connPill = document.getElementById('connection-status-pill');
    els.connText = document.getElementById('connection-status-text');
    els.sessionPill = document.getElementById('teacher-session-pill');
    els.btnLockSession = document.getElementById('btn-lock-session');

    els.endpointBanner = document.getElementById('endpoint-config-banner');
    els.endpointForm = document.getElementById('endpoint-config-form');
    els.inputWebAppUrl = document.getElementById('input-webapp-url');

    els.authSection = document.getElementById('teacher-auth-section');
    els.loginForm = document.getElementById('teacher-login-form');
    els.selectLoginTeacher = document.getElementById('select-login-teacher');
    els.inputTeacherPin = document.getElementById('input-teacher-pin');
    els.btnTeacherLogin = document.getElementById('btn-teacher-login');
    els.btnReloadLogin = document.getElementById('btn-reload-login-options');
    els.authFeedback = document.getElementById('teacher-auth-feedback');

    els.workspace = document.getElementById('teacher-workspace');
    els.activeTeacherLabel = document.getElementById('active-teacher-label');
    els.selectSwitchTeacher = document.getElementById('select-switch-teacher');
    els.recordOutForm = document.getElementById('record-out-form');
    els.selectClass = document.getElementById('select-class');
    els.selectDestination = document.getElementById('select-destination');
    els.destChipsContainer = document.getElementById('destination-chips-container');
    els.inputStudentSearch = document.getElementById('input-student-search');
    els.rosterContainer = document.getElementById('student-roster-container');
    els.selectedStudentIdInput = document.getElementById('selected-student-id');
    els.inputNotes = document.getElementById('input-movement-notes');
    els.selectedStudentSummary = document.getElementById('selected-student-summary');
    els.btnMarkOut = document.getElementById('btn-mark-out');

    els.currentOutCount = document.getElementById('current-out-count');
    els.lastRefreshTime = document.getElementById('last-refresh-time');
    els.btnManualRefresh = document.getElementById('btn-manual-refresh');
    els.currentOutTbody = document.getElementById('current-out-tbody');
  }

  function updateConnectionIndicator(connState) {
    if (!els.connPill || !els.connText) return;
    els.connPill.className = 'status-pill';
    if (connState.status === 'connected') {
      els.connPill.classList.add('conn-connected');
      els.connText.textContent = 'Connected to Google Sheets';
    } else if (connState.status === 'loading') {
      els.connPill.classList.add('conn-loading');
      els.connText.textContent = connState.message || 'Syncing...';
    } else if (connState.status === 'offline') {
      els.connPill.classList.add('conn-offline');
      els.connText.textContent = 'Offline';
    } else if (connState.status === 'unconfigured') {
      els.connPill.classList.add('conn-unconfigured');
      els.connText.textContent = 'Endpoint Not Configured';
    } else {
      els.connPill.classList.add('conn-error');
      els.connText.textContent = connState.message || 'Connection Error';
    }

    var hint = document.getElementById('endpoint-banner-hint');
    var pill = document.getElementById('endpoint-env-pill');
    if (pill) {
      var isProd = SMT_CONFIG.isUsingProduction();
      var isLoc = SMT_CONFIG.isUsingLocalOverride();
      pill.textContent = isProd ? 'Production' : (isLoc ? 'Local Test Proxy' : 'Custom');
      pill.className = 'endpoint-env-pill' + (isProd ? '' : ' local-dev');
      if (hint) {
        hint.textContent = isProd
          ? 'Shared Google Apps Script Web App (Production)'
          : (isLoc ? 'Local Test Mock API (http://localhost:3000/api)' : SMT_CONFIG.getWebAppUrl());
      }
    }
  }

  function renderTeacherSelectOptions() {
    var html = ['<option value="">— Select Active Teacher —</option>'];
    var switchHtml = ['<option value="">— Switch Teacher —</option>'];
    for (var i = 0; i < state.teachers.length; i++) {
      var t = state.teachers[i];
      var id = t.TeacherID || t.teacherId || '';
      var name = t.TeacherName || t.teacherName || '';
      var opt = '<option value="' + SMT_UTILS.escapeHtml(id) + '">' +
        SMT_UTILS.escapeHtml(name) + ' (' + SMT_UTILS.escapeHtml(id) + ')</option>';
      html.push(opt);
      switchHtml.push(opt);
    }
    els.selectLoginTeacher.innerHTML = html.join('');
    els.selectSwitchTeacher.innerHTML = switchHtml.join('');
  }

  function renderDestinationControls() {
    var opts = ['<option value="">— Select Destination —</option>'];
    var chips = [];
    for (var i = 0; i < state.destinations.length; i++) {
      var d = state.destinations[i];
      var isSel = state.selectedDestination === d;
      opts.push('<option value="' + SMT_UTILS.escapeHtml(d) + '"' + (isSel ? ' selected' : '') + '>' + SMT_UTILS.escapeHtml(d) + '</option>');
      chips.push(
        '<button type="button" class="dest-chip' + (isSel ? ' selected' : '') + '" data-destination="' + SMT_UTILS.escapeHtml(d) + '">' +
        SMT_UTILS.escapeHtml(d) + '</button>'
      );
    }
    els.selectDestination.innerHTML = opts.join('');
    els.destChipsContainer.innerHTML = chips.join('');
  }

  function renderClassOptions() {
    var html = ['<option value="">— Select Class —</option>'];
    for (var i = 0; i < state.classes.length; i++) {
      var c = state.classes[i];
      var sel = state.selectedClassId === c.ClassID ? ' selected' : '';
      html.push(
        '<option value="' + SMT_UTILS.escapeHtml(c.ClassID) + '"' + sel + '>' +
        SMT_UTILS.escapeHtml(c.ClassName) + ' (' + c.activeEnrolledCount + ' students)</option>'
      );
    }
    els.selectClass.innerHTML = html.join('');
  }

  function renderStudentRoster() {
    if (!state.selectedClassId) {
      els.rosterContainer.innerHTML = '<div class="empty-state">Select an active class above to view its students.</div>';
      updateMarkOutButtonState();
      return;
    }

    var search = String(els.inputStudentSearch.value || '').trim().toLowerCase();
    var outStudentMap = {};
    for (var m = 0; m < state.currentOutRecords.length; m++) {
      outStudentMap[state.currentOutRecords[m].StudentID] = state.currentOutRecords[m];
    }

    var filtered = state.studentsInClass.filter(function (s) {
      if (!search) return true;
      return s.StudentName.toLowerCase().indexOf(search) !== -1 || s.StudentID.toLowerCase().indexOf(search) !== -1;
    });

    if (filtered.length === 0) {
      els.rosterContainer.innerHTML = '<div class="empty-state">No matching active students found in this class.</div>';
      updateMarkOutButtonState();
      return;
    }

    var rowsHtml = [];
    for (var i = 0; i < filtered.length; i++) {
      var stu = filtered[i];
      var activeMov = outStudentMap[stu.StudentID] || (stu.isCurrentlyOut ? { MovementID: stu.activeMovementId } : null);
      var isOut = Boolean(activeMov);
      var isSelected = !isOut && state.selectedStudentId === stu.StudentID;
      var itemCls = 'student-roster-item' + (isSelected ? ' selected' : '') + (isOut ? ' disabled' : '');

      rowsHtml.push(
        '<div class="' + itemCls + '" role="option" aria-selected="' + (isSelected ? 'true' : 'false') + '" data-student-id="' + SMT_UTILS.escapeHtml(stu.StudentID) + '" data-currently-out="' + (isOut ? '1' : '0') + '">' +
        '  <div>' +
        '    <strong>' + SMT_UTILS.escapeHtml(stu.StudentName) + '</strong> ' +
        '    <span style="color: var(--text-muted); font-size: 0.82rem;">(' + SMT_UTILS.escapeHtml(stu.StudentID) + ')</span>' +
        '  </div>' +
        '  <div>' +
        (isOut
          ? '<span class="badge badge-warning">Already OUT</span>'
          : (isSelected ? '<span class="badge badge-active">Selected</span>' : '<span style="font-size: 0.8rem; color: var(--brand-primary); font-weight: 600;">Select</span>')) +
        '  </div>' +
        '</div>'
      );
    }

    els.rosterContainer.innerHTML = rowsHtml.join('');
    updateMarkOutButtonState();
  }

  function updateMarkOutButtonState() {
    var session = SMT_API.getSession('teacher');
    var stu = null;
    for (var i = 0; i < state.studentsInClass.length; i++) {
      if (state.studentsInClass[i].StudentID === state.selectedStudentId) {
        stu = state.studentsInClass[i];
        break;
      }
    }
    if (stu) {
      els.selectedStudentSummary.innerHTML = 'Selected: <strong>' + SMT_UTILS.escapeHtml(stu.StudentName) + '</strong> (' + SMT_UTILS.escapeHtml(stu.StudentID) + ')';
    } else {
      els.selectedStudentSummary.textContent = 'No student selected';
    }

    var canSubmit = Boolean(
      session &&
      session.token &&
      state.selectedClassId &&
      state.selectedDestination &&
      state.selectedStudentId &&
      !state.submittingOut
    );
    els.btnMarkOut.disabled = !canSubmit;
  }

  function renderCurrentOutTable() {
    els.currentOutCount.textContent = String(state.currentOutRecords.length);
    if (state.currentOutRecords.length === 0) {
      els.currentOutTbody.innerHTML = '<tr><td colspan="8" class="empty-state">No students are currently outside the classroom.</td></tr>';
      return;
    }

    var nowMs = Date.now();
    var html = [];
    for (var i = 0; i < state.currentOutRecords.length; i++) {
      var rec = state.currentOutRecords[i];
      var liveMetrics = SMT_UTILS.calculateMovementStatus(
        rec.TimeOut,
        null,
        rec.MovementID,
        state.warningMinutes,
        state.criticalMinutes,
        nowMs
      );
      var rowClass = '';
      if (liveMetrics.attentionLevel === 'Critical') rowClass = 'row-critical';
      else if (liveMetrics.attentionLevel === 'Warning') rowClass = 'row-warning';

      var isReturning = Boolean(state.returningIds[rec.MovementID]);

      html.push(
        '<tr class="' + rowClass + '" data-movement-id="' + SMT_UTILS.escapeHtml(rec.MovementID) + '">' +
        '  <td><strong>' + SMT_UTILS.escapeHtml(rec.StudentName) + '</strong><br><small style="color: var(--text-muted);">' + SMT_UTILS.escapeHtml(rec.StudentID) + '</small></td>' +
        '  <td>' + SMT_UTILS.escapeHtml(rec.ClassName) + '</td>' +
        '  <td>' + SMT_UTILS.escapeHtml(rec.Destination) + '</td>' +
        '  <td>' + SMT_UTILS.escapeHtml(SMT_UTILS.formatTimeShort(rec.TimeOut, rec.MovementID)) + '</td>' +
        '  <td><strong>' + SMT_UTILS.escapeHtml(liveMetrics.formattedDuration) + '</strong></td>' +
        '  <td>' + SMT_UTILS.attentionBadgeHtml(liveMetrics.attentionLevel) + '</td>' +
        '  <td>' + SMT_UTILS.escapeHtml(rec.TeacherName) + '</td>' +
        '  <td>' +
        '    <button type="button" class="btn btn-success btn-sm btn-return-student" ' +
        '      data-movement-id="' + SMT_UTILS.escapeHtml(rec.MovementID) + '" ' +
        '      data-student-id="' + SMT_UTILS.escapeHtml(rec.StudentID) + '" ' +
        '      data-student-name="' + SMT_UTILS.escapeHtml(rec.StudentName) + '"' +
        (isReturning ? ' disabled' : '') + '>' +
        (isReturning ? 'Saving...' : 'Mark RETURNED') +
        '    </button>' +
        '  </td>' +
        '</tr>'
      );
    }
    els.currentOutTbody.innerHTML = html.join('');
  }

  async function loadBootstrapOptions() {
    els.authFeedback.textContent = '';
    try {
      var res = await SMT_API.get('getTeacherLoginOptions');
      var data = res.data || res || {};
      state.teachers = data.teachers || res.teachers || [];
      state.destinations = data.destinations || res.destinations || (data.settings && data.settings.destinations) || [];
      state.warningMinutes = Number(data.warningMinutes || data.warning_minutes || (data.settings && data.settings.warning_minutes) || 10);
      state.criticalMinutes = Number(data.criticalMinutes || data.critical_minutes || (data.settings && data.settings.critical_minutes) || 15);
      var schoolName = data.schoolName || data.school_name || (data.settings && data.settings.school_name);
      if (schoolName) {
        els.schoolName.textContent = schoolName + ' — Movement Tracker';
      }
      renderTeacherSelectOptions();
      renderDestinationControls();
    } catch (err) {
      els.authFeedback.textContent = err.message || 'Unable to load teacher directory.';
    }
  }

  async function loadAuthenticatedWorkspaceData() {
    var session = SMT_API.getSession('teacher');
    if (!session || !session.token) {
      lockOrClearTeacherSession();
      return;
    }

    els.sessionPill.textContent = 'Teacher: ' + session.teacherName + ' (' + session.teacherId + ')';
    els.activeTeacherLabel.textContent = session.teacherName + ' (' + session.teacherId + ')';
    els.btnLockSession.style.display = 'inline-flex';
    els.authSection.style.display = 'none';
    els.workspace.style.display = 'grid';

    try {
      var results = await Promise.all([
        SMT_API.get('getClasses', {}, { role: 'teacher' }),
        SMT_API.get('getCurrentOut', {}, { role: 'teacher' })
      ]);
      state.classes = (results[0].data && results[0].data.classes) || [];
      state.currentOutRecords = (results[1].data && results[1].data.records) || [];
      els.lastRefreshTime.textContent = new Date().toLocaleTimeString();
      renderClassOptions();
      renderCurrentOutTable();
      if (state.selectedClassId) {
        await loadStudentsForSelectedClass(state.selectedClassId);
      } else {
        renderStudentRoster();
      }
    } catch (err) {
      SMT_UTILS.showToast(err.message || 'Failed loading classroom data', 'error');
      if (!SMT_API.getSession('teacher')) {
        lockOrClearTeacherSession();
      }
    }
  }

  async function loadStudentsForSelectedClass(classId) {
    if (!classId) {
      state.studentsInClass = [];
      state.selectedStudentId = '';
      renderStudentRoster();
      return;
    }
    try {
      var res = await SMT_API.get('getStudents', { classId: classId }, { role: 'teacher', silentLoading: true });
      state.studentsInClass = (res.data && res.data.students) || [];
      renderStudentRoster();
    } catch (err) {
      SMT_UTILS.showToast(err.message || 'Could not load students for class', 'error');
    }
  }

  async function refreshCurrentOutOnly(silent) {
    var session = SMT_API.getSession('teacher');
    if (!session || !session.token) return;
    try {
      var res = await SMT_API.get('getCurrentOut', {}, { role: 'teacher', silentLoading: Boolean(silent) });
      state.currentOutRecords = (res.data && res.data.records) || [];
      els.lastRefreshTime.textContent = new Date().toLocaleTimeString();
      renderCurrentOutTable();
      renderStudentRoster();
    } catch (err) {
      if (!silent) {
        SMT_UTILS.showToast(err.message || 'Could not refresh OUT records', 'error');
      }
      if (!SMT_API.getSession('teacher')) {
        lockOrClearTeacherSession();
      }
    }
  }

  /**
   * Clears the previous teacher's session and requires re-authentication.
   */
  function lockOrClearTeacherSession(preselectTeacherId) {
    SMT_API.clearSession('teacher');
    state.selectedStudentId = '';
    els.sessionPill.textContent = 'Locked (No Active Teacher)';
    els.btnLockSession.style.display = 'none';
    els.workspace.style.display = 'none';
    els.authSection.style.display = 'block';
    els.inputTeacherPin.value = '';
    if (preselectTeacherId) {
      els.selectLoginTeacher.value = preselectTeacherId;
      els.inputTeacherPin.focus();
    }
    updateMarkOutButtonState();
  }

  function bindEvents() {
    SMT_API.onConnectionStateChange(updateConnectionIndicator);

    window.addEventListener('online', function () {
      refreshCurrentOutOnly(true);
    });
    window.addEventListener('offline', function () {
      updateConnectionIndicator({ status: 'offline', message: 'Offline' });
    });

    els.endpointForm.addEventListener('submit', function (e) {
      e.preventDefault();
      var url = String(els.inputWebAppUrl.value || '').trim();
      SMT_CONFIG.setWebAppUrl(url);
      loadBootstrapOptions();
    });

    els.btnReloadLogin.addEventListener('click', function () {
      loadBootstrapOptions();
    });

    // Clearing any existing session immediately if user changes teacher in login dropdown
    els.selectLoginTeacher.addEventListener('change', function () {
      SMT_API.clearSession('teacher');
      els.authFeedback.textContent = '';
    });

    // Switching teachers inside the workspace immediately clears the previous teacher's session
    els.selectSwitchTeacher.addEventListener('change', function () {
      var nextTeacherId = els.selectSwitchTeacher.value;
      els.selectSwitchTeacher.value = '';
      lockOrClearTeacherSession(nextTeacherId);
      SMT_UTILS.showToast('Previous teacher session cleared. Please enter PIN for the selected teacher.', 'info');
    });

    els.btnLockSession.addEventListener('click', function () {
      lockOrClearTeacherSession();
      SMT_UTILS.showToast('Teacher session locked.', 'info');
    });

    els.loginForm.addEventListener('submit', async function (e) {
      e.preventDefault();
      var teacherId = els.selectLoginTeacher.value;
      var pin = els.inputTeacherPin.value;
      if (!teacherId || !pin) {
        els.authFeedback.textContent = 'Please select your profile and enter the Teacher App PIN.';
        return;
      }
      els.btnTeacherLogin.disabled = true;
      els.authFeedback.textContent = '';
      try {
        // Clear any prior teacher session before authenticating
        SMT_API.clearSession('teacher');
        var res = await SMT_API.post('verifyTeacherPasscode', {
          teacherId: teacherId,
          pin: pin
        });
        SMT_API.setSession('teacher', res.data);
        els.inputTeacherPin.value = '';
        SMT_UTILS.showToast('Signed in as ' + res.data.teacherName, 'success');
        await loadAuthenticatedWorkspaceData();
      } catch (err) {
        els.authFeedback.textContent = err.message || 'Authentication failed.';
      } finally {
        els.btnTeacherLogin.disabled = false;
      }
    });

    els.selectClass.addEventListener('change', function () {
      state.selectedClassId = els.selectClass.value;
      state.selectedStudentId = '';
      loadStudentsForSelectedClass(state.selectedClassId);
    });

    els.selectDestination.addEventListener('change', function () {
      state.selectedDestination = els.selectDestination.value;
      renderDestinationControls();
      updateMarkOutButtonState();
    });

    els.destChipsContainer.addEventListener('click', function (e) {
      var btn = e.target.closest('.dest-chip');
      if (!btn) return;
      state.selectedDestination = btn.getAttribute('data-destination') || '';
      renderDestinationControls();
      updateMarkOutButtonState();
    });

    els.inputStudentSearch.addEventListener('input', function () {
      renderStudentRoster();
    });

    els.rosterContainer.addEventListener('click', function (e) {
      var item = e.target.closest('.student-roster-item');
      if (!item) return;
      if (item.getAttribute('data-currently-out') === '1') {
        SMT_UTILS.showToast('This student already has an active OUT movement.', 'warning');
        return;
      }
      var stuId = item.getAttribute('data-student-id') || '';
      state.selectedStudentId = stuId;
      els.selectedStudentIdInput.value = stuId;
      renderStudentRoster();
    });

    els.recordOutForm.addEventListener('submit', async function (e) {
      e.preventDefault();
      if (state.submittingOut) return;
      var session = SMT_API.getSession('teacher');
      if (!session || !session.token) {
        lockOrClearTeacherSession();
        return;
      }
      if (!state.selectedClassId || !state.selectedDestination || !state.selectedStudentId) {
        SMT_UTILS.showToast('Select class, destination, and student first.', 'warning');
        return;
      }
      state.submittingOut = true;
      updateMarkOutButtonState();
      try {
        var res = await SMT_API.post('recordOut', {
          studentId: state.selectedStudentId,
          classId: state.selectedClassId,
          destination: state.selectedDestination,
          notes: els.inputNotes.value
        }, { role: 'teacher' });
        SMT_UTILS.showToast(res.message || 'Student marked OUT', 'success');
        state.selectedStudentId = '';
        els.inputNotes.value = '';
        await Promise.all([
          refreshCurrentOutOnly(true),
          loadStudentsForSelectedClass(state.selectedClassId)
        ]);
      } catch (err) {
        SMT_UTILS.showToast(err.message || 'Failed to record movement OUT', 'error');
      } finally {
        state.submittingOut = false;
        updateMarkOutButtonState();
      }
    });

    els.currentOutTbody.addEventListener('click', async function (e) {
      var btn = e.target.closest('.btn-return-student');
      if (!btn) return;
      var movementId = btn.getAttribute('data-movement-id');
      var studentId = btn.getAttribute('data-student-id');
      if (!movementId || state.returningIds[movementId]) return;

      state.returningIds[movementId] = true;
      renderCurrentOutTable();
      try {
        var res = await SMT_API.post('recordReturn', {
          movementId: movementId,
          studentId: studentId
        }, { role: 'teacher' });
        SMT_UTILS.showToast(res.message || 'Student marked RETURNED', 'success');
        delete state.returningIds[movementId];
        await Promise.all([
          refreshCurrentOutOnly(true),
          state.selectedClassId ? loadStudentsForSelectedClass(state.selectedClassId) : Promise.resolve()
        ]);
      } catch (err) {
        delete state.returningIds[movementId];
        renderCurrentOutTable();
        SMT_UTILS.showToast(err.message || 'Failed to record return', 'error');
      }
    });

    els.btnManualRefresh.addEventListener('click', function () {
      refreshCurrentOutOnly(false);
    });
  }

  function startTimers() {
    if (state.refreshTimer) clearInterval(state.refreshTimer);
    state.refreshTimer = setInterval(function () {
      if (SMT_API.getSession('teacher')) {
        refreshCurrentOutOnly(true);
      }
    }, SMT_CONFIG.AUTO_REFRESH_INTERVAL_MS || 10000);

    if (state.durationTickTimer) clearInterval(state.durationTickTimer);
    state.durationTickTimer = setInterval(function () {
      if (state.currentOutRecords.length > 0 && SMT_API.getSession('teacher')) {
        renderCurrentOutTable();
      }
    }, 1000);
  }

  async function init() {
    cacheElements();
    bindEvents();
    var currentUrl = SMT_CONFIG.getWebAppUrl();
    els.inputWebAppUrl.value = currentUrl;
    if (currentUrl) {
      await loadBootstrapOptions();
      var existingSession = SMT_API.getSession('teacher');
      if (existingSession && existingSession.token) {
        await loadAuthenticatedWorkspaceData();
      }
    }
    startTimers();
  }

  if (typeof document !== 'undefined') {
    document.addEventListener('DOMContentLoaded', init);
  }
})();
