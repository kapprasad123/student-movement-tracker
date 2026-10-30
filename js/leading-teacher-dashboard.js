/**
 * Student Movement Tracker — Leading Teacher Dashboard Controller (js/leading-teacher-dashboard.js)
 */
(function () {
  var state = {
    leadingTeachers: [],
    assignedClasses: [],
    destinations: [],
    liveOutRecords: [],
    historyRecords: [],
    warningMinutes: 10,
    criticalMinutes: 15,
    refreshTimer: null,
    tickTimer: null
  };

  var els = {};

  function cacheElements() {
    els.schoolName = document.getElementById('header-school-name');
    els.connPill = document.getElementById('connection-status-pill');
    els.connText = document.getElementById('connection-status-text');
    els.sessionPill = document.getElementById('lt-session-pill');
    els.btnLogout = document.getElementById('btn-logout-lt');

    els.endpointForm = document.getElementById('endpoint-config-form');
    els.inputWebAppUrl = document.getElementById('input-webapp-url');

    els.authSection = document.getElementById('lt-auth-section');
    els.loginForm = document.getElementById('lt-login-form');
    els.selectProfile = document.getElementById('select-lt-profile');
    els.inputPasscode = document.getElementById('input-lt-passcode');
    els.btnLogin = document.getElementById('btn-lt-login');
    els.btnReloadOptions = document.getElementById('btn-reload-lt-options');
    els.authFeedback = document.getElementById('lt-auth-feedback');

    els.workspace = document.getElementById('lt-workspace');
    els.profileName = document.getElementById('lt-profile-name');
    els.assignedBadges = document.getElementById('lt-assigned-classes-badges');
    els.lastUpdate = document.getElementById('lt-last-update');

    els.statAssignedClasses = document.getElementById('lt-stat-assigned-classes');
    els.statCurrentOut = document.getElementById('lt-stat-current-out');
    els.statWarning = document.getElementById('lt-stat-warning');
    els.statCritical = document.getElementById('lt-stat-critical');

    els.btnRefresh = document.getElementById('btn-lt-refresh');
    els.filterClass = document.getElementById('lt-filter-class');
    els.filterDestination = document.getElementById('lt-filter-destination');
    els.filterAttention = document.getElementById('lt-filter-attention');
    els.filterSearch = document.getElementById('lt-filter-search');
    els.liveTbody = document.getElementById('lt-live-tbody');

    els.histRange = document.getElementById('lt-hist-range');
    els.histClass = document.getElementById('lt-hist-class');
    els.histStatus = document.getElementById('lt-hist-status');
    els.histSearch = document.getElementById('lt-hist-search');
    els.btnLoadHistory = document.getElementById('btn-lt-load-history');
    els.btnExportCsv = document.getElementById('btn-lt-export-csv');
    els.historyTbody = document.getElementById('lt-history-tbody');
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
  }

  async function loadLoginOptions() {
    els.authFeedback.textContent = '';
    try {
      var res = await SMT_API.get('getLeadingTeacherLoginOptions');
      var data = res.data || {};
      state.leadingTeachers = data.leadingTeachers || [];
      if (data.schoolName) {
        els.schoolName.textContent = data.schoolName + ' — Leading Teacher';
      }
      var opts = ['<option value="">— Select Leading Teacher —</option>'];
      for (var i = 0; i < state.leadingTeachers.length; i++) {
        var lt = state.leadingTeachers[i];
        opts.push(
          '<option value="' + SMT_UTILS.escapeHtml(lt.LeadingTeacherID) + '">' +
          SMT_UTILS.escapeHtml(lt.LeadingTeacherName) + ' (' + SMT_UTILS.escapeHtml(lt.LeadingTeacherID) + ')</option>'
        );
      }
      els.selectProfile.innerHTML = opts.join('');
    } catch (err) {
      els.authFeedback.textContent = err.message || 'Could not load Leading Teacher profiles.';
    }
  }

  function renderAssignedClassesUI() {
    els.statAssignedClasses.textContent = String(state.assignedClasses.length);
    if (state.assignedClasses.length === 0) {
      els.assignedBadges.innerHTML = '<span class="badge badge-warning">No active classes assigned</span>';
    } else {
      els.assignedBadges.innerHTML = state.assignedClasses.map(function (c) {
        return '<span class="badge badge-active" style="margin-right: 0.35rem;">' +
          SMT_UTILS.escapeHtml(c.ClassName) + ' (' + SMT_UTILS.escapeHtml(c.ClassID) + ')</span>';
      }).join('');
    }

    var curLiveClass = els.filterClass.value;
    var curHistClass = els.histClass.value;
    var liveOpts = ['<option value="">All Assigned Classes</option>'];
    var histOpts = ['<option value="">All Assigned Classes</option>'];
    for (var i = 0; i < state.assignedClasses.length; i++) {
      var cls = state.assignedClasses[i];
      liveOpts.push(
        '<option value="' + SMT_UTILS.escapeHtml(cls.ClassID) + '"' +
        (curLiveClass === cls.ClassID ? ' selected' : '') + '>' +
        SMT_UTILS.escapeHtml(cls.ClassName) + '</option>'
      );
      histOpts.push(
        '<option value="' + SMT_UTILS.escapeHtml(cls.ClassID) + '"' +
        (curHistClass === cls.ClassID ? ' selected' : '') + '>' +
        SMT_UTILS.escapeHtml(cls.ClassName) + '</option>'
      );
    }
    els.filterClass.innerHTML = liveOpts.join('');
    els.histClass.innerHTML = histOpts.join('');

    var curDest = els.filterDestination.value;
    var destOpts = ['<option value="">All Destinations</option>'];
    for (var j = 0; j < state.destinations.length; j++) {
      var d = state.destinations[j];
      destOpts.push(
        '<option value="' + SMT_UTILS.escapeHtml(d) + '"' +
        (curDest === d ? ' selected' : '') + '>' +
        SMT_UTILS.escapeHtml(d) + '</option>'
      );
    }
    els.filterDestination.innerHTML = destOpts.join('');
  }

  function renderLiveTable() {
    var nowMs = Date.now();
    var warnCount = 0;
    var critCount = 0;
    var enriched = [];

    for (var i = 0; i < state.liveOutRecords.length; i++) {
      var r = state.liveOutRecords[i];
      var m = SMT_UTILS.calculateMovementStatus(
        r.TimeOut,
        null,
        r.MovementID,
        state.warningMinutes,
        state.criticalMinutes,
        nowMs
      );
      if (m.attentionLevel === 'Warning') warnCount++;
      if (m.attentionLevel === 'Critical') critCount++;
      enriched.push(Object.assign({}, r, {
        elapsedSeconds: m.elapsedSeconds,
        elapsedMinutes: m.elapsedMinutes,
        formattedDuration: m.formattedDuration,
        attentionLevel: m.attentionLevel
      }));
    }

    els.statCurrentOut.textContent = String(state.liveOutRecords.length);
    els.statWarning.textContent = String(warnCount);
    els.statCritical.textContent = String(critCount);

    var fClass = els.filterClass.value;
    var fDest = els.filterDestination.value;
    var fAttn = els.filterAttention.value;
    var fSearch = String(els.filterSearch.value || '').trim().toLowerCase();

    var filtered = enriched.filter(function (rec) {
      if (fClass && rec.ClassID !== fClass) return false;
      if (fDest && rec.Destination !== fDest) return false;
      if (fAttn && rec.attentionLevel !== fAttn) return false;
      if (fSearch) {
        if (
          rec.StudentName.toLowerCase().indexOf(fSearch) === -1 &&
          rec.StudentID.toLowerCase().indexOf(fSearch) === -1
        ) {
          return false;
        }
      }
      return true;
    });

    if (filtered.length === 0) {
      els.liveTbody.innerHTML = '<tr><td colspan="8" class="empty-state">No currently OUT students in your assigned classes match the selected filters.</td></tr>';
      return;
    }

    var html = [];
    for (var k = 0; k < filtered.length; k++) {
      var item = filtered[k];
      var rowCls = '';
      if (item.attentionLevel === 'Critical') rowCls = 'row-critical';
      else if (item.attentionLevel === 'Warning') rowCls = 'row-warning';

      html.push(
        '<tr class="' + rowCls + '">' +
        '  <td data-label="Student Name"><strong>' + SMT_UTILS.escapeHtml(item.StudentName) + '</strong></td>' +
        '  <td data-label="Student ID">' + SMT_UTILS.escapeHtml(item.StudentID) + '</td>' +
        '  <td data-label="Class">' + SMT_UTILS.escapeHtml(item.ClassName) + '</td>' +
        '  <td data-label="Destination">' + SMT_UTILS.escapeHtml(item.Destination) + '</td>' +
        '  <td data-label="Teacher">' + SMT_UTILS.escapeHtml(item.TeacherName) + '</td>' +
        '  <td data-label="Checkout Time">' + SMT_UTILS.escapeHtml(SMT_UTILS.formatTimeShort(item.TimeOut, item.MovementID)) + '</td>' +
        '  <td data-label="Elapsed"><strong>' + SMT_UTILS.escapeHtml(item.formattedDuration) + '</strong></td>' +
        '  <td data-label="Attention">' + SMT_UTILS.attentionBadgeHtml(item.attentionLevel) + '</td>' +
        '</tr>'
      );
    }
    els.liveTbody.innerHTML = html.join('');
  }

  async function loadHistoryForAssignedClasses() {
    var session = SMT_API.getSession('leading_teacher');
    if (!session || !session.token) return;
    try {
      var res = await SMT_API.get('getMovementHistory', {
        dateRange: els.histRange.value,
        classId: els.histClass.value,
        status: els.histStatus.value,
        search: els.histSearch.value
      }, { role: 'leading_teacher' });
      state.historyRecords = (res.data && res.data.records) || [];
      renderHistoryTable();
    } catch (err) {
      SMT_UTILS.showToast(err.message || 'Could not load movement history', 'error');
    }
  }

  function renderHistoryTable() {
    if (state.historyRecords.length === 0) {
      els.historyTbody.innerHTML = '<tr><td colspan="9" class="empty-state">No movement history records found for your assigned classes.</td></tr>';
      return;
    }
    var rows = [];
    for (var i = 0; i < state.historyRecords.length; i++) {
      var r = state.historyRecords[i];
      var rowCls = r.attentionLevel === 'Critical' ? 'row-critical' : (r.attentionLevel === 'Warning' ? 'row-warning' : '');
      rows.push(
        '<tr class="' + rowCls + '">' +
        '  <td data-label="Movement ID"><small>' + SMT_UTILS.escapeHtml(r.MovementID) + '</small></td>' +
        '  <td data-label="Student"><strong>' + SMT_UTILS.escapeHtml(r.StudentName) + '</strong> (' + SMT_UTILS.escapeHtml(r.StudentID) + ')</td>' +
        '  <td data-label="Class">' + SMT_UTILS.escapeHtml(r.ClassName) + '</td>' +
        '  <td data-label="Destination">' + SMT_UTILS.escapeHtml(r.Destination) + '</td>' +
        '  <td data-label="Checkout">' + SMT_UTILS.escapeHtml(SMT_UTILS.formatDateTime(r.TimeOut, r.MovementID)) + '</td>' +
        '  <td data-label="Returned">' + (r.TimeIn ? SMT_UTILS.escapeHtml(SMT_UTILS.formatDateTime(r.TimeIn)) : '—') + '</td>' +
        '  <td data-label="Duration"><strong>' + SMT_UTILS.escapeHtml(r.formattedDuration) + '</strong></td>' +
        '  <td data-label="Attention">' + SMT_UTILS.attentionBadgeHtml(r.attentionLevel) + '</td>' +
        '  <td data-label="Status"><span class="badge ' + (r.Status === 'OUT' ? 'badge-warning' : 'badge-active') + '">' + SMT_UTILS.escapeHtml(r.Status) + '</span></td>' +
        '</tr>'
      );
    }
    els.historyTbody.innerHTML = rows.join('');
  }

  async function loadWorkspaceData(silent) {
    var session = SMT_API.getSession('leading_teacher');
    if (!session || !session.token) {
      lockSession();
      return;
    }

    els.sessionPill.textContent = 'LT: ' + (session.leadingTeacherName || session.leadingTeacherId);
    els.profileName.textContent = (session.leadingTeacherName || '') + ' (' + (session.leadingTeacherId || '') + ')';
    els.btnLogout.style.display = 'inline-flex';
    els.authSection.style.display = 'none';
    els.workspace.style.display = 'block';

    try {
      var results = await Promise.all([
        SMT_API.get('getSettings', {}, { silentLoading: Boolean(silent) }),
        SMT_API.get('getClasses', {}, { role: 'leading_teacher', silentLoading: Boolean(silent) }),
        SMT_API.get('getCurrentOut', {}, { role: 'leading_teacher', silentLoading: Boolean(silent) })
      ]);
      var settings = results[0].data || {};
      state.warningMinutes = Number(settings.warningMinutes || 10);
      state.criticalMinutes = Number(settings.criticalMinutes || 15);
      state.destinations = settings.destinations || [];
      state.assignedClasses = (results[1].data && results[1].data.classes) || [];
      state.liveOutRecords = (results[2].data && results[2].data.records) || [];

      els.lastUpdate.textContent = new Date().toLocaleTimeString();
      renderAssignedClassesUI();
      renderLiveTable();
    } catch (err) {
      if (!silent) {
        SMT_UTILS.showToast(err.message || 'Failed loading Leading Teacher data', 'error');
      }
      if (!SMT_API.getSession('leading_teacher')) {
        lockSession();
      }
    }
  }

  function lockSession() {
    SMT_API.clearSession('leading_teacher');
    els.sessionPill.textContent = 'Locked';
    els.btnLogout.style.display = 'none';
    els.workspace.style.display = 'none';
    els.authSection.style.display = 'block';
    els.inputPasscode.value = '';
  }

  function bindEvents() {
    SMT_API.onConnectionStateChange(updateConnectionIndicator);

    els.endpointForm.addEventListener('submit', function (e) {
      e.preventDefault();
      SMT_CONFIG.setWebAppUrl(els.inputWebAppUrl.value);
      loadLoginOptions();
    });

    els.btnReloadOptions.addEventListener('click', loadLoginOptions);

    els.loginForm.addEventListener('submit', async function (e) {
      e.preventDefault();
      var ltId = els.selectProfile.value;
      var pass = els.inputPasscode.value;
      if (!ltId || !pass) {
        els.authFeedback.textContent = 'Select your Leading Teacher profile and enter the passcode.';
        return;
      }
      els.btnLogin.disabled = true;
      els.authFeedback.textContent = '';
      try {
        SMT_API.clearSession('leading_teacher');
        var res = await SMT_API.post('verifyPasscode', {
          role: 'leading_teacher',
          leadingTeacherId: ltId,
          passcode: pass
        });
        SMT_API.setSession('leading_teacher', res.data);
        els.inputPasscode.value = '';
        SMT_UTILS.showToast('Signed in as ' + res.data.leadingTeacherName, 'success');
        await loadWorkspaceData(false);
        await loadHistoryForAssignedClasses();
      } catch (err) {
        els.authFeedback.textContent = err.message || 'Authentication failed.';
      } finally {
        els.btnLogin.disabled = false;
      }
    });

    els.btnLogout.addEventListener('click', function () {
      lockSession();
      SMT_UTILS.showToast('Leading Teacher session locked.', 'info');
    });

    els.btnRefresh.addEventListener('click', function () {
      loadWorkspaceData(false);
    });

    els.filterClass.addEventListener('change', renderLiveTable);
    els.filterDestination.addEventListener('change', renderLiveTable);
    els.filterAttention.addEventListener('change', renderLiveTable);
    els.filterSearch.addEventListener('input', renderLiveTable);

    var tabBtns = document.querySelectorAll('[data-lt-tab]');
    for (var i = 0; i < tabBtns.length; i++) {
      tabBtns[i].addEventListener('click', function (e) {
        var btn = e.currentTarget;
        var targetTab = btn.getAttribute('data-lt-tab');
        for (var j = 0; j < tabBtns.length; j++) {
          tabBtns[j].classList.remove('active');
        }
        btn.classList.add('active');
        if (btn.scrollIntoView) {
          btn.scrollIntoView({ behavior: 'smooth', inline: 'center', block: 'nearest' });
        }
        document.getElementById('lt-tab-live').classList.toggle('active', targetTab === 'live');
        document.getElementById('lt-tab-history').classList.toggle('active', targetTab === 'history');
        if (targetTab === 'history') {
          loadHistoryForAssignedClasses();
        }
      });
    }

    els.btnLoadHistory.addEventListener('click', loadHistoryForAssignedClasses);
    els.btnExportCsv.addEventListener('click', function () {
      var csv = SMT_UTILS.movementRecordsToCsv(state.historyRecords);
      SMT_UTILS.downloadCsvFile('leading-teacher-movements.csv', csv);
    });
  }

  function startTimers() {
    if (state.refreshTimer) clearInterval(state.refreshTimer);
    state.refreshTimer = setInterval(function () {
      if (SMT_API.getSession('leading_teacher')) {
        loadWorkspaceData(true);
      }
    }, SMT_CONFIG.AUTO_REFRESH_INTERVAL_MS || 10000);

    if (state.tickTimer) clearInterval(state.tickTimer);
    state.tickTimer = setInterval(function () {
      if (SMT_API.getSession('leading_teacher') && state.liveOutRecords.length > 0) {
        renderLiveTable();
      }
    }, 1000);
  }

  async function init() {
    cacheElements();
    bindEvents();
    var currentUrl = SMT_CONFIG.getWebAppUrl();
    els.inputWebAppUrl.value = currentUrl;
    if (currentUrl) {
      await loadLoginOptions();
      if (SMT_API.getSession('leading_teacher')) {
        await loadWorkspaceData(false);
        await loadHistoryForAssignedClasses();
      }
    }
    startTimers();
  }

  if (typeof document !== 'undefined') {
    document.addEventListener('DOMContentLoaded', init);
  }
})();
