/**
 * Student Movement Tracker — Principal Dashboard Controller (js/principal-dashboard.js)
 */
(function () {
  var state = {
    allOutRecords: [],
    classes: [],
    destinations: [],
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
    els.sessionPill = document.getElementById('principal-session-pill');
    els.btnLogout = document.getElementById('btn-logout-principal');

    els.endpointForm = document.getElementById('endpoint-config-form');
    els.inputWebAppUrl = document.getElementById('input-webapp-url');

    els.authSection = document.getElementById('principal-auth-section');
    els.loginForm = document.getElementById('principal-login-form');
    els.inputPasscode = document.getElementById('input-principal-passcode');
    els.btnLogin = document.getElementById('btn-principal-login');
    els.authFeedback = document.getElementById('principal-auth-feedback');

    els.workspace = document.getElementById('principal-workspace');
    els.statTotalOut = document.getElementById('stat-total-out');
    els.statAffectedClasses = document.getElementById('stat-affected-classes');
    els.statOccupiedDest = document.getElementById('stat-occupied-destinations');
    els.statOver10 = document.getElementById('stat-over-10');
    els.statOver15 = document.getElementById('stat-over-15');

    els.lastUpdate = document.getElementById('principal-last-update');
    els.btnRefresh = document.getElementById('btn-principal-refresh');
    els.filterClass = document.getElementById('filter-class');
    els.filterDestination = document.getElementById('filter-destination');
    els.filterAttention = document.getElementById('filter-attention');
    els.filterSearch = document.getElementById('filter-search');
    els.tbody = document.getElementById('principal-movements-tbody');
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

  function renderFilterOptions() {
    var currentClass = els.filterClass.value;
    var classHtml = ['<option value="">All Classes</option>'];
    for (var i = 0; i < state.classes.length; i++) {
      var c = state.classes[i];
      classHtml.push(
        '<option value="' + SMT_UTILS.escapeHtml(c.ClassID) + '"' +
        (currentClass === c.ClassID ? ' selected' : '') + '>' +
        SMT_UTILS.escapeHtml(c.ClassName) + '</option>'
      );
    }
    els.filterClass.innerHTML = classHtml.join('');

    var currentDest = els.filterDestination.value;
    var destHtml = ['<option value="">All Destinations</option>'];
    for (var j = 0; j < state.destinations.length; j++) {
      var d = state.destinations[j];
      destHtml.push(
        '<option value="' + SMT_UTILS.escapeHtml(d) + '"' +
        (currentDest === d ? ' selected' : '') + '>' +
        SMT_UTILS.escapeHtml(d) + '</option>'
      );
    }
    els.filterDestination.innerHTML = destHtml.join('');
  }

  function computeAndRenderLiveView() {
    var nowMs = Date.now();
    var totalOut = state.allOutRecords.length;
    var classesSet = {};
    var destSet = {};
    var over10Count = 0;
    var over15Count = 0;

    var enriched = [];
    for (var i = 0; i < state.allOutRecords.length; i++) {
      var r = state.allOutRecords[i];
      var m = SMT_UTILS.calculateMovementStatus(
        r.TimeOut,
        null,
        r.MovementID,
        state.warningMinutes,
        state.criticalMinutes,
        nowMs
      );
      if (r.ClassID) classesSet[r.ClassID] = true;
      if (r.Destination) destSet[r.Destination] = true;
      if (m.elapsedMinutes >= state.warningMinutes) over10Count++;
      if (m.elapsedMinutes >= state.criticalMinutes) over15Count++;

      enriched.push(Object.assign({}, r, {
        elapsedSeconds: m.elapsedSeconds,
        elapsedMinutes: m.elapsedMinutes,
        formattedDuration: m.formattedDuration,
        attentionLevel: m.attentionLevel
      }));
    }

    els.statTotalOut.textContent = String(totalOut);
    els.statAffectedClasses.textContent = String(Object.keys(classesSet).length);
    els.statOccupiedDest.textContent = String(Object.keys(destSet).length);
    els.statOver10.textContent = String(over10Count);
    els.statOver15.textContent = String(over15Count);

    var fClass = els.filterClass.value;
    var fDest = els.filterDestination.value;
    var fAttn = els.filterAttention.value;
    var fSearch = String(els.filterSearch.value || '').trim().toLowerCase();

    var filtered = enriched.filter(function (item) {
      if (fClass && item.ClassID !== fClass) return false;
      if (fDest && item.Destination !== fDest) return false;
      if (fAttn && item.attentionLevel !== fAttn) return false;
      if (fSearch) {
        if (
          item.StudentName.toLowerCase().indexOf(fSearch) === -1 &&
          item.StudentID.toLowerCase().indexOf(fSearch) === -1
        ) {
          return false;
        }
      }
      return true;
    });

    if (filtered.length === 0) {
      els.tbody.innerHTML = '<tr><td colspan="9" class="empty-state">No active student movements match the current filters.</td></tr>';
      return;
    }

    var rows = [];
    for (var k = 0; k < filtered.length; k++) {
      var rec = filtered[k];
      var rowCls = '';
      if (rec.attentionLevel === 'Critical') rowCls = 'row-critical';
      else if (rec.attentionLevel === 'Warning') rowCls = 'row-warning';

      rows.push(
        '<tr class="' + rowCls + '">' +
        '  <td><strong>' + SMT_UTILS.escapeHtml(rec.StudentName) + '</strong></td>' +
        '  <td>' + SMT_UTILS.escapeHtml(rec.StudentID) + '</td>' +
        '  <td>' + SMT_UTILS.escapeHtml(rec.ClassName) + '</td>' +
        '  <td>' + SMT_UTILS.escapeHtml(rec.TeacherName) + '</td>' +
        '  <td>' + SMT_UTILS.escapeHtml(rec.Destination) + '</td>' +
        '  <td>' + SMT_UTILS.escapeHtml(SMT_UTILS.formatTimeShort(rec.TimeOut, rec.MovementID)) + '</td>' +
        '  <td><strong>' + SMT_UTILS.escapeHtml(rec.formattedDuration) + '</strong></td>' +
        '  <td>' + SMT_UTILS.attentionBadgeHtml(rec.attentionLevel) + '</td>' +
        '  <td><span class="badge badge-warning">' + SMT_UTILS.escapeHtml(rec.Status) + '</span></td>' +
        '</tr>'
      );
    }
    els.tbody.innerHTML = rows.join('');
  }

  async function loadDashboardData(silent) {
    var session = SMT_API.getSession('principal');
    if (!session || !session.token) {
      lockPrincipalSession();
      return;
    }

    els.sessionPill.textContent = 'Principal Session Active';
    els.btnLogout.style.display = 'inline-flex';
    els.authSection.style.display = 'none';
    els.workspace.style.display = 'block';

    try {
      var results = await Promise.all([
        SMT_API.get('getSettings', {}, { silentLoading: Boolean(silent) }),
        SMT_API.get('getClasses', {}, { role: 'principal', silentLoading: Boolean(silent) }),
        SMT_API.get('getCurrentOut', {}, { role: 'principal', silentLoading: Boolean(silent) })
      ]);

      var settings = results[0].data || {};
      state.warningMinutes = Number(settings.warningMinutes || 10);
      state.criticalMinutes = Number(settings.criticalMinutes || 15);
      state.destinations = settings.destinations || [];
      if (settings.schoolName) {
        els.schoolName.textContent = settings.schoolName + ' — Principal Monitor';
      }

      state.classes = (results[1].data && results[1].data.classes) || [];
      state.allOutRecords = (results[2].data && results[2].data.records) || [];

      els.lastUpdate.textContent = new Date().toLocaleTimeString();
      renderFilterOptions();
      computeAndRenderLiveView();
    } catch (err) {
      if (!silent) {
        SMT_UTILS.showToast(err.message || 'Failed loading Principal dashboard data', 'error');
      }
      if (!SMT_API.getSession('principal')) {
        lockPrincipalSession();
      }
    }
  }

  function lockPrincipalSession() {
    SMT_API.clearSession('principal');
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
      if (SMT_API.getSession('principal')) {
        loadDashboardData(false);
      } else {
        SMT_API.get('getSettings').then(function (res) {
          if (res.data && res.data.schoolName) {
            els.schoolName.textContent = res.data.schoolName + ' — Principal Monitor';
          }
        }).catch(function () {});
      }
    });

    els.loginForm.addEventListener('submit', async function (e) {
      e.preventDefault();
      var pass = els.inputPasscode.value;
      if (!pass) return;
      els.btnLogin.disabled = true;
      els.authFeedback.textContent = '';
      try {
        var res = await SMT_API.post('verifyPasscode', {
          role: 'principal',
          passcode: pass
        });
        SMT_API.setSession('principal', res.data);
        els.inputPasscode.value = '';
        SMT_UTILS.showToast('Signed in to Principal Dashboard', 'success');
        await loadDashboardData(false);
      } catch (err) {
        els.authFeedback.textContent = err.message || 'Invalid Principal passcode.';
      } finally {
        els.btnLogin.disabled = false;
      }
    });

    els.btnLogout.addEventListener('click', function () {
      lockPrincipalSession();
      SMT_UTILS.showToast('Principal session locked.', 'info');
    });

    els.btnRefresh.addEventListener('click', function () {
      loadDashboardData(false);
    });

    els.filterClass.addEventListener('change', computeAndRenderLiveView);
    els.filterDestination.addEventListener('change', computeAndRenderLiveView);
    els.filterAttention.addEventListener('change', computeAndRenderLiveView);
    els.filterSearch.addEventListener('input', computeAndRenderLiveView);
  }

  function startTimers() {
    if (state.refreshTimer) clearInterval(state.refreshTimer);
    state.refreshTimer = setInterval(function () {
      if (SMT_API.getSession('principal')) {
        loadDashboardData(true);
      }
    }, SMT_CONFIG.AUTO_REFRESH_INTERVAL_MS || 10000);

    if (state.tickTimer) clearInterval(state.tickTimer);
    state.tickTimer = setInterval(function () {
      if (SMT_API.getSession('principal') && state.allOutRecords.length > 0) {
        computeAndRenderLiveView();
      }
    }, 1000);
  }

  async function init() {
    cacheElements();
    bindEvents();
    var currentUrl = SMT_CONFIG.getWebAppUrl();
    els.inputWebAppUrl.value = currentUrl;
    if (currentUrl) {
      if (SMT_API.getSession('principal')) {
        await loadDashboardData(false);
      } else {
        SMT_API.get('getSettings').then(function (res) {
          var s = res.data || res.settings || res || {};
          var name = s.schoolName || s.school_name;
          if (name) {
            els.schoolName.textContent = name + ' — Principal Monitor';
          }
        }).catch(function () {});
      }
    }
    startTimers();
  }

  if (typeof document !== 'undefined') {
    document.addEventListener('DOMContentLoaded', init);
  }
})();
