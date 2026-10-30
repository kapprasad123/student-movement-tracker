/**
 * Student Movement Tracker — Admin Dashboard Controller (js/admin-dashboard.js)
 */
(function () {
  var state = {
    overview: null,
    students: [],
    classes: [],
    teachers: [],
    leadingTeachers: [],
    teachingAssignments: [],
    ltAssignments: [],
    settings: {
      schoolName: '',
      schoolTimezone: 'Indian/Maldives',
      destinations: [],
      warningMinutes: 10,
      criticalMinutes: 15,
      sessionTimeoutMinutes: 480,
      teacherSessionTimeoutMinutes: 240
    },
    securityMeta: null,
    reportData: null,
    bulkPreview: null,
    bulkTeacherPreview: null
  };

  var els = {};

  function cacheElements() {
    els.schoolName = document.getElementById('header-school-name');
    els.connPill = document.getElementById('connection-status-pill');
    els.connText = document.getElementById('connection-status-text');
    els.sessionPill = document.getElementById('admin-session-pill');
    els.btnLogout = document.getElementById('btn-logout-admin');

    els.endpointBanner = document.getElementById('endpoint-config-banner');
    els.endpointForm = document.getElementById('endpoint-config-form');
    els.inputWebAppUrl = document.getElementById('input-webapp-url');

    els.authSection = document.getElementById('admin-auth-section');
    els.loginForm = document.getElementById('admin-login-form');
    els.inputPasscode = document.getElementById('input-admin-passcode');
    els.btnLogin = document.getElementById('btn-admin-login');
    els.authFeedback = document.getElementById('admin-auth-feedback');

    els.workspace = document.getElementById('admin-workspace');
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

  function lockAdminSession() {
    SMT_API.clearSession('admin');
    els.sessionPill.textContent = 'Locked';
    els.btnLogout.style.display = 'none';
    els.workspace.style.display = 'none';
    els.authSection.style.display = 'block';
    els.inputPasscode.value = '';
  }

  /**
   * Pings the public getSettings endpoint to verify the Web App URL immediately
   * (even before the Admin logs in) so the status badge updates to "Connected to Google Sheets".
   */
  async function verifyBackendConnection(showSuccessToast) {
    var typedUrl = els.inputWebAppUrl ? String(els.inputWebAppUrl.value || '').trim() : '';
    if (typedUrl && !SMT_CONFIG.getWebAppUrl()) {
      SMT_CONFIG.setWebAppUrl(typedUrl);
    }
    if (!SMT_CONFIG.getWebAppUrl()) return false;

    try {
      var res = await SMT_API.get('getSettings', {});
      if (res && res.data) {
        state.settings = res.data;
        if (res.data.schoolName) {
          els.schoolName.textContent = res.data.schoolName + ' — Admin Console';
        }
      }
      if (showSuccessToast) {
        SMT_UTILS.showToast('Connected to Google Sheets backend!', 'success');
      }
      return true;
    } catch (err) {
      if (showSuccessToast) {
        SMT_UTILS.showToast(err.message || 'Failed connecting to Google Apps Script URL', 'error');
      }
      return false;
    }
  }

  async function loadAllAdminData() {
    var session = SMT_API.getSession('admin');
    if (!session || !session.token) {
      lockAdminSession();
      return;
    }

    els.sessionPill.textContent = 'Administrator Active';
    els.btnLogout.style.display = 'inline-flex';
    els.authSection.style.display = 'none';
    els.workspace.style.display = 'block';

    try {
      var results = await Promise.all([
        SMT_API.get('getSettings', {}),
        SMT_API.get('getAdminOverview', {}, { role: 'admin' }),
        SMT_API.get('getClasses', { includeInactive: true }, { role: 'admin' }),
        SMT_API.get('getStudents', { includeInactive: true }, { role: 'admin' }),
        SMT_API.get('getTeachers', {}, { role: 'admin' }),
        SMT_API.get('getLeadingTeachers', {}, { role: 'admin' }),
        SMT_API.get('getTeachingAssignments', {}, { role: 'admin' }),
        SMT_API.get('getLeadingTeacherAssignments', {}, { role: 'admin' }),
        SMT_API.get('getSecuritySettings', {}, { role: 'admin' })
      ]);

      state.settings = results[0].data || state.settings;
      if (state.settings.schoolName) {
        els.schoolName.textContent = state.settings.schoolName + ' — Admin Console';
      }
      state.overview = results[1].data || {};
      state.classes = (results[2].data && results[2].data.classes) || [];
      state.students = (results[3].data && results[3].data.students) || [];
      state.teachers = (results[4].data && results[4].data.teachers) || [];
      state.leadingTeachers = (results[5].data && results[5].data.leadingTeachers) || [];
      state.teachingAssignments = (results[6].data && results[6].data.assignments) || [];
      state.ltAssignments = (results[7].data && results[7].data.assignments) || [];
      state.securityMeta = results[8].data || {};

      renderOverviewTab();
      renderClassDropdowns();
      renderStudentsTab();
      renderClassesTab();
      renderTeachersTab();
      renderLeadingTeachersTab();
      renderTeachingAssignmentsTab();
      renderLtAssignmentsTab();
      renderSettingsTab();
      renderSecurityTab();
      await runMovementReport();
    } catch (err) {
      SMT_UTILS.showToast(err.message || 'Failed loading Admin data', 'error');
      if (!SMT_API.getSession('admin')) {
        lockAdminSession();
      }
    }
  }

  // 1. OVERVIEW RENDERER
  function renderOverviewTab() {
    var ov = state.overview || {};
    document.getElementById('ov-total-students').textContent = String(ov.totalStudents || 0);
    document.getElementById('ov-active-students').textContent = String(ov.activeStudents || 0);
    document.getElementById('ov-active-classes').textContent = String(ov.activeClasses || 0);
    document.getElementById('ov-active-teachers').textContent = String(ov.activeTeachers || 0);
    document.getElementById('ov-active-lt').textContent = String(ov.activeLeadingTeachers || 0);
    document.getElementById('ov-current-out').textContent = String(ov.currentOutCount || 0);
    document.getElementById('ov-critical-count').textContent = String(ov.criticalMovementsCount || 0);

    var recent = ov.recentActivity || [];
    var tbody = document.getElementById('ov-recent-tbody');
    if (recent.length === 0) {
      tbody.innerHTML = '<tr><td colspan="10" class="empty-state">No movement activity recorded yet.</td></tr>';
      return;
    }
    tbody.innerHTML = recent.map(function (r) {
      var rowCls = r.attentionLevel === 'Critical' ? 'row-critical' : (r.attentionLevel === 'Warning' ? 'row-warning' : '');
      return '<tr class="' + rowCls + '">' +
        '<td data-label="Movement ID"><small>' + SMT_UTILS.escapeHtml(r.MovementID) + '</small></td>' +
        '<td data-label="Student"><strong>' + SMT_UTILS.escapeHtml(r.StudentName) + '</strong> (' + SMT_UTILS.escapeHtml(r.StudentID) + ')</td>' +
        '<td data-label="Class">' + SMT_UTILS.escapeHtml(r.ClassName) + '</td>' +
        '<td data-label="Destination">' + SMT_UTILS.escapeHtml(r.Destination) + '</td>' +
        '<td data-label="Teacher">' + SMT_UTILS.escapeHtml(r.TeacherName) + '</td>' +
        '<td data-label="Checkout">' + SMT_UTILS.escapeHtml(SMT_UTILS.formatDateTime(r.TimeOut, r.MovementID)) + '</td>' +
        '<td data-label="Returned">' + (r.TimeIn ? SMT_UTILS.escapeHtml(SMT_UTILS.formatDateTime(r.TimeIn)) : '—') + '</td>' +
        '<td data-label="Duration"><strong>' + SMT_UTILS.escapeHtml(r.formattedDuration) + '</strong></td>' +
        '<td data-label="Attention">' + SMT_UTILS.attentionBadgeHtml(r.attentionLevel) + '</td>' +
        '<td data-label="Status"><span class="badge ' + (r.Status === 'OUT' ? 'badge-warning' : 'badge-active') + '">' + SMT_UTILS.escapeHtml(r.Status) + '</span></td>' +
        '</tr>';
    }).join('');
  }

  function renderClassDropdowns() {
    var activeClasses = state.classes.filter(function (c) {
      return c.Status.toLowerCase() === 'active';
    });

    var activeOpts = ['<option value="">— Select Class —</option>'].concat(
      activeClasses.map(function (c) {
        return '<option value="' + SMT_UTILS.escapeHtml(c.ClassID) + '">' + SMT_UTILS.escapeHtml(c.ClassName) + ' (' + SMT_UTILS.escapeHtml(c.ClassID) + ')</option>';
      })
    ).join('');

    var allFilterOpts = ['<option value="">All Classes</option>'].concat(
      state.classes.map(function (c) {
        return '<option value="' + SMT_UTILS.escapeHtml(c.ClassID) + '">' + SMT_UTILS.escapeHtml(c.ClassName) + '</option>';
      })
    ).join('');

    document.getElementById('stu-select-class').innerHTML = activeOpts;
    document.getElementById('tasg-select-class').innerHTML = activeOpts;
    document.getElementById('ltasg-select-class').innerHTML = activeOpts;
    document.getElementById('stu-filter-class').innerHTML = allFilterOpts;
    document.getElementById('rep-class').innerHTML = allFilterOpts;

    var destOpts = ['<option value="">All Destinations</option>'].concat(
      (state.settings.destinations || []).map(function (d) {
        return '<option value="' + SMT_UTILS.escapeHtml(d) + '">' + SMT_UTILS.escapeHtml(d) + '</option>';
      })
    ).join('');
    document.getElementById('rep-destination').innerHTML = destOpts;
  }

  // 2. STUDENTS & BULK IMPORT RENDERER
  function renderStudentsTab() {
    var fClass = document.getElementById('stu-filter-class').value;
    var fStatus = document.getElementById('stu-filter-status').value;
    var fSearch = String(document.getElementById('stu-filter-search').value || '').trim().toLowerCase();

    var filtered = state.students.filter(function (s) {
      if (fClass && s.ClassID !== fClass) return false;
      if (fStatus && s.Status.toLowerCase() !== fStatus.toLowerCase()) return false;
      if (fSearch) {
        if (s.StudentName.toLowerCase().indexOf(fSearch) === -1 && s.StudentID.toLowerCase().indexOf(fSearch) === -1) {
          return false;
        }
      }
      return true;
    });

    var tbody = document.getElementById('admin-students-tbody');
    if (filtered.length === 0) {
      tbody.innerHTML = '<tr><td colspan="6" class="empty-state">No students match the selected criteria.</td></tr>';
      return;
    }

    tbody.innerHTML = filtered.map(function (s) {
      var isAct = s.Status.toLowerCase() === 'active';
      return '<tr>' +
        '<td data-label="Student ID"><strong>' + SMT_UTILS.escapeHtml(s.StudentID) + '</strong></td>' +
        '<td data-label="Student Name">' + SMT_UTILS.escapeHtml(s.StudentName) + '</td>' +
        '<td data-label="Class">' + SMT_UTILS.escapeHtml(s.ClassName) + ' <small style="color:var(--text-muted)">(' + SMT_UTILS.escapeHtml(s.ClassID) + ')</small></td>' +
        '<td data-label="Status"><span class="badge ' + (isAct ? 'badge-active' : 'badge-inactive') + '">' + SMT_UTILS.escapeHtml(s.Status) + '</span></td>' +
        '<td data-label="Movement State">' + (s.isCurrentlyOut ? '<span class="badge badge-warning">Currently OUT</span>' : '<span style="color:var(--text-muted)">In Class</span>') + '</td>' +
        '<td data-label="Actions"><div>' +
        '  <button type="button" class="btn btn-secondary btn-sm btn-edit-student" data-id="' + SMT_UTILS.escapeHtml(s.StudentID) + '">Edit</button> ' +
        '  <button type="button" class="btn ' + (isAct ? 'btn-danger' : 'btn-success') + ' btn-sm btn-toggle-student" data-id="' + SMT_UTILS.escapeHtml(s.StudentID) + '" data-next="' + (isAct ? 'Inactive' : 'Active') + '">' +
        (isAct ? 'Deactivate' : 'Activate') + '</button>' +
        '</div></td>' +
        '</tr>';
    }).join('');
  }

  // 3. CLASSES RENDERER
  function renderClassesTab() {
    var search = String(document.getElementById('cls-filter-search').value || '').trim().toLowerCase();
    var filtered = state.classes.filter(function (c) {
      if (!search) return true;
      return c.ClassName.toLowerCase().indexOf(search) !== -1 || c.ClassID.toLowerCase().indexOf(search) !== -1;
    });

    var tbody = document.getElementById('admin-classes-tbody');
    if (filtered.length === 0) {
      tbody.innerHTML = '<tr><td colspan="6" class="empty-state">No classes found.</td></tr>';
      return;
    }

    tbody.innerHTML = filtered.map(function (c) {
      var isAct = c.Status.toLowerCase() === 'active';
      return '<tr>' +
        '<td data-label="Class ID"><strong>' + SMT_UTILS.escapeHtml(c.ClassID) + '</strong></td>' +
        '<td data-label="Class Name">' + SMT_UTILS.escapeHtml(c.ClassName) + '</td>' +
        '<td data-label="Active Students"><strong>' + c.activeEnrolledCount + '</strong></td>' +
        '<td data-label="Total Students">' + c.enrolledCount + '</td>' +
        '<td data-label="Status"><span class="badge ' + (isAct ? 'badge-active' : 'badge-inactive') + '">' + SMT_UTILS.escapeHtml(c.Status) + '</span></td>' +
        '<td data-label="Actions"><div>' +
        '  <button type="button" class="btn btn-secondary btn-sm btn-edit-class" data-id="' + SMT_UTILS.escapeHtml(c.ClassID) + '">Edit / Rename</button> ' +
        '  <button type="button" class="btn ' + (isAct ? 'btn-danger' : 'btn-success') + ' btn-sm btn-toggle-class" data-id="' + SMT_UTILS.escapeHtml(c.ClassID) + '" data-name="' + SMT_UTILS.escapeHtml(c.ClassName) + '" data-enrolled="' + c.activeEnrolledCount + '" data-next="' + (isAct ? 'Inactive' : 'Active') + '">' +
        (isAct ? 'Deactivate' : 'Activate') + '</button>' +
        '</div></td>' +
        '</tr>';
    }).join('');
  }

  // 4. TEACHERS & BULK TEACHER IMPORT RENDERER
  function renderTeachersTab() {
    var activeTeachers = state.teachers.filter(function (t) { return t.Status.toLowerCase() === 'active'; });
    document.getElementById('tasg-select-teacher').innerHTML = ['<option value="">— Select Teacher —</option>'].concat(
      activeTeachers.map(function (t) {
        return '<option value="' + SMT_UTILS.escapeHtml(t.TeacherID) + '">' + SMT_UTILS.escapeHtml(t.TeacherName) + ' (' + SMT_UTILS.escapeHtml(t.TeacherID) + ')</option>';
      })
    ).join('');

    var searchEl = document.getElementById('tch-filter-search');
    var search = searchEl ? String(searchEl.value || '').trim().toLowerCase() : '';
    var filtered = state.teachers.filter(function (t) {
      if (!search) return true;
      return t.TeacherName.toLowerCase().indexOf(search) !== -1 || t.TeacherID.toLowerCase().indexOf(search) !== -1;
    });

    var tbody = document.getElementById('admin-teachers-tbody');
    if (filtered.length === 0) {
      tbody.innerHTML = '<tr><td colspan="5" class="empty-state">No matching teachers found.</td></tr>';
      return;
    }

    tbody.innerHTML = filtered.map(function (t) {
      var isAct = t.Status.toLowerCase() === 'active';
      var classesBadges = (t.assignedClasses || []).map(function (ac) {
        return '<span class="badge badge-normal" style="margin-right:0.25rem;">' + SMT_UTILS.escapeHtml(ac.ClassName) + '</span>';
      }).join('') || '<span style="color:var(--text-muted)">None assigned</span>';

      return '<tr>' +
        '<td data-label="Teacher ID"><strong>' + SMT_UTILS.escapeHtml(t.TeacherID) + '</strong></td>' +
        '<td data-label="Teacher Name">' + SMT_UTILS.escapeHtml(t.TeacherName) + '</td>' +
        '<td data-label="Assigned Classes"><div>' + classesBadges + '</div></td>' +
        '<td data-label="Status"><span class="badge ' + (isAct ? 'badge-active' : 'badge-inactive') + '">' + SMT_UTILS.escapeHtml(t.Status) + '</span></td>' +
        '<td data-label="Actions"><div>' +
        '  <button type="button" class="btn btn-secondary btn-sm btn-edit-teacher" data-id="' + SMT_UTILS.escapeHtml(t.TeacherID) + '">Edit</button> ' +
        '  <button type="button" class="btn ' + (isAct ? 'btn-danger' : 'btn-success') + ' btn-sm btn-toggle-teacher" data-id="' + SMT_UTILS.escapeHtml(t.TeacherID) + '" data-next="' + (isAct ? 'Inactive' : 'Active') + '">' +
        (isAct ? 'Deactivate' : 'Activate') + '</button>' +
        '</div></td>' +
        '</tr>';
    }).join('');
  }

  // 5. LEADING TEACHERS RENDERER
  function renderLeadingTeachersTab() {
    var activeLts = state.leadingTeachers.filter(function (lt) { return lt.Status.toLowerCase() === 'active'; });
    document.getElementById('ltasg-select-lt').innerHTML = ['<option value="">— Select Leading Teacher —</option>'].concat(
      activeLts.map(function (lt) {
        return '<option value="' + SMT_UTILS.escapeHtml(lt.LeadingTeacherID) + '">' + SMT_UTILS.escapeHtml(lt.LeadingTeacherName) + ' (' + SMT_UTILS.escapeHtml(lt.LeadingTeacherID) + ')</option>';
      })
    ).join('');

    var tbody = document.getElementById('admin-lt-tbody');
    if (state.leadingTeachers.length === 0) {
      tbody.innerHTML = '<tr><td colspan="5" class="empty-state">No Leading Teachers configured yet.</td></tr>';
      return;
    }

    tbody.innerHTML = state.leadingTeachers.map(function (lt) {
      var isAct = lt.Status.toLowerCase() === 'active';
      var activeAsgs = (lt.assignments || []).filter(function (a) { return a.Status.toLowerCase() === 'active'; });
      var badges = activeAsgs.map(function (a) {
        return '<span class="badge badge-normal" style="margin-right:0.25rem;">' + SMT_UTILS.escapeHtml(a.ClassName) + ' (' + SMT_UTILS.escapeHtml(a.Period) + ')</span>';
      }).join('') || '<span style="color:var(--text-muted)">No active classes</span>';

      return '<tr>' +
        '<td data-label="LT ID"><strong>' + SMT_UTILS.escapeHtml(lt.LeadingTeacherID) + '</strong></td>' +
        '<td data-label="Name">' + SMT_UTILS.escapeHtml(lt.LeadingTeacherName) + '</td>' +
        '<td data-label="Assigned Classes"><div>' + badges + '</div></td>' +
        '<td data-label="Status"><span class="badge ' + (isAct ? 'badge-active' : 'badge-inactive') + '">' + SMT_UTILS.escapeHtml(lt.Status) + '</span></td>' +
        '<td data-label="Actions"><div>' +
        '  <button type="button" class="btn btn-secondary btn-sm btn-edit-lt" data-id="' + SMT_UTILS.escapeHtml(lt.LeadingTeacherID) + '">Edit</button> ' +
        '  <button type="button" class="btn ' + (isAct ? 'btn-danger' : 'btn-success') + ' btn-sm btn-toggle-lt" data-id="' + SMT_UTILS.escapeHtml(lt.LeadingTeacherID) + '" data-next="' + (isAct ? 'Inactive' : 'Active') + '">' +
        (isAct ? 'Deactivate' : 'Activate') + '</button>' +
        '</div></td>' +
        '</tr>';
    }).join('');
  }

  // 6. TEACHING ASSIGNMENTS RENDERER
  function renderTeachingAssignmentsTab() {
    var tbody = document.getElementById('admin-tasg-tbody');
    if (state.teachingAssignments.length === 0) {
      tbody.innerHTML = '<tr><td colspan="6" class="empty-state">No teaching assignments recorded yet.</td></tr>';
      return;
    }
    tbody.innerHTML = state.teachingAssignments.map(function (a) {
      var isAct = a.Status.toLowerCase() === 'active';
      return '<tr>' +
        '<td data-label="Assignment ID"><small>' + SMT_UTILS.escapeHtml(a.AssignmentID) + '</small></td>' +
        '<td data-label="Teacher"><strong>' + SMT_UTILS.escapeHtml(a.TeacherName) + '</strong> (' + SMT_UTILS.escapeHtml(a.TeacherID) + ')</td>' +
        '<td data-label="Class">' + SMT_UTILS.escapeHtml(a.ClassName) + ' (' + SMT_UTILS.escapeHtml(a.ClassID) + ')</td>' +
        '<td data-label="Status"><span class="badge ' + (isAct ? 'badge-active' : 'badge-inactive') + '">' + SMT_UTILS.escapeHtml(a.Status) + '</span></td>' +
        '<td data-label="Updated"><small>' + SMT_UTILS.escapeHtml(SMT_UTILS.formatDateTime(a.UpdatedAt || a.CreatedAt)) + '</small></td>' +
        '<td data-label="Actions">' +
        '  <button type="button" class="btn ' + (isAct ? 'btn-danger' : 'btn-success') + ' btn-sm btn-toggle-tasg" data-id="' + SMT_UTILS.escapeHtml(a.AssignmentID) + '" data-next="' + (isAct ? 'Inactive' : 'Active') + '">' +
        (isAct ? 'Deactivate' : 'Activate') + '</button>' +
        '</td>' +
        '</tr>';
    }).join('');
  }

  // 7. LEADING TEACHER CLASS ASSIGNMENTS RENDERER
  function renderLtAssignmentsTab() {
    var tbody = document.getElementById('admin-ltasg-tbody');
    if (state.ltAssignments.length === 0) {
      tbody.innerHTML = '<tr><td colspan="6" class="empty-state">No Leading Teacher class assignments recorded yet.</td></tr>';
      return;
    }
    tbody.innerHTML = state.ltAssignments.map(function (a) {
      var isAct = a.Status.toLowerCase() === 'active';
      return '<tr>' +
        '<td data-label="Assignment ID"><small>' + SMT_UTILS.escapeHtml(a.AssignmentID) + '</small></td>' +
        '<td data-label="Leading Teacher"><strong>' + SMT_UTILS.escapeHtml(a.LeadingTeacherName) + '</strong> (' + SMT_UTILS.escapeHtml(a.LeadingTeacherID) + ')</td>' +
        '<td data-label="Class">' + SMT_UTILS.escapeHtml(a.ClassName) + ' (' + SMT_UTILS.escapeHtml(a.ClassID) + ')</td>' +
        '<td data-label="Period">' + SMT_UTILS.escapeHtml(a.Period) + '</td>' +
        '<td data-label="Status"><span class="badge ' + (isAct ? 'badge-active' : 'badge-inactive') + '">' + SMT_UTILS.escapeHtml(a.Status) + '</span></td>' +
        '<td data-label="Actions">' +
        '  <button type="button" class="btn ' + (isAct ? 'btn-danger' : 'btn-success') + ' btn-sm btn-toggle-ltasg" data-id="' + SMT_UTILS.escapeHtml(a.AssignmentID) + '" data-next="' + (isAct ? 'Inactive' : 'Active') + '">' +
        (isAct ? 'Deactivate' : 'Activate') + '</button>' +
        '</td>' +
        '</tr>';
    }).join('');
  }

  // 8. SETTINGS & DESTINATIONS RENDERER
  function renderSettingsTab() {
    var s = state.settings;
    document.getElementById('set-school-name').value = s.schoolName || '';
    document.getElementById('set-school-tz').value = s.schoolTimezone || 'Indian/Maldives';
    document.getElementById('set-warn-mins').value = s.warningMinutes || 10;
    document.getElementById('set-crit-mins').value = s.criticalMinutes || 15;
    document.getElementById('set-session-timeout').value = s.sessionTimeoutMinutes || 480;
    document.getElementById('set-teacher-timeout').value = s.teacherSessionTimeoutMinutes || 240;

    var destTbody = document.getElementById('admin-destinations-tbody');
    var dests = s.destinations || [];
    if (dests.length === 0) {
      destTbody.innerHTML = '<tr><td colspan="2" class="empty-state">No destinations configured.</td></tr>';
      return;
    }
    destTbody.innerHTML = dests.map(function (d) {
      return '<tr>' +
        '<td><strong>' + SMT_UTILS.escapeHtml(d) + '</strong></td>' +
        '<td><button type="button" class="btn btn-danger btn-sm btn-remove-dest" data-dest="' + SMT_UTILS.escapeHtml(d) + '">Remove</button></td>' +
        '</tr>';
    }).join('');
  }

  // 9. SECURITY & PASSCODES RENDERER
  function renderSecurityTab() {
    var box = document.getElementById('security-metadata-box');
    var meta = state.securityMeta || {};
    var cfg = meta.credentialsConfigured || {};
    var upd = meta.lastUpdated || {};
    box.innerHTML = [
      '<strong>Credential Provisioning Status (Salted SHA-256 Hashes Only — Plaintext Never Stored):</strong>',
      '<ul style="margin: 0.4rem 0 0; padding-left: 1.2rem;">',
      '  <li>Teacher App PIN: <strong>' + (cfg.teacher_pin ? 'Configured' : 'Not Set') + '</strong>' + (upd.teacher_pin ? ' (Updated: ' + SMT_UTILS.escapeHtml(SMT_UTILS.formatDateTime(upd.teacher_pin)) + ')' : '') + '</li>',
      '  <li>Leading Teacher Passcode: <strong>' + (cfg.leading_teacher ? 'Configured' : 'Not Set') + '</strong>' + (upd.leading_teacher ? ' (Updated: ' + SMT_UTILS.escapeHtml(SMT_UTILS.formatDateTime(upd.leading_teacher)) + ')' : '') + '</li>',
      '  <li>Principal Passcode: <strong>' + (cfg.principal ? 'Configured' : 'Not Set') + '</strong>' + (upd.principal ? ' (Updated: ' + SMT_UTILS.escapeHtml(SMT_UTILS.formatDateTime(upd.principal)) + ')' : '') + '</li>',
      '  <li>Admin Passcode: <strong>' + (cfg.admin ? 'Configured' : 'Not Set') + '</strong>' + (upd.admin ? ' (Updated: ' + SMT_UTILS.escapeHtml(SMT_UTILS.formatDateTime(upd.admin)) + ')' : '') + '</li>',
      '</ul>'
    ].join('');
  }

  // 10. REPORTS & ANALYTICS
  async function runMovementReport() {
    var session = SMT_API.getSession('admin');
    if (!session || !session.token) return;
    try {
      var res = await SMT_API.get('getMovementReports', {
        dateRange: document.getElementById('rep-date-range').value,
        startDate: document.getElementById('rep-start-date').value,
        endDate: document.getElementById('rep-end-date').value,
        classId: document.getElementById('rep-class').value,
        destination: document.getElementById('rep-destination').value,
        status: document.getElementById('rep-status').value,
        search: document.getElementById('rep-search').value
      }, { role: 'admin', silentLoading: true });
      state.reportData = res.data || {};
      renderReportTab();
    } catch (err) {
      SMT_UTILS.showToast(err.message || 'Failed generating report', 'error');
    }
  }

  function renderReportTab() {
    var rep = state.reportData || {};
    var sum = rep.summary || {};
    document.getElementById('rep-stat-total').textContent = String(sum.totalMovements || 0);
    document.getElementById('rep-stat-out').textContent = String(sum.currentOutCount || 0);
    document.getElementById('rep-stat-returned').textContent = String(sum.returnedCount || 0);
    document.getElementById('rep-stat-avg').textContent = (sum.averageDurationMinutes || 0) + 'm';
    document.getElementById('rep-stat-max').textContent = (sum.maxDurationMinutes || 0) + 'm';
    document.getElementById('rep-stat-critical').textContent = String(sum.criticalMovementsCount || 0);

    var byDest = rep.byDestination || [];
    document.getElementById('rep-by-destination-tbody').innerHTML = byDest.length
      ? byDest.map(function (d) {
          return '<tr><td>' + SMT_UTILS.escapeHtml(d.destination) + '</td><td><strong>' + d.count + '</strong></td></tr>';
        }).join('')
      : '<tr><td colspan="2" class="empty-state">No data</td></tr>';

    var byCls = rep.byClass || [];
    document.getElementById('rep-by-class-tbody').innerHTML = byCls.length
      ? byCls.map(function (c) {
          return '<tr><td>' + SMT_UTILS.escapeHtml(c.className) + '</td><td><strong>' + c.count + '</strong></td></tr>';
        }).join('')
      : '<tr><td colspan="2" class="empty-state">No data</td></tr>';

    var repeated = rep.repeatedStudents || [];
    document.getElementById('rep-repeated-students-tbody').innerHTML = repeated.length
      ? repeated.map(function (s) {
          return '<tr>' +
            '<td><strong>' + SMT_UTILS.escapeHtml(s.StudentName) + '</strong> (' + SMT_UTILS.escapeHtml(s.StudentID) + ')</td>' +
            '<td>' + SMT_UTILS.escapeHtml(s.ClassName) + '</td>' +
            '<td><strong>' + s.movementCount + '</strong></td>' +
            '<td>' + s.totalDurationMinutes + 'm</td>' +
            '<td>' + s.criticalCount + '</td>' +
            '</tr>';
        }).join('')
      : '<tr><td colspan="5" class="empty-state">No repeated student movements in selected period.</td></tr>';

    var records = rep.records || [];
    document.getElementById('rep-records-tbody').innerHTML = records.length
      ? records.map(function (r) {
          var rowCls = r.attentionLevel === 'Critical' ? 'row-critical' : (r.attentionLevel === 'Warning' ? 'row-warning' : '');
          return '<tr class="' + rowCls + '">' +
            '<td data-label="Movement ID"><small>' + SMT_UTILS.escapeHtml(r.MovementID) + '</small></td>' +
            '<td data-label="Student"><strong>' + SMT_UTILS.escapeHtml(r.StudentName) + '</strong> (' + SMT_UTILS.escapeHtml(r.StudentID) + ')</td>' +
            '<td data-label="Class">' + SMT_UTILS.escapeHtml(r.ClassName) + '</td>' +
            '<td data-label="Destination">' + SMT_UTILS.escapeHtml(r.Destination) + '</td>' +
            '<td data-label="Teacher">' + SMT_UTILS.escapeHtml(r.TeacherName) + '</td>' +
            '<td data-label="Time Out">' + SMT_UTILS.escapeHtml(SMT_UTILS.formatDateTime(r.TimeOut, r.MovementID)) + '</td>' +
            '<td data-label="Time In">' + (r.TimeIn ? SMT_UTILS.escapeHtml(SMT_UTILS.formatDateTime(r.TimeIn)) : '—') + '</td>' +
            '<td data-label="Duration"><strong>' + SMT_UTILS.escapeHtml(r.formattedDuration) + '</strong></td>' +
            '<td data-label="Attention">' + SMT_UTILS.attentionBadgeHtml(r.attentionLevel) + '</td>' +
            '<td data-label="Status"><span class="badge ' + (r.Status === 'OUT' ? 'badge-warning' : 'badge-active') + '">' + SMT_UTILS.escapeHtml(r.Status) + '</span></td>' +
            '</tr>';
        }).join('')
      : '<tr><td colspan="10" class="empty-state">No movement records match the selected report filters.</td></tr>';
  }

  function bindEvents() {
    SMT_API.onConnectionStateChange(updateConnectionIndicator);

    els.endpointForm.addEventListener('submit', async function (e) {
      e.preventDefault();
      var url = String(els.inputWebAppUrl.value || '').trim();
      SMT_CONFIG.setWebAppUrl(url);
      if (SMT_API.getSession('admin')) {
        await loadAllAdminData();
      } else {
        await verifyBackendConnection(true);
      }
    });

    els.loginForm.addEventListener('submit', async function (e) {
      e.preventDefault();
      // If user typed a URL in the endpoint box without clicking Save & Connect first, save it automatically
      var typedUrl = els.inputWebAppUrl ? String(els.inputWebAppUrl.value || '').trim() : '';
      if (typedUrl && typedUrl !== SMT_CONFIG.getWebAppUrl()) {
        SMT_CONFIG.setWebAppUrl(typedUrl);
      }
      var pass = els.inputPasscode.value;
      if (!pass) return;
      els.btnLogin.disabled = true;
      els.authFeedback.textContent = '';
      try {
        var res = await SMT_API.post('verifyPasscode', {
          role: 'admin',
          passcode: pass
        });
        SMT_API.setSession('admin', res.data);
        els.inputPasscode.value = '';
        SMT_UTILS.showToast('Signed in to Admin Dashboard', 'success');
        await loadAllAdminData();
      } catch (err) {
        els.authFeedback.textContent = err.message || 'Invalid Administrator passcode.';
      } finally {
        els.btnLogin.disabled = false;
      }
    });

    els.btnLogout.addEventListener('click', function () {
      lockAdminSession();
      SMT_UTILS.showToast('Admin session locked.', 'info');
    });

    // Horizontal Scrollable Tabs Navigation with auto-reveal
    var tabBtns = document.querySelectorAll('[data-admin-tab]');
    for (var i = 0; i < tabBtns.length; i++) {
      tabBtns[i].addEventListener('click', function (e) {
        var btn = e.currentTarget;
        var tabName = btn.getAttribute('data-admin-tab');
        for (var j = 0; j < tabBtns.length; j++) {
          tabBtns[j].classList.remove('active');
        }
        btn.classList.add('active');
        if (btn.scrollIntoView) {
          btn.scrollIntoView({ behavior: 'smooth', inline: 'center', block: 'nearest' });
        }
        var panels = document.querySelectorAll('.tab-panel');
        for (var k = 0; k < panels.length; k++) {
          panels[k].classList.remove('active');
        }
        var targetPanel = document.getElementById('tab-' + tabName);
        if (targetPanel) targetPanel.classList.add('active');
      });
    }

    document.getElementById('btn-refresh-overview').addEventListener('click', loadAllAdminData);

    // STUDENT CRUD & FILTERS
    document.getElementById('stu-filter-class').addEventListener('change', renderStudentsTab);
    document.getElementById('stu-filter-status').addEventListener('change', renderStudentsTab);
    document.getElementById('stu-filter-search').addEventListener('input', renderStudentsTab);

    var resetStudentForm = function () {
      document.getElementById('student-edit-mode').value = 'add';
      document.getElementById('student-form-heading').textContent = 'Add New Student';
      var idInput = document.getElementById('stu-input-id');
      idInput.value = '';
      idInput.readOnly = false;
      document.getElementById('stu-input-name').value = '';
      document.getElementById('stu-select-class').value = '';
      document.getElementById('stu-select-status').value = 'Active';
    };
    document.getElementById('btn-reset-student-form').addEventListener('click', resetStudentForm);

    document.getElementById('student-form').addEventListener('submit', async function (e) {
      e.preventDefault();
      var mode = document.getElementById('student-edit-mode').value;
      var studentId = document.getElementById('stu-input-id').value.trim();
      var studentName = document.getElementById('stu-input-name').value.trim();
      var classId = document.getElementById('stu-select-class').value;
      var status = document.getElementById('stu-select-status').value;

      try {
        var action = mode === 'edit' ? 'updateStudent' : 'addStudent';
        var res = await SMT_API.post(action, {
          studentId: studentId,
          studentName: studentName,
          classId: classId,
          status: status
        }, { role: 'admin' });
        SMT_UTILS.showToast(res.message || 'Student saved', 'success');
        resetStudentForm();
        await loadAllAdminData();
      } catch (err) {
        SMT_UTILS.showToast(err.message || 'Failed saving student', 'error');
      }
    });

    document.getElementById('admin-students-tbody').addEventListener('click', async function (e) {
      var editBtn = e.target.closest('.btn-edit-student');
      if (editBtn) {
        var sid = editBtn.getAttribute('data-id');
        var stu = state.students.find(function (s) { return s.StudentID === sid; });
        if (!stu) return;
        document.getElementById('student-edit-mode').value = 'edit';
        document.getElementById('student-form-heading').textContent = 'Edit Student (' + stu.StudentID + ')';
        var idInput = document.getElementById('stu-input-id');
        idInput.value = stu.StudentID;
        idInput.readOnly = true;
        document.getElementById('stu-input-name').value = stu.StudentName;
        document.getElementById('stu-select-class').value = stu.ClassID;
        document.getElementById('stu-select-status').value = stu.Status;
        idInput.scrollIntoView({ behavior: 'smooth', block: 'center' });
        return;
      }

      var toggleBtn = e.target.closest('.btn-toggle-student');
      if (toggleBtn) {
        var targetId = toggleBtn.getAttribute('data-id');
        var nextStatus = toggleBtn.getAttribute('data-next');
        var confirmed = await SMT_UTILS.confirmDialog({
          title: nextStatus === 'Inactive' ? 'Deactivate Student' : 'Activate Student',
          message: 'Are you sure you want to mark student "' + targetId + '" as ' + nextStatus + '?',
          confirmText: nextStatus === 'Inactive' ? 'Deactivate' : 'Activate',
          danger: nextStatus === 'Inactive'
        });
        if (!confirmed) return;
        try {
          await SMT_API.post('deactivateStudent', { studentId: targetId, status: nextStatus }, { role: 'admin' });
          SMT_UTILS.showToast('Student status updated to ' + nextStatus, 'success');
          await loadAllAdminData();
        } catch (err) {
          SMT_UTILS.showToast(err.message || 'Could not update student status', 'error');
        }
      }
    });

    // BULK STUDENT CSV / EXCEL IMPORT
    document.getElementById('btn-download-student-template').addEventListener('click', function () {
      var sample = 'StudentID,StudentName,Class\r\nSTU-2001,Mohamed Rasheed,Grade 10A\r\nSTU-2002,Mariyam Niuma,Grade 10B\r\n';
      SMT_UTILS.downloadCsvFile('sample-students-import.csv', sample);
    });

    document.getElementById('bulk-csv-file').addEventListener('change', function (e) {
      var file = e.target.files && e.target.files[0];
      if (!file) return;
      SMT_UTILS.readSpreadsheetFileAsText(file).then(function (text) {
        document.getElementById('bulk-paste-textarea').value = text;
        SMT_UTILS.showToast('Loaded file "' + file.name + '". Click Validate & Preview Import.', 'info');
      }).catch(function (err) {
        SMT_UTILS.showToast(err.message || 'Could not read spreadsheet file', 'error');
      });
    });

    document.getElementById('btn-preview-bulk').addEventListener('click', function () {
      var raw = document.getElementById('bulk-paste-textarea').value;
      var preview = SMT_UTILS.parseAndValidateStudentImport(raw, state.classes, state.students);
      state.bulkPreview = preview;

      document.getElementById('bulk-preview-container').style.display = 'block';
      document.getElementById('bulk-valid-count').textContent = String(preview.valid.length);
      document.getElementById('bulk-skipped-count').textContent = String(preview.skipped.length);
      document.getElementById('bulk-invalid-count').textContent = String(preview.invalid.length);
      document.getElementById('bulk-valid-label').textContent = String(preview.valid.length);

      var issues = preview.skipped.concat(preview.invalid);
      document.getElementById('bulk-issues-label').textContent = String(issues.length);

      document.getElementById('bulk-valid-tbody').innerHTML = preview.valid.length
        ? preview.valid.map(function (v) {
            return '<tr><td>' + v.rowNumber + '</td><td><strong>' + SMT_UTILS.escapeHtml(v.studentId) + '</strong></td><td>' + SMT_UTILS.escapeHtml(v.studentName) + '</td><td>' + SMT_UTILS.escapeHtml(v.className) + '</td></tr>';
          }).join('')
        : '<tr><td colspan="4" class="empty-state">No valid rows to import.</td></tr>';

      document.getElementById('bulk-issues-tbody').innerHTML = issues.length
        ? issues.map(function (iss) {
            return '<tr><td>' + iss.rowNumber + '</td><td>' + SMT_UTILS.escapeHtml(iss.studentId || '—') + '</td><td>' + SMT_UTILS.escapeHtml(iss.reason) + '</td></tr>';
          }).join('')
        : '<tr><td colspan="3" class="empty-state">No invalid or skipped rows.</td></tr>';

      var confirmBtn = document.getElementById('btn-confirm-bulk-import');
      confirmBtn.disabled = preview.valid.length === 0;
      confirmBtn.textContent = 'Confirm & Import Valid Rows (' + preview.valid.length + ')';
    });

    document.getElementById('btn-confirm-bulk-import').addEventListener('click', async function () {
      if (!state.bulkPreview || state.bulkPreview.valid.length === 0) return;
      var ok = await SMT_UTILS.confirmDialog({
        title: 'Confirm Bulk Student Import',
        message: 'Import ' + state.bulkPreview.valid.length + ' validated student(s) into Google Sheets? Existing students and Movement_Log will not be modified.',
        confirmText: 'Import ' + state.bulkPreview.valid.length + ' Students'
      });
      if (!ok) return;

      try {
        var res = await SMT_API.post('bulkAddStudents', {
          students: state.bulkPreview.valid
        }, { role: 'admin' });
        var d = res.data || {};
        document.getElementById('bulk-import-result-summary').textContent =
          'Import Summary: ' + (d.addedCount || 0) + ' added, ' +
          ((d.skippedCount || 0) + state.bulkPreview.skipped.length) + ' skipped (existing), ' +
          ((d.invalidCount || 0) + state.bulkPreview.invalid.length) + ' invalid.';
        SMT_UTILS.showToast(res.message || 'Bulk import completed', 'success');
        document.getElementById('btn-confirm-bulk-import').disabled = true;
        await loadAllAdminData();
      } catch (err) {
        SMT_UTILS.showToast(err.message || 'Bulk import failed', 'error');
      }
    });

    document.getElementById('btn-clear-bulk').addEventListener('click', function () {
      state.bulkPreview = null;
      document.getElementById('bulk-csv-file').value = '';
      document.getElementById('bulk-paste-textarea').value = '';
      document.getElementById('bulk-preview-container').style.display = 'none';
      document.getElementById('bulk-import-result-summary').textContent = '';
      var confirmBtn = document.getElementById('btn-confirm-bulk-import');
      confirmBtn.disabled = true;
      confirmBtn.textContent = 'Confirm & Import Valid Rows (0)';
    });

    // CLASS CRUD
    document.getElementById('cls-filter-search').addEventListener('input', renderClassesTab);
    var resetClassForm = function () {
      document.getElementById('class-edit-mode').value = 'add';
      document.getElementById('class-form-heading').textContent = 'Add / Edit Class';
      var idEl = document.getElementById('cls-input-id');
      idEl.value = '';
      idEl.readOnly = false;
      document.getElementById('cls-input-name').value = '';
      document.getElementById('cls-select-status').value = 'Active';
    };
    document.getElementById('btn-reset-class-form').addEventListener('click', resetClassForm);

    document.getElementById('class-form').addEventListener('submit', async function (e) {
      e.preventDefault();
      var mode = document.getElementById('class-edit-mode').value;
      var classId = document.getElementById('cls-input-id').value.trim();
      var className = document.getElementById('cls-input-name').value.trim();
      var status = document.getElementById('cls-select-status').value;

      try {
        var action = mode === 'edit' ? 'updateClass' : 'addClass';
        var res = await SMT_API.post(action, {
          classId: classId,
          className: className,
          status: status,
          confirmEnrolledWarning: true
        }, { role: 'admin' });
        SMT_UTILS.showToast(res.message || 'Class saved', 'success');
        resetClassForm();
        await loadAllAdminData();
      } catch (err) {
        SMT_UTILS.showToast(err.message || 'Failed saving class', 'error');
      }
    });

    document.getElementById('admin-classes-tbody').addEventListener('click', async function (e) {
      var editBtn = e.target.closest('.btn-edit-class');
      if (editBtn) {
        var cid = editBtn.getAttribute('data-id');
        var cls = state.classes.find(function (c) { return c.ClassID === cid; });
        if (!cls) return;
        document.getElementById('class-edit-mode').value = 'edit';
        document.getElementById('class-form-heading').textContent = 'Edit / Rename Class (' + cls.ClassID + ')';
        var idEl = document.getElementById('cls-input-id');
        idEl.value = cls.ClassID;
        idEl.readOnly = true;
        document.getElementById('cls-input-name').value = cls.ClassName;
        document.getElementById('cls-select-status').value = cls.Status;
        idEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
        return;
      }

      var toggleBtn = e.target.closest('.btn-toggle-class');
      if (toggleBtn) {
        var targetId = toggleBtn.getAttribute('data-id');
        var targetName = toggleBtn.getAttribute('data-name');
        var enrolled = Number(toggleBtn.getAttribute('data-enrolled') || 0);
        var nextStatus = toggleBtn.getAttribute('data-next');
        var warnMsg = nextStatus === 'Inactive' && enrolled > 0
          ? 'Warning: Class "' + targetName + '" currently has ' + enrolled + ' active enrolled student(s). Are you sure you want to deactivate this class?'
          : 'Change status of class "' + targetName + '" to ' + nextStatus + '?';
        var ok = await SMT_UTILS.confirmDialog({
          title: nextStatus === 'Inactive' ? 'Deactivate Class' : 'Activate Class',
          message: warnMsg,
          confirmText: nextStatus === 'Inactive' ? 'Deactivate Class' : 'Activate Class',
          danger: nextStatus === 'Inactive'
        });
        if (!ok) return;
        try {
          await SMT_API.post('deactivateClass', {
            classId: targetId,
            status: nextStatus,
            confirmEnrolledWarning: true
          }, { role: 'admin' });
          SMT_UTILS.showToast('Class status updated to ' + nextStatus, 'success');
          await loadAllAdminData();
        } catch (err) {
          SMT_UTILS.showToast(err.message || 'Could not update class status', 'error');
        }
      }
    });

    // TEACHER CRUD & BULK TEACHER IMPORT (CSV / EXCEL)
    var tchSearchEl = document.getElementById('tch-filter-search');
    if (tchSearchEl) {
      tchSearchEl.addEventListener('input', renderTeachersTab);
    }

    var resetTeacherForm = function () {
      document.getElementById('teacher-edit-mode').value = 'add';
      document.getElementById('teacher-form-heading').textContent = 'Add / Edit Teacher Profile';
      var idEl = document.getElementById('tch-input-id');
      idEl.value = '';
      idEl.readOnly = false;
      document.getElementById('tch-input-name').value = '';
      document.getElementById('tch-select-status').value = 'Active';
    };
    document.getElementById('btn-reset-teacher-form').addEventListener('click', resetTeacherForm);

    document.getElementById('teacher-form').addEventListener('submit', async function (e) {
      e.preventDefault();
      var mode = document.getElementById('teacher-edit-mode').value;
      var teacherId = document.getElementById('tch-input-id').value.trim();
      var teacherName = document.getElementById('tch-input-name').value.trim();
      var status = document.getElementById('tch-select-status').value;
      try {
        var action = mode === 'edit' ? 'updateTeacher' : 'addTeacher';
        var res = await SMT_API.post(action, {
          teacherId: teacherId,
          teacherName: teacherName,
          status: status
        }, { role: 'admin' });
        SMT_UTILS.showToast(res.message || 'Teacher saved', 'success');
        resetTeacherForm();
        await loadAllAdminData();
      } catch (err) {
        SMT_UTILS.showToast(err.message || 'Failed saving teacher', 'error');
      }
    });

    document.getElementById('btn-download-teacher-template').addEventListener('click', function () {
      var sample = [
        'TeacherID,TeacherName,Classes,Status',
        'TCH-101,Ahmed Naseer,Grade 10A; Grade 10B,Active',
        'TCH-102,Aminath Rasheeda,Grade 10A,Active',
        'TCH-103,Ibrahim Shareef,,Active'
      ].join('\r\n');
      SMT_UTILS.downloadCsvFile('sample-teachers-import.csv', sample);
    });

    document.getElementById('bulk-teacher-file').addEventListener('change', function (e) {
      var file = e.target.files && e.target.files[0];
      if (!file) return;
      SMT_UTILS.readSpreadsheetFileAsText(file).then(function (text) {
        document.getElementById('bulk-teacher-textarea').value = text;
        SMT_UTILS.showToast('Loaded "' + file.name + '". Click Validate & Preview Teachers.', 'info');
      }).catch(function (err) {
        SMT_UTILS.showToast(err.message || 'Could not read teacher spreadsheet file', 'error');
      });
    });

    document.getElementById('btn-preview-bulk-teachers').addEventListener('click', function () {
      var raw = document.getElementById('bulk-teacher-textarea').value;
      var preview = SMT_UTILS.parseAndValidateTeacherImport(raw, state.teachers, state.classes);
      state.bulkTeacherPreview = preview;

      document.getElementById('bulk-teacher-preview-container').style.display = 'block';
      document.getElementById('bulk-tch-valid-count').textContent = String(preview.valid.length);
      document.getElementById('bulk-tch-skipped-count').textContent = String(preview.skipped.length);
      document.getElementById('bulk-tch-invalid-count').textContent = String(preview.invalid.length);
      document.getElementById('bulk-tch-valid-label').textContent = String(preview.valid.length);

      var issues = preview.skipped.concat(preview.invalid);
      document.getElementById('bulk-tch-issues-label').textContent = String(issues.length);

      document.getElementById('bulk-tch-valid-tbody').innerHTML = preview.valid.length
        ? preview.valid.map(function (v) {
            var clsText = (v.classNames && v.classNames.length > 0) ? SMT_UTILS.escapeHtml(v.classNames.join(', ')) : '<span style="color:var(--text-muted)">—</span>';
            return '<tr><td>' + v.rowNumber + '</td><td><strong>' + SMT_UTILS.escapeHtml(v.teacherId) + '</strong></td><td>' + SMT_UTILS.escapeHtml(v.teacherName) + '</td><td>' + clsText + '</td></tr>';
          }).join('')
        : '<tr><td colspan="4" class="empty-state">No valid teacher rows to import.</td></tr>';

      document.getElementById('bulk-tch-issues-tbody').innerHTML = issues.length
        ? issues.map(function (iss) {
            return '<tr><td>' + iss.rowNumber + '</td><td>' + SMT_UTILS.escapeHtml(iss.teacherId || '—') + '</td><td>' + SMT_UTILS.escapeHtml(iss.reason) + '</td></tr>';
          }).join('')
        : '<tr><td colspan="3" class="empty-state">No invalid or skipped rows.</td></tr>';

      var confirmBtn = document.getElementById('btn-confirm-bulk-teachers');
      confirmBtn.disabled = preview.valid.length === 0;
      confirmBtn.textContent = 'Confirm & Import Valid Teachers (' + preview.valid.length + ')';
    });

    document.getElementById('btn-confirm-bulk-teachers').addEventListener('click', async function () {
      if (!state.bulkTeacherPreview || state.bulkTeacherPreview.valid.length === 0) return;
      var ok = await SMT_UTILS.confirmDialog({
        title: 'Confirm Bulk Teacher Import',
        message: 'Import ' + state.bulkTeacherPreview.valid.length + ' validated teacher(s) into Google Sheets? Existing teachers and Movement_Log will not be modified.',
        confirmText: 'Import ' + state.bulkTeacherPreview.valid.length + ' Teachers'
      });
      if (!ok) return;

      try {
        var res = await SMT_API.post('bulkAddTeachers', {
          teachers: state.bulkTeacherPreview.valid
        }, { role: 'admin' });
        var d = res.data || {};
        document.getElementById('bulk-teacher-result-summary').textContent =
          'Teacher Import Summary: ' + (d.addedCount || 0) + ' teachers added, ' +
          (d.assignmentsCreatedCount || 0) + ' class assignments created, ' +
          ((d.skippedCount || 0) + state.bulkTeacherPreview.skipped.length) + ' skipped (existing), ' +
          ((d.invalidCount || 0) + state.bulkTeacherPreview.invalid.length) + ' invalid.';
        SMT_UTILS.showToast(res.message || 'Bulk teacher import completed', 'success');
        document.getElementById('btn-confirm-bulk-teachers').disabled = true;
        await loadAllAdminData();
      } catch (err) {
        SMT_UTILS.showToast(err.message || 'Bulk teacher import failed', 'error');
      }
    });

    document.getElementById('btn-clear-bulk-teachers').addEventListener('click', function () {
      state.bulkTeacherPreview = null;
      document.getElementById('bulk-teacher-file').value = '';
      document.getElementById('bulk-teacher-textarea').value = '';
      document.getElementById('bulk-teacher-preview-container').style.display = 'none';
      document.getElementById('bulk-teacher-result-summary').textContent = '';
      var confirmBtn = document.getElementById('btn-confirm-bulk-teachers');
      confirmBtn.disabled = true;
      confirmBtn.textContent = 'Confirm & Import Valid Teachers (0)';
    });

    document.getElementById('admin-teachers-tbody').addEventListener('click', async function (e) {
      var editBtn = e.target.closest('.btn-edit-teacher');
      if (editBtn) {
        var tid = editBtn.getAttribute('data-id');
        var t = state.teachers.find(function (item) { return item.TeacherID === tid; });
        if (!t) return;
        document.getElementById('teacher-edit-mode').value = 'edit';
        document.getElementById('teacher-form-heading').textContent = 'Edit Teacher (' + t.TeacherID + ')';
        var idEl = document.getElementById('tch-input-id');
        idEl.value = t.TeacherID;
        idEl.readOnly = true;
        document.getElementById('tch-input-name').value = t.TeacherName;
        document.getElementById('tch-select-status').value = t.Status;
        idEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
        return;
      }

      var toggleBtn = e.target.closest('.btn-toggle-teacher');
      if (toggleBtn) {
        var targetId = toggleBtn.getAttribute('data-id');
        var nextStatus = toggleBtn.getAttribute('data-next');
        var ok = await SMT_UTILS.confirmDialog({
          title: nextStatus === 'Inactive' ? 'Deactivate Teacher' : 'Activate Teacher',
          message: 'Change status of teacher "' + targetId + '" to ' + nextStatus + '?',
          confirmText: nextStatus,
          danger: nextStatus === 'Inactive'
        });
        if (!ok) return;
        try {
          await SMT_API.post('deactivateTeacher', { teacherId: targetId, status: nextStatus }, { role: 'admin' });
          SMT_UTILS.showToast('Teacher updated to ' + nextStatus, 'success');
          await loadAllAdminData();
        } catch (err) {
          SMT_UTILS.showToast(err.message || 'Failed updating teacher', 'error');
        }
      }
    });

    // LEADING TEACHER CRUD
    var resetLtForm = function () {
      document.getElementById('lt-edit-mode').value = 'add';
      document.getElementById('lt-form-heading').textContent = 'Add / Edit Leading Teacher Profile';
      var idEl = document.getElementById('lt-input-id');
      idEl.value = '';
      idEl.readOnly = false;
      document.getElementById('lt-input-name').value = '';
      document.getElementById('lt-select-status').value = 'Active';
    };
    document.getElementById('btn-reset-lt-form').addEventListener('click', resetLtForm);

    document.getElementById('lt-form').addEventListener('submit', async function (e) {
      e.preventDefault();
      var mode = document.getElementById('lt-edit-mode').value;
      var ltId = document.getElementById('lt-input-id').value.trim();
      var ltName = document.getElementById('lt-input-name').value.trim();
      var status = document.getElementById('lt-select-status').value;
      try {
        var action = mode === 'edit' ? 'updateLeadingTeacher' : 'addLeadingTeacher';
        var res = await SMT_API.post(action, {
          leadingTeacherId: ltId,
          leadingTeacherName: ltName,
          status: status
        }, { role: 'admin' });
        SMT_UTILS.showToast(res.message || 'Leading Teacher saved', 'success');
        resetLtForm();
        await loadAllAdminData();
      } catch (err) {
        SMT_UTILS.showToast(err.message || 'Failed saving Leading Teacher', 'error');
      }
    });

    document.getElementById('admin-lt-tbody').addEventListener('click', async function (e) {
      var editBtn = e.target.closest('.btn-edit-lt');
      if (editBtn) {
        var id = editBtn.getAttribute('data-id');
        var lt = state.leadingTeachers.find(function (item) { return item.LeadingTeacherID === id; });
        if (!lt) return;
        document.getElementById('lt-edit-mode').value = 'edit';
        document.getElementById('lt-form-heading').textContent = 'Edit Leading Teacher (' + lt.LeadingTeacherID + ')';
        var idEl = document.getElementById('lt-input-id');
        idEl.value = lt.LeadingTeacherID;
        idEl.readOnly = true;
        document.getElementById('lt-input-name').value = lt.LeadingTeacherName;
        document.getElementById('lt-select-status').value = lt.Status;
        idEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
        return;
      }

      var toggleBtn = e.target.closest('.btn-toggle-lt');
      if (toggleBtn) {
        var targetId = toggleBtn.getAttribute('data-id');
        var nextStatus = toggleBtn.getAttribute('data-next');
        try {
          await SMT_API.post('deactivateLeadingTeacher', { leadingTeacherId: targetId, status: nextStatus }, { role: 'admin' });
          SMT_UTILS.showToast('Leading Teacher status updated', 'success');
          await loadAllAdminData();
        } catch (err) {
          SMT_UTILS.showToast(err.message || 'Failed updating Leading Teacher', 'error');
        }
      }
    });

    // TEACHING ASSIGNMENTS
    document.getElementById('teaching-assignment-form').addEventListener('submit', async function (e) {
      e.preventDefault();
      var teacherId = document.getElementById('tasg-select-teacher').value;
      var classId = document.getElementById('tasg-select-class').value;
      try {
        var res = await SMT_API.post('addTeachingAssignment', { teacherId: teacherId, classId: classId }, { role: 'admin' });
        SMT_UTILS.showToast(res.message || 'Assignment created', 'success');
        await loadAllAdminData();
      } catch (err) {
        SMT_UTILS.showToast(err.message || 'Could not add teaching assignment', 'error');
      }
    });

    document.getElementById('admin-tasg-tbody').addEventListener('click', async function (e) {
      var btn = e.target.closest('.btn-toggle-tasg');
      if (!btn) return;
      var id = btn.getAttribute('data-id');
      var nextStatus = btn.getAttribute('data-next');
      try {
        await SMT_API.post('deactivateTeachingAssignment', { assignmentId: id, status: nextStatus }, { role: 'admin' });
        SMT_UTILS.showToast('Assignment updated to ' + nextStatus, 'success');
        await loadAllAdminData();
      } catch (err) {
        SMT_UTILS.showToast(err.message || 'Failed updating assignment', 'error');
      }
    });

    // LEADING TEACHER CLASS ASSIGNMENTS
    document.getElementById('lt-assignment-form').addEventListener('submit', async function (e) {
      e.preventDefault();
      var ltId = document.getElementById('ltasg-select-lt').value;
      var classId = document.getElementById('ltasg-select-class').value;
      var period = document.getElementById('ltasg-input-period').value.trim();
      try {
        var res = await SMT_API.post('addLeadingTeacherAssignment', {
          leadingTeacherId: ltId,
          classId: classId,
          period: period
        }, { role: 'admin' });
        SMT_UTILS.showToast(res.message || 'Leading Teacher assigned to class', 'success');
        await loadAllAdminData();
      } catch (err) {
        SMT_UTILS.showToast(err.message || 'Failed assigning class', 'error');
      }
    });

    document.getElementById('admin-ltasg-tbody').addEventListener('click', async function (e) {
      var btn = e.target.closest('.btn-toggle-ltasg');
      if (!btn) return;
      var id = btn.getAttribute('data-id');
      var nextStatus = btn.getAttribute('data-next');
      try {
        await SMT_API.post('deactivateLeadingTeacherAssignment', { assignmentId: id, status: nextStatus }, { role: 'admin' });
        SMT_UTILS.showToast('Leading Teacher assignment updated to ' + nextStatus, 'success');
        await loadAllAdminData();
      } catch (err) {
        SMT_UTILS.showToast(err.message || 'Failed updating assignment', 'error');
      }
    });

    // SETTINGS & DESTINATIONS
    document.getElementById('school-settings-form').addEventListener('submit', async function (e) {
      e.preventDefault();
      try {
        await SMT_API.post('updateSettings', {
          schoolName: document.getElementById('set-school-name').value.trim(),
          schoolTimezone: document.getElementById('set-school-tz').value.trim(),
          warningMinutes: Number(document.getElementById('set-warn-mins').value),
          criticalMinutes: Number(document.getElementById('set-crit-mins').value),
          sessionTimeoutMinutes: Number(document.getElementById('set-session-timeout').value),
          teacherSessionTimeoutMinutes: Number(document.getElementById('set-teacher-timeout').value)
        }, { role: 'admin' });
        SMT_UTILS.showToast('School settings updated', 'success');
        await loadAllAdminData();
      } catch (err) {
        SMT_UTILS.showToast(err.message || 'Failed updating settings', 'error');
      }
    });

    document.getElementById('add-destination-form').addEventListener('submit', async function (e) {
      e.preventDefault();
      var input = document.getElementById('input-new-destination');
      var dest = input.value.trim();
      if (!dest) return;
      try {
        await SMT_API.post('manageDestinations', { operation: 'add', destination: dest }, { role: 'admin' });
        input.value = '';
        SMT_UTILS.showToast('Destination added', 'success');
        await loadAllAdminData();
      } catch (err) {
        SMT_UTILS.showToast(err.message || 'Could not add destination', 'error');
      }
    });

    document.getElementById('admin-destinations-tbody').addEventListener('click', async function (e) {
      var btn = e.target.closest('.btn-remove-dest');
      if (!btn) return;
      var dest = btn.getAttribute('data-dest');
      var ok = await SMT_UTILS.confirmDialog({
        title: 'Remove Destination',
        message: 'Remove "' + dest + '" from the active destinations list?',
        confirmText: 'Remove',
        danger: true
      });
      if (!ok) return;
      try {
        await SMT_API.post('manageDestinations', { operation: 'remove', destination: dest }, { role: 'admin' });
        SMT_UTILS.showToast('Destination removed', 'success');
        await loadAllAdminData();
      } catch (err) {
        SMT_UTILS.showToast(err.message || 'Could not remove destination', 'error');
      }
    });

    // PASSCODE & PIN MANAGEMENT
    document.getElementById('change-passcode-form').addEventListener('submit', async function (e) {
      e.preventDefault();
      var targetRole = document.getElementById('sec-target-role').value;
      var newPass = document.getElementById('sec-new-passcode').value;
      var confirmPass = document.getElementById('sec-confirm-passcode').value;
      var curAdminPass = document.getElementById('sec-current-admin-passcode').value;

      if (newPass !== confirmPass) {
        SMT_UTILS.showToast('New credential and confirmation do not match.', 'error');
        return;
      }

      var ok = await SMT_UTILS.confirmDialog({
        title: 'Confirm Credential Rotation',
        message: 'Update credential for "' + targetRole + '" and invalidate all existing active sessions for that role?',
        confirmText: 'Update & Rotate Secret',
        danger: true
      });
      if (!ok) return;

      try {
        var res = await SMT_API.post('changePasscode', {
          targetRole: targetRole,
          newPasscode: newPass,
          currentAdminPasscode: curAdminPass
        }, { role: 'admin' });
        if (res.data && res.data.newAdminSession) {
          SMT_API.setSession('admin', res.data.newAdminSession);
        }
        document.getElementById('sec-new-passcode').value = '';
        document.getElementById('sec-confirm-passcode').value = '';
        document.getElementById('sec-current-admin-passcode').value = '';
        SMT_UTILS.showToast(res.message || 'Credential updated', 'success');
        await loadAllAdminData();
      } catch (err) {
        SMT_UTILS.showToast(err.message || 'Failed updating credential', 'error');
      }
    });

    // REPORTS & EXPORT
    document.getElementById('btn-run-report').addEventListener('click', runMovementReport);
    document.getElementById('rep-date-range').addEventListener('change', runMovementReport);
    document.getElementById('rep-class').addEventListener('change', runMovementReport);
    document.getElementById('rep-destination').addEventListener('change', runMovementReport);
    document.getElementById('rep-status').addEventListener('change', runMovementReport);

    document.getElementById('btn-export-report-csv').addEventListener('click', function () {
      var recs = (state.reportData && state.reportData.records) || [];
      var csv = SMT_UTILS.movementRecordsToCsv(recs);
      SMT_UTILS.downloadCsvFile('student-movement-report.csv', csv);
    });

    document.getElementById('btn-print-report').addEventListener('click', function () {
      window.print();
    });
  }

  async function init() {
    cacheElements();
    bindEvents();
    var currentUrl = SMT_CONFIG.getWebAppUrl();
    els.inputWebAppUrl.value = currentUrl;
    if (currentUrl) {
      if (SMT_API.getSession('admin')) {
        await loadAllAdminData();
      } else {
        await verifyBackendConnection(false);
      }
    }
  }

  if (typeof document !== 'undefined') {
    document.addEventListener('DOMContentLoaded', init);
  }
})();
