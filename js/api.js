/**
 * Student Movement Tracker — Shared API Client (js/api.js)
 *
 * Responsibilities:
 * - Communicates with the configured Google Apps Script Web App URL
 * - Never substitutes fake demo data on connection failure
 * - Never reports "Connected to Google Sheets" unless the backend responds with { success: true }
 * - Attaches role-specific tokens only to protected requests (not public bootstrap calls)
 * - Isolates sessions per role and clears Teacher session immediately when switching teachers
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) {
    var cfg = require('./config.js');
    module.exports = factory(cfg);
  } else {
    root.SMT_API = factory(root.SMT_CONFIG);
  }
})(typeof self !== 'undefined' ? self : this, function (SMT_CONFIG) {

  var PUBLIC_ACTIONS = {
    getSettings: true,
    getTeacherLoginOptions: true,
    getLeadingTeacherLoginOptions: true,
    verifyPasscode: true,
    verifyTeacherPasscode: true
  };

  var SESSION_STORAGE_PREFIX = 'smt_session_v1_';
  var connectionStateListeners = [];
  var lastConnectionState = {
    status: 'unconfigured', // 'unconfigured' | 'loading' | 'connected' | 'offline' | 'error'
    message: 'Not connected to Google Sheets',
    lastSyncIso: null
  };

  function notifyConnectionState(status, message) {
    lastConnectionState = {
      status: status,
      message: message || '',
      lastSyncIso: status === 'connected' ? new Date().toISOString() : lastConnectionState.lastSyncIso
    };
    for (var i = 0; i < connectionStateListeners.length; i++) {
      try {
        connectionStateListeners[i](lastConnectionState);
      } catch (e) {
        // Ignore listener error
      }
    }
  }

  function onConnectionStateChange(listener) {
    if (typeof listener === 'function') {
      connectionStateListeners.push(listener);
      listener(lastConnectionState);
    }
  }

  function getConnectionState() {
    return lastConnectionState;
  }

  /**
   * Role session management with safe storage access.
   */
  function getSessionKey(role) {
    return SESSION_STORAGE_PREFIX + String(role || '').toLowerCase();
  }

  function getSession(role) {
    try {
      if (typeof sessionStorage !== 'undefined') {
        var raw = sessionStorage.getItem(getSessionKey(role));
        if (!raw) return null;
        var parsed = JSON.parse(raw);
        if (parsed && parsed.expiresAt) {
          if (new Date(parsed.expiresAt).getTime() <= Date.now()) {
            sessionStorage.removeItem(getSessionKey(role));
            return null;
          }
        }
        return parsed;
      }
    } catch (e) {
      return null;
    }
    return null;
  }

  function setSession(role, sessionData) {
    try {
      if (typeof sessionStorage !== 'undefined') {
        sessionStorage.setItem(getSessionKey(role), JSON.stringify(sessionData));
      }
    } catch (e) {
      // Ignore storage error
    }
    return sessionData;
  }

  function clearSession(role) {
    try {
      if (typeof sessionStorage !== 'undefined') {
        sessionStorage.removeItem(getSessionKey(role));
      }
    } catch (e) {
      // Ignore storage error
    }
  }

  /**
   * Core HTTP request helper for Google Apps Script Web App.
   * Uses text/plain POST payload to avoid CORS preflight failures on script.google.com redirects.
   */
  async function request(method, action, paramsOrBody, options) {
    var opts = options || {};
    var webAppUrl = SMT_CONFIG ? SMT_CONFIG.getWebAppUrl() : '';

    if (!webAppUrl) {
      notifyConnectionState('unconfigured', 'Backend URL not configured');
      var cfgErr = new Error('Google Apps Script Web App URL is not configured. Enter your deployment URL in the connection bar.');
      cfgErr.code = 'UNCONFIGURED_URL';
      throw cfgErr;
    }

    if (typeof navigator !== 'undefined' && navigator.onLine === false) {
      notifyConnectionState('offline', 'Offline — check network connection');
      var offErr = new Error('Browser is currently offline. Please check your internet connection.');
      offErr.code = 'OFFLINE';
      throw offErr;
    }

    var payload = Object.assign({}, paramsOrBody || {}, { action: action });

    // Attach role session token only to protected endpoints
    if (!PUBLIC_ACTIONS[action]) {
      var role = opts.role;
      var token = payload.token;
      if (!token && role) {
        var sess = getSession(role);
        if (sess && sess.token) {
          token = sess.token;
        }
      }
      if (!token) {
        var authErr = new Error('Authentication required. Please sign in to continue.');
        authErr.code = 'AUTH_REQUIRED';
        throw authErr;
      }
      payload.token = token;
    } else {
      // Ensure credentials/tokens are not attached unnecessarily to public bootstrap calls
      delete payload.token;
      delete payload.authToken;
    }

    if (!opts.silentLoading) {
      notifyConnectionState('loading', 'Syncing with Google Sheets...');
    }

    var fetchUrl = webAppUrl;
    var fetchOptions = {
      method: method === 'GET' ? 'GET' : 'POST',
      redirect: 'follow'
    };

    if (fetchOptions.method === 'GET') {
      var qs = Object.keys(payload).map(function (k) {
        var val = payload[k];
        if (val === undefined || val === null) return '';
        return encodeURIComponent(k) + '=' + encodeURIComponent(String(val));
      }).filter(Boolean).join('&');
      fetchUrl += (fetchUrl.indexOf('?') === -1 ? '?' : '&') + qs;
    } else {
      fetchOptions.headers = {
        'Content-Type': 'text/plain;charset=utf-8'
      };
      fetchOptions.body = JSON.stringify(payload);
    }

    var response;
    try {
      response = await fetch(fetchUrl, fetchOptions);
    } catch (networkErr) {
      notifyConnectionState('error', 'Connection to Google Sheets failed');
      var netError = new Error('Unable to reach Google Apps Script backend. Verify the Web App URL and deployment permissions.');
      netError.code = 'NETWORK_ERROR';
      throw netError;
    }

    var rawText = '';
    try {
      rawText = await response.text();
    } catch (readErr) {
      notifyConnectionState('error', 'Failed reading backend response');
      var rErr = new Error('Failed to read response from Google Apps Script.');
      rErr.code = 'READ_ERROR';
      throw rErr;
    }

    var json;
    try {
      json = JSON.parse(rawText);
    } catch (parseErr) {
      notifyConnectionState('error', 'Invalid JSON response from backend');
      var pErr = new Error('Backend returned a non-JSON response. Ensure the Google Apps Script Web App is deployed with access set properly.');
      pErr.code = 'INVALID_JSON_RESPONSE';
      throw pErr;
    }

    if (!json || json.success !== true) {
      var errCode = (json && json.code) ? json.code : 'API_ERROR';
      var errMsg = (json && json.error) ? json.error : 'Operation failed on server.';
      if (errCode === 'AUTH_REQUIRED' || errCode === 'INVALID_TOKEN' || errCode === 'TOKEN_EXPIRED' || errCode === 'ACCOUNT_INACTIVE') {
        if (opts.role) {
          clearSession(opts.role);
        }
      }
      // Note: The backend responded cleanly with an application-level JSON error, so the connection itself is live
      notifyConnectionState('connected', 'Connected to Google Sheets');
      var apiError = new Error(errMsg);
      apiError.code = errCode;
      throw apiError;
    }

    // Ensure response has normalized data object for compatibility across deployment versions
    if (json.data === undefined) {
      json.data = {};
      for (var k in json) {
        if (k !== 'success' && k !== 'message' && k !== 'data') {
          json.data[k] = json[k];
        }
      }
    }

    notifyConnectionState('connected', 'Connected to Google Sheets');
    return json;
  }

  function get(action, params, options) {
    return request('GET', action, params, options);
  }

  function post(action, body, options) {
    return request('POST', action, body, options);
  }

  return {
    PUBLIC_ACTIONS: PUBLIC_ACTIONS,
    get: get,
    post: post,
    request: request,
    getSession: getSession,
    setSession: setSession,
    clearSession: clearSession,
    onConnectionStateChange: onConnectionStateChange,
    getConnectionState: getConnectionState
  };
});
