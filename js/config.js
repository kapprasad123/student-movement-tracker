/**
 * Student Movement Tracker — Shared Configuration (js/config.js)
 *
 * Central configuration for Google Apps Script Web App endpoints and runtime thresholds.
 * Defaults to the production Google Apps Script deployment URL while supporting
 * intentional local-development overrides.
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = factory();
  } else {
    root.SMT_CONFIG = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  var PRODUCTION_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbygZM_HFzRnom0e3AFk4pSyvHufSLyF0hCpPba9oSr9wkyGXQrT-8SsMTU3MoLI-vtB/exec';
  var STORAGE_URL_KEY = 'smt_apps_script_webapp_url';
  var explicitOverride = null;

  var config = {
    PRODUCTION_WEB_APP_URL: PRODUCTION_WEB_APP_URL,
    WEB_APP_URL: PRODUCTION_WEB_APP_URL,
    LOCAL_DEV_URL: 'http://localhost:3000/api',
    SCHOOL_TIMEZONE: 'Indian/Maldives',
    AUTO_REFRESH_INTERVAL_MS: 10000,
    THRESHOLDS: {
      WARNING_MINUTES: 10,
      CRITICAL_MINUTES: 15
    },

    getWebAppUrl: function () {
      if (typeof window !== 'undefined' && window.SMT_RUNTIME_WEB_APP_URL !== undefined) {
        return String(window.SMT_RUNTIME_WEB_APP_URL).trim();
      }
      if (explicitOverride !== null) {
        return String(explicitOverride).trim();
      }
      if (typeof window !== 'undefined' && window.location && window.location.search) {
        try {
          var params = new URLSearchParams(window.location.search);
          if (params.has('endpoint')) {
            return String(params.get('endpoint')).trim();
          }
          if (params.get('local') === 'true' || params.get('local') === '1') {
            return window.location.origin + '/api';
          }
        } catch (e) {}
      }
      try {
        if (typeof localStorage !== 'undefined') {
          var saved = localStorage.getItem(STORAGE_URL_KEY);
          if (saved !== null && String(saved).trim() !== '') {
            return String(saved).trim();
          }
        }
      } catch (e) {}
      if (typeof window !== 'undefined' && window.location && (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1')) {
        return window.location.origin + '/api';
      }
      return String(config.WEB_APP_URL || PRODUCTION_WEB_APP_URL).trim();
    },

    setWebAppUrl: function (url) {
      var clean = String(url || '').trim();
      explicitOverride = clean;
      config.WEB_APP_URL = clean;
      try {
        if (typeof localStorage !== 'undefined') {
          if (clean && clean !== PRODUCTION_WEB_APP_URL) {
            localStorage.setItem(STORAGE_URL_KEY, clean);
          } else {
            localStorage.removeItem(STORAGE_URL_KEY);
          }
        }
      } catch (e) {}
      return clean;
    },

    resetToDefault: function () {
      explicitOverride = null;
      config.WEB_APP_URL = PRODUCTION_WEB_APP_URL;
      try {
        if (typeof localStorage !== 'undefined') {
          localStorage.removeItem(STORAGE_URL_KEY);
        }
      } catch (e) {}
      return PRODUCTION_WEB_APP_URL;
    },

    isUsingProduction: function () {
      return config.getWebAppUrl() === PRODUCTION_WEB_APP_URL;
    },

    isUsingLocalOverride: function () {
      var url = config.getWebAppUrl();
      return url.indexOf('localhost') !== -1 || url.indexOf('127.0.0.1') !== -1;
    }
  };

  return config;
});
