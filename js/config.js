/**
 * Student Movement Tracker — Shared Configuration (js/config.js)
 *
 * Stores externalized runtime configuration without inventing or hardcoding
 * fake Google Apps Script deployment URLs or credentials.
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = factory();
  } else {
    root.SMT_CONFIG = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  var STORAGE_URL_KEY = 'smt_apps_script_webapp_url';

  var config = {
    // Set your deployed Google Apps Script Web App URL here or via the UI connection bar / local dev server:
    WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbygZM_HFzRnom0e3AFk4pSyvHufSLyF0hCpPba9oSr9wkyGXQrT-8SsMTU3MoLI-vtB/exec',
    SCHOOL_TIMEZONE: 'Indian/Maldives',
    AUTO_REFRESH_INTERVAL_MS: 10000,
    THRESHOLDS: {
      WARNING_MINUTES: 10,
      CRITICAL_MINUTES: 15
    },

    getWebAppUrl: function () {
      if (typeof window !== 'undefined' && window.SMT_RUNTIME_WEB_APP_URL) {
        return String(window.SMT_RUNTIME_WEB_APP_URL).trim();
      }
      if (config.WEB_APP_URL && String(config.WEB_APP_URL).trim() !== '') {
        return String(config.WEB_APP_URL).trim();
      }
      try {
        if (typeof localStorage !== 'undefined') {
          var saved = localStorage.getItem(STORAGE_URL_KEY);
          if (saved && String(saved).trim() !== '') {
            return String(saved).trim();
          }
        }
      } catch (e) {
        // Ignore storage errors in restricted environments
      }
      return '';
    },

    setWebAppUrl: function (url) {
      var clean = String(url || '').trim();
      config.WEB_APP_URL = clean;
      try {
        if (typeof localStorage !== 'undefined') {
          if (clean) {
            localStorage.setItem(STORAGE_URL_KEY, clean);
          } else {
            localStorage.removeItem(STORAGE_URL_KEY);
          }
        }
      } catch (e) {
        // Ignore storage errors
      }
      return clean;
    }
  };

  return config;
});
