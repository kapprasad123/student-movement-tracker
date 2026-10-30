/**
 * Student Movement Tracker — Mock Google Apps Script & Spreadsheet Runtime (tests/gas-mock.js)
 *
 * Provides a faithful in-memory simulation of:
 * - SpreadsheetApp (Spreadsheet, Sheet, Range with 1-based indexing)
 * - PropertiesService (ScriptProperties)
 * - LockService (ScriptLock with waitLock / releaseLock tracking)
 * - Utilities (SHA_256 digest, HMAC-SHA256 signature, base64EncodeWebSafe, base64DecodeWebSafe, newBlob, getUuid)
 * - ContentService (JSON TextOutput)
 */

const crypto = require('crypto');

class MockRange {
  constructor(sheet, startRow, startCol, numRows, numCols) {
    this.sheet = sheet;
    this.startRow = startRow; // 1-based
    this.startCol = startCol; // 1-based
    this.numRows = numRows;
    this.numCols = numCols;
  }

  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowIdx = this.startRow - 1 + r;
      const srcRow = this.sheet.grid[rowIdx] || [];
      const rowArr = [];
      for (let c = 0; c < this.numCols; c++) {
        const colIdx = this.startCol - 1 + c;
        const val = srcRow[colIdx];
        rowArr.push(val !== undefined ? val : '');
      }
      out.push(rowArr);
    }
    return out;
  }

  setValues(values) {
    if (!Array.isArray(values) || values.length !== this.numRows) {
      throw new Error(`setValues row count mismatch: expected ${this.numRows}, got ${values ? values.length : 0}`);
    }
    for (let r = 0; r < this.numRows; r++) {
      const rowIdx = this.startRow - 1 + r;
      if (!this.sheet.grid[rowIdx]) {
        this.sheet.grid[rowIdx] = [];
      }
      const rowVals = values[r];
      if (!Array.isArray(rowVals) || rowVals.length !== this.numCols) {
        throw new Error(`setValues col count mismatch at row ${r}`);
      }
      for (let c = 0; c < this.numCols; c++) {
        const colIdx = this.startCol - 1 + c;
        this.sheet.grid[rowIdx][colIdx] = rowVals[c];
      }
    }
    return this;
  }
}

class MockSheet {
  constructor(name) {
    this.name = name;
    this.grid = [];
  }

  getName() {
    return this.name;
  }

  getLastRow() {
    return this.grid.length;
  }

  getLastColumn() {
    let maxCols = 0;
    for (let r = 0; r < this.grid.length; r++) {
      if (this.grid[r] && this.grid[r].length > maxCols) {
        maxCols = this.grid[r].length;
      }
    }
    return maxCols;
  }

  getRange(row, col, numRows = 1, numCols = 1) {
    return new MockRange(this, row, col, numRows, numCols);
  }
}

class MockSpreadsheet {
  constructor(id = 'MOCK_SPREADSHEET_ID') {
    this.id = id;
    this.sheets = new Map();
  }

  getId() {
    return this.id;
  }

  getSheetByName(name) {
    return this.sheets.get(name) || null;
  }

  insertSheet(name) {
    if (this.sheets.has(name)) {
      throw new Error(`Sheet "${name}" already exists`);
    }
    const sheet = new MockSheet(name);
    this.sheets.set(name, sheet);
    return sheet;
  }

  getSheets() {
    return Array.from(this.sheets.values());
  }
}

function createGasEnvironment(options = {}) {
  const activeSpreadsheet = options.spreadsheet || new MockSpreadsheet('MOCK_SS_001');
  const scriptProps = new Map(Object.entries(options.scriptProperties || {}));
  let lockAcquireCount = 0;
  let lockReleaseCount = 0;

  const PropertiesService = {
    getScriptProperties() {
      return {
        getProperty(key) {
          return scriptProps.has(key) ? scriptProps.get(key) : null;
        },
        setProperty(key, value) {
          scriptProps.set(String(key), String(value));
        },
        deleteProperty(key) {
          scriptProps.delete(String(key));
        },
        getProperties() {
          return Object.fromEntries(scriptProps.entries());
        }
      };
    }
  };

  const SpreadsheetApp = {
    getActiveSpreadsheet() {
      return activeSpreadsheet;
    },
    openById(id) {
      if (id === activeSpreadsheet.getId()) {
        return activeSpreadsheet;
      }
      return null;
    }
  };

  const LockService = {
    getScriptLock() {
      return {
        waitLock() {
          lockAcquireCount++;
          return true;
        },
        releaseLock() {
          lockReleaseCount++;
        }
      };
    },
    _getStats() {
      return { lockAcquireCount, lockReleaseCount };
    }
  };

  const Utilities = {
    DigestAlgorithm: {
      SHA_256: 'SHA_256'
    },
    Charset: {
      UTF_8: 'UTF-8'
    },
    computeDigest(algo, value) {
      const buf = crypto.createHash('sha256').update(String(value), 'utf8').digest();
      return Array.from(buf).map((b) => (b > 127 ? b - 256 : b));
    },
    computeHmacSha256Signature(value, key) {
      const buf = crypto.createHmac('sha256', String(key)).update(String(value), 'utf8').digest();
      return Array.from(buf).map((b) => (b > 127 ? b - 256 : b));
    },
    base64EncodeWebSafe(str) {
      return Buffer.from(String(str), 'utf8')
        .toString('base64')
        .replace(/\+/g, '-')
        .replace(/\//g, '_');
    },
    base64DecodeWebSafe(b64) {
      const norm = String(b64).replace(/-/g, '+').replace(/_/g, '/');
      const buf = Buffer.from(norm, 'base64');
      return Array.from(buf);
    },
    newBlob(bytes) {
      const u8 = Uint8Array.from(bytes.map((b) => (b < 0 ? b + 256 : b)));
      return {
        getDataAsString(encoding = 'utf8') {
          return Buffer.from(u8).toString(encoding.toLowerCase().replace('-', ''));
        }
      };
    },
    getUuid() {
      return crypto.randomUUID();
    }
  };

  const ContentService = {
    MimeType: {
      JSON: 'application/json'
    },
    createTextOutput(text) {
      let mime = 'text/plain';
      return {
        setMimeType(m) {
          mime = m;
          return this;
        },
        getContent() {
          return text;
        },
        getMimeType() {
          return mime;
        }
      };
    }
  };

  // Install globals for Code.gs execution
  global.PropertiesService = PropertiesService;
  global.SpreadsheetApp = SpreadsheetApp;
  global.LockService = LockService;
  global.Utilities = Utilities;
  global.ContentService = ContentService;

  return {
    spreadsheet: activeSpreadsheet,
    scriptProps,
    PropertiesService,
    SpreadsheetApp,
    LockService,
    Utilities,
    ContentService
  };
}

module.exports = {
  MockSpreadsheet,
  MockSheet,
  MockRange,
  createGasEnvironment
};
