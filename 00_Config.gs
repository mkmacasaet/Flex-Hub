// ===================================================================
// FILE: 00_Config.gs
// PROJECT: SISFU Flex Hub
// PURPOSE: Foundation layer. Every other .gs file depends on this.
//          Contains app config, the HTML include utility, the shared
//          column-finder, the password hasher, and sheet accessor helpers
//          that replace 18+ repeated 3-line open patterns in the old code.
// ===================================================================

const CONFIG_SHEET = '_Config';
let globalConfigCache = null;

/**
 * Serves the initial HTML file to the user's browser.
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('SISFU Flex Hub')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Retrieves global IDs from the _Config sheet.
 * PERFORMANCE: Caches the entire config map on first run.
 */
function getConfig(key) {
  if (!globalConfigCache) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
    if (!sheet) return "";
    const data = sheet.getDataRange().getValues();
    globalConfigCache = {};
    for (let i = 0; i < data.length; i++) {
      globalConfigCache[data[i][0]] = data[i][1];
    }
  }
  return globalConfigCache[key] || "";
}

/**
 * UTILITY: Standard Apps Script templating function.
 * Used in Index.html to inject CSS and module HTML files.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===================================================================
// SHARED UTILITIES (DRY replacements — used across all modules)
// ===================================================================

/**
 * colIdx — Shared column finder.
 * Replaces the 3 identical `const getIdx` declarations that were
 * copy-pasted inside getMasterlistColMap, getAccountsColMap, getLoaColMap.
 *
 * @param {string[]} headers - The lowercased header row array.
 * @param {string[]} terms   - Possible header name substrings to match.
 * @returns {number} Column index, or -1 if not found.
 */
const colIdx = (headers, terms) =>
  headers.findIndex(h => terms.some(t => h.includes(t)));

/**
 * hashPassword — Single source of truth for SHA-256 hashing.
 * Replaces 4 copy-pasted digest chains in attemptLogin,
 * updateAccountAdmin, createNewAccount, and changeMyPassword.
 *
 * @param {string} raw - The plain-text password to hash.
 * @returns {string} Hex-encoded SHA-256 hash.
 */
function hashPassword(raw) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(raw).trim())
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// ===================================================================
// SHEET ACCESSOR HELPERS
// Replaces 18+ repeated "getConfig → openById → getSheetByName" blocks.
// Each helper throws a descriptive error if the config key is missing,
// so callers get a clear message instead of a cryptic null crash.
// ===================================================================

function getMasterSheet() {
  const id = getConfig('MasterStudentListId');
  if (!id) throw new Error('MasterStudentListId missing in _Config.');
  const ss = SpreadsheetApp.openById(id);
  return ss.getSheetByName('MasterList') || ss.getSheets()[0];
}

function getDeferralSheet() {
  const id = getConfig('DeferralMonitorID');
  if (!id) throw new Error('DeferralMonitorID missing in _Config.');
  return SpreadsheetApp.openById(id).getSheetByName('_Deferrals');
}

function getLOASheet() {
  const id = getConfig('WithdrawalLOAID');
  if (!id) throw new Error('WithdrawalLOAID missing in _Config.');
  const ss = SpreadsheetApp.openById(id);
  return ss.getSheets()[0];
}

function getStudentIndexSheet() {
  const id = getConfig('StudentIndexID');
  if (!id) throw new Error('StudentIndexID missing in _Config.');
  const ss = SpreadsheetApp.openById(id);
  return ss.getSheetByName('_StudentIndex') || ss.getSheets()[0];
}

function getAYSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_AcademicYears');
}

// -------------------------------------------------------------------
// UNIVERSAL NAME FORMATTER
// -------------------------------------------------------------------
function formatUniversalName(firstName, middleInitial, lastName) {
  const fName = (firstName     || "").trim();
  const lName = (lastName      || "").trim();
  const mi    = (middleInitial || "").trim() ? (middleInitial || "").trim() : "";
  return [fName, mi, lName].filter(Boolean).join(" ");
}
