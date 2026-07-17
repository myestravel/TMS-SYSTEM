/**
 * E.S. TRAVEL TMS Enterprise v2.0 — Sheet Module
 * Google Sheets synchronization façade.
 *
 * Compatibility note: the official legacy runtime remains in script.js.
 * This module exposes a stable namespaced API without changing existing UI behaviour.
 */
(function (global) {
  'use strict';
  const call = (name, ...args) => {
    const fn = global[name];
    if (typeof fn !== 'function') throw new Error(`TMS function not available: ${name}`);
    return fn(...args);
  };
  global.ESTMS = global.ESTMS || { version: '2.0.0' };
  global.ESTMS.Sheet = Object.freeze({
    syncAllToGoogleSheet: (...args) => call('syncAllToGoogleSheet', ...args),
    loadAllFromGoogleSheet: (...args) => call('loadAllFromGoogleSheet', ...args),
    sheetPayload: (...args) => call('sheetPayload', ...args),
  });
})(window);
