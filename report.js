/**
 * E.S. TRAVEL TMS Enterprise v2.0 — Report Module
 * Operational report extension point.
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
  global.ESTMS.Report = Object.freeze({
    reportRow: (...args) => call('reportRow', ...args),
    createMonthlyReportsForCurrentMonth: (...args) => call('createMonthlyReportsForCurrentMonth', ...args),
    createMonthlyReportsForPreviousMonth: (...args) => call('createMonthlyReportsForPreviousMonth', ...args),
  });
})(window);
