/**
 * E.S. TRAVEL TMS Enterprise v2.0 — Assignment Module
 * Driver and vehicle assignment workflow.
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
  global.ESTMS.Assignment = Object.freeze({
    getSelectedAssignmentBooking: (...args) => call('getSelectedAssignmentBooking', ...args),
    findAssignmentConflict: (...args) => call('findAssignmentConflict', ...args),
    renderAssignmentPage: (...args) => call('renderAssignmentPage', ...args),
  });
})(window);
