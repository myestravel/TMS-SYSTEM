/**
 * E.S. TRAVEL TMS Enterprise v2.0 — Booking Module
 * Booking creation, editing, confirmation and booking-number workflow.
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
  global.ESTMS.Booking = Object.freeze({
    nextBookingNo: (...args) => call('nextBookingNo', ...args),
    collectBookingFromForm: (...args) => call('collectBookingFromForm', ...args),
    editBooking: (...args) => call('editBooking', ...args),
    cancelBooking: (...args) => call('cancelBooking', ...args),
    completeBooking: (...args) => call('completeBooking', ...args),
  });
})(window);
