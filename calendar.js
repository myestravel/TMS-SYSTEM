/**
 * E.S. TRAVEL TMS Enterprise v2.0 — Calendar Module
 * Google Calendar event detail and launch helpers.
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
  global.ESTMS.Calendar = Object.freeze({
    buildCalendarDetails: (...args) => call('buildCalendarDetails', ...args),
    buildCalendarTitle: (...args) => call('buildCalendarTitle', ...args),
    buildGoogleCalendarUrl: (...args) => call('buildGoogleCalendarUrl', ...args),
    openGoogleCalendarForBooking: (...args) => call('openGoogleCalendarForBooking', ...args),
  });
})(window);
