/**
 * E.S. TRAVEL TMS Enterprise v2.0 — Payment Module
 * Payment state and outstanding-day automation.
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
  global.ESTMS.Payment = Object.freeze({
    paymentFor: (...args) => call('paymentFor', ...args),
    paid: (...args) => call('paid', ...args),
    outstanding: (...args) => call('outstanding', ...args),
    autoStatuses: (...args) => call('autoStatuses', ...args),
  });
})(window);
