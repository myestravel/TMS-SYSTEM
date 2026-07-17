/******************************************************
 * E.S.TRAVEL TMS - Google Sheet Storage + Monthly Report No Overwrite
 * IMPORTANT: Run setupTmsProject() once from the Apps Script editor.
 ******************************************************/

var TMS_SPREADSHEET_PROPERTY = 'TMS_SPREADSHEET_ID';
var TMS_API_VERSION = '2.0.9';

function setupTmsProject() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No active spreadsheet found. Open Extensions > Apps Script from the target Google Sheet, then run setupTmsProject().');
  }
  PropertiesService.getScriptProperties().setProperty(TMS_SPREADSHEET_PROPERTY, ss.getId());
  ensureTmsSettings_(ss);
  SpreadsheetApp.flush();
  var result = {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    message: 'TMS connection completed successfully.'
  };
  console.log(JSON.stringify(result));
  return result;
}

/** Run this after setupTmsProject() to verify the binding without using Spreadsheet UI. */
function testTmsConnection() {
  var ss = getTmsSpreadsheet_();
  var result = {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    spreadsheetUrl: ss.getUrl(),
    timezone: ss.getSpreadsheetTimeZone(),
    checkedAt: new Date().toISOString()
  };
  console.log(JSON.stringify(result));
  return result;
}

function getTmsSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(TMS_SPREADSHEET_PROPERTY);
  if (id) return SpreadsheetApp.openById(id);
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    props.setProperty(TMS_SPREADSHEET_PROPERTY, active.getId());
    return active;
  }
  throw new Error('TMS spreadsheet is not connected. Run setupTmsProject() once from Apps Script.');
}

function ensureTmsSettings_(ss) {
  var sh = ss.getSheetByName('SETTINGS') || ss.insertSheet('SETTINGS');
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,2).setValues([['key','value']]);
  var values = sh.getLastRow() > 1 ? sh.getRange(2,1,sh.getLastRow()-1,2).getValues() : [];
  var keys = {};
  values.forEach(function(r){ keys[String(r[0]||'').toUpperCase()] = true; });
  var add = [];
  if (!keys.SYSTEM_NAME) add.push(['SYSTEM_NAME','E.S. TRAVEL TMS Enterprise v2.0']);
  if (!keys.TIMEZONE) add.push(['TIMEZONE','Asia/Kuala_Lumpur']);
  if (!keys.BOOKING_PREFIX) add.push(['BOOKING_PREFIX','EST']);
  if (!keys.CALENDAR_ID) add.push(['CALENDAR_ID','PRIMARY']);
  if (add.length) sh.getRange(sh.getLastRow()+1,1,add.length,2).setValues(add);
}

function doPost(e) {
  try {
    var raw = '';
    if (e && e.parameter && e.parameter.payload) raw = e.parameter.payload;
    else if (e && e.postData && e.postData.contents) raw = e.postData.contents;
    var payload = JSON.parse(raw || '{}');
    if (payload.action !== 'sync_all') return jsonResponse({ok:false, error:'Invalid action'});
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      syncAllToSheets(payload.data || {});
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
    return jsonResponse({ok:true, message:'Synced to Google Sheet', updatedAt:new Date().toISOString(), bookingCount:(payload.data && payload.data.bookings || []).length});
  } catch (err) {
    return jsonResponse({ok:false, error:String(err), stack:err && err.stack ? String(err.stack) : ''});
  }
}

function doGet(e) {
  try {
    var params = (e && e.parameter) || {};
    var action = params.action || '';
    var callback = params.callback || '';
    if (action === 'health') {
      var ss = getTmsSpreadsheet_();
      var calendar = getTmsCalendar_();
      return jsonOrJsonp_({ok:true, service:'E.S. TRAVEL TMS', apiVersion:TMS_API_VERSION, spreadsheetId:ss.getId(), spreadsheetName:ss.getName(), calendarId:calendar.getId(), calendarName:calendar.getName(), updatedAt:new Date().toISOString()}, callback);
    }
    if (action === 'load_all') {
      var data = readCloudData_();
      return jsonOrJsonp_({ok:true, data:data, updatedAt: data.cloudUpdatedAt || ''}, callback);
    }
    return jsonOrJsonp_({ok:true, message:'TMS Sheet Web App is running', apiVersion:TMS_API_VERSION}, callback);
  } catch (err) {
    return jsonOrJsonp_({ok:false, error:String(err)}, (e && e.parameter && e.parameter.callback) || '');
  }
}

function jsonOrJsonp_(obj, callback) {
  var out = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(String(callback).replace(/[^A-Za-z0-9_.$]/g,'') + '(' + out + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonResponse(obj);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}


/**
 * Normalizes the TMS payload so Apps Script functions are safe when invoked
 * manually from the editor or when an older frontend omits optional arrays.
 * Manual execution will reuse the last cloud snapshot and will never access
 * data.bookings on an undefined value.
 */
function normalizeTmsData_(data) {
  if (!data || typeof data !== 'object') {
    try {
      data = readCloudData_();
    } catch (err) {
      data = {};
    }
  }
  data = data || {};
  data.bookings = Array.isArray(data.bookings) ? data.bookings : [];
  data.prices = Array.isArray(data.prices) ? data.prices : [];
  data.agents = Array.isArray(data.agents) ? data.agents : [];
  data.suppliers = Array.isArray(data.suppliers) ? data.suppliers : [];
  data.drivers = Array.isArray(data.drivers) ? data.drivers : [];
  data.vehicles = Array.isArray(data.vehicles) ? data.vehicles : [];
  return data;
}

/**
 * Safe editor test. This does not create a fake booking and does not clear data.
 * It reports the currently stored cloud snapshot and Calendar connection.
 */
function testSyncEnvironment() {
  var ss = getTmsSpreadsheet_();
  var data = normalizeTmsData_(null);
  var calendar = getTmsCalendar_();
  var result = {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    calendarId: calendar.getId(),
    calendarName: calendar.getName(),
    bookingCount: data.bookings.length,
    message: 'Environment is ready. Create/confirm a Booking from the TMS website to test live synchronization.'
  };
  console.log(JSON.stringify(result));
  return result;
}

function syncAllToSheets(data) {
  var ss = getTmsSpreadsheet_();
  data = normalizeTmsData_(data);
  // Booking/Sheet persistence must never depend on Google Calendar.
  // Calendar authorization or CALENDAR_ID errors are recorded per booking,
  // but the order is still written to Google Sheets first.
  try {
    syncCalendarEvents_(data);
  } catch (calendarErr) {
    var calendarMessage = String(calendarErr && calendarErr.message ? calendarErr.message : calendarErr);
    (data.bookings || []).forEach(function(b){
      if (b) b.calendarSyncError = calendarMessage;
    });
    console.error('Calendar unavailable; continuing Google Sheet save: ' + calendarMessage);
  }
  writeCloudData_(ss, data);

  writeSheet(ss, 'BOOKINGS', [
    'bookingNo','bookingDate','bookingType','status','bookingReference','tripCode','voucherCode','salesPerson','agentName','supplierName','departureDate','departureTime','returnDate','returnTime','needGuide','guideLanguage','noOfPax','noOfLuggage','vehicleType','vehicleQty','specialRequests','driverRemark','internalRemark','originalBookingInfo','calendarDepartureEventId','calendarDepartureUrl','calendarReturnEventId','calendarReturnUrl','createdAt','updatedAt'
  ], (data.bookings || []).map(function(b){
    return [b.bookingNo,b.bookingDate,b.bookingType,b.status,b.bookingReference,b.tripCode,b.voucherCode,b.salesPerson,partyName(data.agents,b.agentId),partyName(data.suppliers,b.supplierId),b.departureDate,b.departureTime,b.returnDate,b.returnTime,b.needGuide,b.guideLanguage,b.noOfPax,b.noOfLuggage,b.vehicleType,b.vehicleQty,(b.specialRequests||[]).join(', '),b.driverRemark,b.internalRemark,b.originalBookingInfo,(b.calendarEvents&&b.calendarEvents.Departure&&b.calendarEvents.Departure.id)||'',(b.calendarEvents&&b.calendarEvents.Departure&&b.calendarEvents.Departure.url)||'',(b.calendarEvents&&b.calendarEvents.Return&&b.calendarEvents.Return.id)||'',(b.calendarEvents&&b.calendarEvents.Return&&b.calendarEvents.Return.url)||'',b.createdAt,b.updatedAt];
  }));

  var locations = [];
  (data.bookings || []).forEach(function(b){
    (b.locations || []).forEach(function(x){
      locations.push([b.bookingNo,x.seq,x.type,cleanAddressForSheet(x.address),x.lat,x.lon,x.placeId,x.addressConfirmed]);
    });
  });
  writeSheet(ss, 'BOOKING_LOCATIONS', ['bookingNo','seq','type','fullAddress','lat','lon','placeId','addressConfirmed'], locations);

  var passengers = [];
  (data.bookings || []).forEach(function(b){
    (b.passengers || []).forEach(function(p){
      passengers.push([b.bookingNo,p.seq,p.name,p.nationality,p.nationalityOther,p.phone]);
    });
  });
  writeSheet(ss, 'BOOKING_PASSENGERS', ['bookingNo','seq','passengerName','nationality','nationalityOther','phone'], passengers);

  var assignments = [];
  (data.bookings || []).forEach(function(b){
    (b.assignments || []).forEach(function(a){
      assignments.push([b.bookingNo,a.segment || 'All Trip',driverName(data.drivers,a.driverId),driverPhone(data.drivers,a.driverId),vehiclePlate(data.vehicles,a.vehicleId),vehicleType(data.vehicles,a.vehicleId),a.remark,a.assignedAt]);
    });
  });
  writeSheet(ss, 'BOOKING_ASSIGNMENTS', ['bookingNo','tripSegment','driverName','driverPhone','vehiclePlate','vehicleType','remark','assignedAt'], assignments);

  var prices = (data.prices || []).map(function(p){
    return [bookingNo(data.bookings,p.bookingId),p.cost,p.selling,p.currency,p.orderStatus,p.paymentReceived,p.paymentMethod,p.updatedAt];
  });
  writeSheet(ss, 'ADMIN_PRICE_PAYMENT', ['bookingNo','cost','selling','currency','orderStatus','paymentReceived','paymentMethod','updatedAt'], prices);

  writeSheet(ss, 'AGENTS', ['id','companyName','phone','email','address','usageCount','lastUsedDate','favorite'], (data.agents || []).map(function(x){ return [x.id,x.companyName,x.phone,x.email,x.address,x.usageCount,x.lastUsedDate,x.favorite]; }));
  writeSheet(ss, 'SUPPLIERS', ['id','companyName','phone','email','address','usageCount','lastUsedDate','favorite'], (data.suppliers || []).map(function(x){ return [x.id,x.companyName,x.phone,x.email,x.address,x.usageCount,x.lastUsedDate,x.favorite]; }));
  writeSheet(ss, 'DRIVERS', ['id','name','phone','licenseNo','status'], (data.drivers || []).map(function(x){ return [x.id,x.name,x.phone,x.licenseNo,x.status]; }));
  writeSheet(ss, 'VEHICLES', ['id','plateNo','vehicleType','seat','status'], (data.vehicles || []).map(function(x){ return [x.id,x.plateNo,x.vehicleType,x.seat,x.status]; }));
  writeDashboardV2_(ss, data);
  appendSyncLog(ss, [[new Date(), (data.bookings||[]).length, (data.agents||[]).length, (data.suppliers||[]).length]]);
}


/**
 * Automatically creates or updates Google Calendar events.
 * One event is maintained for Departure and, where applicable, one for Return.
 * Event IDs are stored inside each booking so Assignment updates the same event.
 */
function syncCalendarEvents_(data) {
  data = normalizeTmsData_(data);
  var calendar = getTmsCalendar_();
  var createdOrUpdated = 0;
  (data.bookings || []).forEach(function(b){
    if (!b || !b.bookingNo) return;
    b.calendarEvents = b.calendarEvents || {};
    b.calendarSyncError = '';
    try {
      if (!b.departureDate || !b.departureTime) {
        throw new Error('Departure Date and Departure Time are required for Calendar Event.');
      }
      upsertBookingCalendarSegment_(calendar, data, b, 'Departure');
      createdOrUpdated++;
      if (isReturnCalendarBooking_(b) && b.returnDate && b.returnTime) {
        upsertBookingCalendarSegment_(calendar, data, b, 'Return');
        createdOrUpdated++;
      }
    } catch (err) {
      b.calendarSyncError = String(err && err.message ? err.message : err);
      console.error('Calendar sync failed for ' + b.bookingNo + ': ' + b.calendarSyncError);
    }
  });
  return createdOrUpdated;
}

function getTmsCalendar_() {
  var ss = getTmsSpreadsheet_();
  var calendarId = '';
  var settings = ss.getSheetByName('SETTINGS');
  if (settings && settings.getLastRow() >= 2) {
    var values = settings.getRange(2,1,settings.getLastRow()-1,2).getValues();
    values.forEach(function(r){
      if (String(r[0] || '').trim().toUpperCase() === 'CALENDAR_ID') calendarId = String(r[1] || '').trim();
    });
  }
  try {
    var calendar = (!calendarId || calendarId.toUpperCase() === 'PRIMARY')
      ? CalendarApp.getDefaultCalendar()
      : CalendarApp.getCalendarById(calendarId);
    if (!calendar) throw new Error('Calendar not found for CALENDAR_ID: ' + (calendarId || 'PRIMARY'));
    return calendar;
  } catch (err) {
    throw new Error('Google Calendar connection failed. Authorize Calendar access and verify SETTINGS!CALENDAR_ID. ' + String(err));
  }
}

function isReturnCalendarBooking_(b) {
  return /round trip|return transfer|vehicle charter/i.test(String(b.bookingType || ''));
}

function upsertBookingCalendarSegment_(calendar, data, b, segment) {
  var isReturn = segment === 'Return';
  var date = isReturn ? b.returnDate : b.departureDate;
  var time = isReturn ? b.returnTime : b.departureTime;
  var start = calendarDateTime_(date, time);
  if (!start) throw new Error('Invalid Calendar date/time: ' + date + ' ' + time);
  var end = new Date(start.getTime() + 60 * 60 * 1000);
  var title = calendarTitle_(data, b, segment);
  var description = calendarDescription_(data, b, segment);
  var location = calendarLocation_(b, segment);
  var existing = b.calendarEvents[segment] || {};
  var event = null;
  if (existing.id) {
    try { event = calendar.getEventById(existing.id); } catch (err) {}
  }
  if (event) {
    event.setTitle(title).setTime(start, end).setDescription(description).setLocation(location || '');
  } else {
    event = calendar.createEvent(title, start, end, {description:description, location:location || ''});
  }
  var eventId = event.getId();
  b.calendarEvents[segment] = {
    id: eventId,
    url: calendarEventUrl_(eventId, calendar.getId()),
    updatedAt: new Date().toISOString()
  };
}

function calendarDateTime_(dateValue, timeValue) {
  var y, m, d, hh, mm;
  if (Object.prototype.toString.call(dateValue) === '[object Date]' && !isNaN(dateValue)) {
    y = dateValue.getFullYear(); m = dateValue.getMonth() + 1; d = dateValue.getDate();
  } else {
    var ds = String(dateValue || '').trim();
    var iso = ds.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    var dmy = ds.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
    if (iso) { y=Number(iso[1]); m=Number(iso[2]); d=Number(iso[3]); }
    else if (dmy) { d=Number(dmy[1]); m=Number(dmy[2]); y=Number(dmy[3]); }
    else return null;
  }
  if (Object.prototype.toString.call(timeValue) === '[object Date]' && !isNaN(timeValue)) {
    hh=timeValue.getHours(); mm=timeValue.getMinutes();
  } else {
    var ts=String(timeValue || '').trim();
    var tm=ts.match(/^(\d{1,2}):(\d{2})(?:\s*(AM|PM))?/i);
    if (!tm) return null;
    hh=Number(tm[1]); mm=Number(tm[2]);
    if (tm[3]) { var ap=tm[3].toUpperCase(); if (ap==='PM' && hh<12) hh+=12; if (ap==='AM' && hh===12) hh=0; }
  }
  var result = new Date(y, m-1, d, hh, mm, 0, 0);
  return isNaN(result.getTime()) ? null : result;
}

function calendarEventUrl_(eventId, calendarId) {
  var eid = Utilities.base64EncodeWebSafe(String(eventId || '') + ' ' + String(calendarId || '')).replace(/=+$/,'');
  return 'https://calendar.google.com/calendar/event?eid=' + encodeURIComponent(eid);
}

function findMaster_(list, id) {
  return (list || []).filter(function(x){ return String(x.id || '') === String(id || ''); })[0] || {};
}

function segmentAssignments_(b, segment) {
  return (b.assignments || []).filter(function(a){
    var s = String(a.segment || 'All').toLowerCase();
    return s === 'all' || s === String(segment || '').toLowerCase();
  });
}

function calendarTitle_(data, b, segment) {
  var agent = findMaster_(data.agents, b.agentId);
  var assignments = segmentAssignments_(b, segment);
  var a = assignments[assignments.length - 1] || {};
  var d = findMaster_(data.drivers, a.driverId);
  var v = findMaster_(data.vehicles, a.vehicleId);
  var parts = [
    agent.companyName || agent.name || 'BOOKING',
    b.bookingNo,
    segment.toUpperCase(),
    v.vehicleType || b.vehicleType || '',
    d.name || '',
    v.plateNo || ''
  ].filter(Boolean);
  return parts.join(' | ').toUpperCase();
}

function calendarDescription_(data, b, segment) {
  var lines = [];
  function section(title, body){ if (String(body || '').trim()) lines.push(title + '\n' + String(body).trim()); }
  var info = [
    'BOOKING NO : ' + (b.bookingNo || ''),
    'BOOKING TYPE : ' + (b.bookingType || ''),
    b.bookingReference ? 'BOOKING REFERENCE : ' + b.bookingReference : '',
    b.voucherCode ? 'VOUCHER CODE : ' + b.voucherCode : '',
    segment === 'Return' ? 'RETURN : ' + b.returnDate + ' ' + b.returnTime : 'DEPARTURE : ' + b.departureDate + ' ' + b.departureTime
  ].filter(Boolean).join('\n');
  section('BOOKING INFORMATION', info);
  section('ROUTE', calendarRouteText_(b, segment));
  var pax = (b.passengers || []).filter(function(p){return p.name || p.phone || p.nationality;}).map(function(p,i){
    return (i+1) + '. ' + (p.name || '') + (p.phone ? '\n' + p.phone : '') + (p.nationality ? '\n' + (String(p.nationality).toUpperCase()==='OTHER' ? p.nationalityOther : p.nationality) : '');
  }).join('\n\n');
  section('PASSENGER', [(b.noOfPax ? 'PAX : ' + b.noOfPax : ''), (b.noOfLuggage ? 'LUGGAGE : ' + b.noOfLuggage : ''), pax].filter(Boolean).join('\n\n'));
  var requests = (b.specialRequests || []).slice(); if (b.specialRequestOther) requests.push(b.specialRequestOther);
  section('ADDITIONAL REQUEST', requests.join('\n'));
  section('REMARK', [b.driverRemark,b.internalRemark].filter(Boolean).join('\n'));
  var asg = segmentAssignments_(b, segment).map(function(a,i){
    var d=findMaster_(data.drivers,a.driverId), v=findMaster_(data.vehicles,a.vehicleId);
    return (i+1)+'. DRIVER : '+(d.name||'')+(d.phone?' ('+d.phone+')':'')+'\nVEHICLE : '+(v.vehicleType||b.vehicleType||'')+' · '+(v.plateNo||'')+(a.remark?'\nASSIGNMENT REMARK : '+a.remark:'');
  }).join('\n\n');
  section('VEHICLE & DRIVER DETAIL', asg);
  section('ORIGINAL BOOKING INFORMATION', b.originalBookingInfo || '');
  return lines.join('\n\n').toUpperCase();
}

function calendarRouteText_(b, segment) {
  var wanted = segment === 'Return' ? ['Return Pickup','Return Drop-off'] : ['Departure Pickup','Departure Drop-off'];
  return (b.locations || []).filter(function(x){ return wanted.indexOf(String(x.type || '')) >= 0; })
    .sort(function(a,b){return Number(a.seq||0)-Number(b.seq||0);})
    .map(function(x){return (x.type || '') + ' : ' + cleanAddressForSheet(x.address || '');}).join('\n');
}

function calendarLocation_(b, segment) {
  var type = segment === 'Return' ? 'Return Pickup' : 'Departure Pickup';
  var x = (b.locations || []).filter(function(l){return String(l.type || '') === type && l.address;})[0];
  return x ? cleanAddressForSheet(x.address) : '';
}

function writeSheet(ss, name, headers, rows) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clearContents();
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  if (rows.length) sh.getRange(2,1,rows.length,headers.length).setValues(rows);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
}

function writeMonthlySheetNoOverwrite(ss, baseName, headers, rows) {
  var finalName = nextAvailableMonthlySheetName(ss, baseName);
  var sh = ss.insertSheet(finalName);
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  if (rows.length) sh.getRange(2,1,rows.length,headers.length).setValues(rows);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  return finalName;
}

function nextAvailableMonthlySheetName(ss, baseName) {
  baseName = safeSheetName(baseName);
  if (!ss.getSheetByName(baseName)) return baseName;
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  var candidate = safeSheetName(baseName + '_RUN_' + stamp);
  var n = 2;
  while (ss.getSheetByName(candidate)) {
    candidate = safeSheetName(baseName + '_RUN_' + stamp + '_' + n);
    n++;
  }
  return candidate;
}

function appendSyncLog(ss, rows) {
  var sh = ss.getSheetByName('SYNC_LOG') || ss.insertSheet('SYNC_LOG');
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,4).setValues([['syncedAt','bookingCount','agentCount','supplierCount']]).setFontWeight('bold');
  sh.getRange(sh.getLastRow()+1,1,rows.length,4).setValues(rows);
  sh.autoResizeColumns(1,4);
}

function partyName(list,id){ var x=(list||[]).find(function(i){return i.id===id;}); return x ? x.companyName || x.name || '' : ''; }
function bookingNo(list,id){ var x=(list||[]).find(function(i){return i.id===id;}); return x ? x.bookingNo || '' : ''; }
function driverName(list,id){ var x=(list||[]).find(function(i){return i.id===id;}); return x ? x.name || '' : ''; }
function driverPhone(list,id){ var x=(list||[]).find(function(i){return i.id===id;}); return x ? x.phone || '' : ''; }
function vehiclePlate(list,id){ var x=(list||[]).find(function(i){return i.id===id;}); return x ? x.plateNo || '' : ''; }
function vehicleType(list,id){ var x=(list||[]).find(function(i){return i.id===id;}); return x ? x.vehicleType || '' : ''; }
function cleanAddressForSheet(v){ return String(v || '').replace(/\s*\([-+]?\d+\.\d+\s*,\s*[-+]?\d+\.\d+\)\s*$/,'').trim(); }


function writeCloudData_(ss, data) {
  var sh = ss.getSheetByName('TMS_CLOUD_DATA') || ss.insertSheet('TMS_CLOUD_DATA');
  var clean = {
    bookings: data.bookings || [],
    prices: data.prices || [],
    agents: data.agents || [],
    suppliers: data.suppliers || [],
    drivers: data.drivers || [],
    vehicles: data.vehicles || [],
    cloudUpdatedAt: new Date().toISOString()
  };
  sh.clearContents();
  sh.getRange(1,1,1,2).setValues([['key','json']]).setFontWeight('bold');
  sh.getRange(2,1,1,2).setValues([['data', JSON.stringify(clean)]]);
  sh.hideSheet();
}

function readCloudData_() {
  var ss = getTmsSpreadsheet_();
  var sh = ss.getSheetByName('TMS_CLOUD_DATA');
  if (sh && sh.getLastRow() >= 2) {
    var json = sh.getRange(2,2).getValue();
    if (json) {
      try { return JSON.parse(json); } catch(err) {}
    }
  }
  return readDataFromSheets_(ss);
}

function readDataFromSheets_(ss) {
  return {
    bookings: [],
    prices: [],
    agents: sheetObjects(ss, 'AGENTS').map(function(x){ return {id:x.id, companyName:x.companyName, phone:x.phone, email:x.email, address:x.address, usageCount:Number(x.usageCount||0), lastUsedDate:x.lastUsedDate, favorite:String(x.favorite).toLowerCase()==='true'}; }),
    suppliers: sheetObjects(ss, 'SUPPLIERS').map(function(x){ return {id:x.id, companyName:x.companyName, phone:x.phone, email:x.email, address:x.address, usageCount:Number(x.usageCount||0), lastUsedDate:x.lastUsedDate, favorite:String(x.favorite).toLowerCase()==='true'}; }),
    drivers: [],
    vehicles: [],
    cloudUpdatedAt: ''
  };
}


/** Creates and immediately deletes a test event to confirm Calendar authorization. */
function testCalendarConnection() {
  var calendar = getTmsCalendar_();
  var start = new Date(Date.now() + 10 * 60 * 1000);
  var end = new Date(start.getTime() + 10 * 60 * 1000);
  var event = calendar.createEvent('E.S. TRAVEL TMS CALENDAR TEST', start, end, {description:'Temporary authorization test. This event will be deleted automatically.'});
  var result = {ok:true, apiVersion:TMS_API_VERSION, calendarId:calendar.getId(), calendarName:calendar.getName(), eventId:event.getId(), message:'Calendar create permission verified successfully.'};
  event.deleteEvent();
  console.log(JSON.stringify(result));
  return result;
}

/********************* MONTHLY REPORT *********************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('E.S. TRAVEL TMS')
    .addItem('Connect this Google Sheet', 'setupTmsProject')
    .addSeparator()
    .addItem('Create current month report', 'createMonthlyReportsForCurrentMonth')
    .addItem('Create previous month report', 'createMonthlyReportsForPreviousMonth')
    .addSeparator()
    .addItem('Install monthly auto report trigger', 'installMonthlyReportTrigger')
    .addItem('Remove monthly report triggers', 'removeMonthlyReportTriggers')
    .addToUi();
}

function installMonthlyReportTrigger() {
  removeMonthlyReportTriggers();
  ScriptApp.newTrigger('createMonthlyReportsForPreviousMonth')
    .timeBased()
    .onMonthDay(1)
    .atHour(2)
    .create();
  console.log('Done. Monthly report trigger installed. It runs on day 1 every month around 2am and exports previous month.');
}

function removeMonthlyReportTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'createMonthlyReportsForPreviousMonth') ScriptApp.deleteTrigger(t);
  });
}

function createMonthlyReportsForCurrentMonth() {
  var now = new Date();
  createMonthlyReports(now.getFullYear(), now.getMonth() + 1);
}

function createMonthlyReportsForPreviousMonth() {
  var now = new Date();
  var y = now.getFullYear();
  var m = now.getMonth(); // previous month in 1-based style because Jan returns 0
  if (m === 0) { y -= 1; m = 12; }
  createMonthlyReports(y, m);
}

function createMonthlyReports(year, month) {
  var ss = getTmsSpreadsheet_();
  var bookings = sheetObjects(ss, 'BOOKINGS');
  var locations = sheetObjects(ss, 'BOOKING_LOCATIONS');
  var passengers = sheetObjects(ss, 'BOOKING_PASSENGERS');
  var assignments = sheetObjects(ss, 'BOOKING_ASSIGNMENTS');
  var prices = sheetObjects(ss, 'ADMIN_PRICE_PAYMENT');

  var ym = year + '-' + pad2(month);
  var filtered = bookings.filter(function(b){
    var d = parseDateOnly(b.departureDate || b.bookingDate || b.createdAt);
    return d && d.getFullYear() === year && (d.getMonth() + 1) === month;
  });

  var indexRows = [];
  var companyHeaders = [
    'bookingNo','status','bookingType','departureDate','departureTime','returnDate','returnTime','agentName','supplierName','passengers','passengerPhones','noOfPax','noOfLuggage','route','assignedDrivers','assignedVehicles','paymentReceived','paymentMethod','currency','selling','driverRemark','internalRemark'
  ];

  var companyRows = filtered.map(function(b){
    return reportRow(b, locations, passengers, assignments, prices);
  });

  var companyBaseName = 'MONTHLY_COMPANY_' + ym;
  var companyName = writeMonthlySheetNoOverwrite(ss, companyBaseName, companyHeaders, companyRows);
  indexRows.push([new Date(), ym, 'Company monthly all orders', companyName, companyRows.length]);

  var driverMap = {};
  filtered.forEach(function(b){
    assignments.filter(function(a){ return a.bookingNo === b.bookingNo; }).forEach(function(a){
      var dn = String(a.driverName || 'Unassigned Driver').trim() || 'Unassigned Driver';
      if (!driverMap[dn]) driverMap[dn] = [];
      driverMap[dn].push(reportRow(b, locations, passengers, assignments, prices));
    });
  });

  Object.keys(driverMap).sort().forEach(function(driver){
    var base = safeSheetName('DRIVER_' + ym + '_' + driver);
    var safe = writeMonthlySheetNoOverwrite(ss, base, companyHeaders, driverMap[driver]);
    indexRows.push([new Date(), ym, 'Driver monthly itinerary: ' + driver, safe, driverMap[driver].length]);
  });

  appendMonthlyIndex(ss, indexRows);
  SpreadsheetApp.flush();
}

function reportRow(b, locations, passengers, assignments, prices) {
  var ps = passengers.filter(function(p){ return p.bookingNo === b.bookingNo; });
  var asg = assignments.filter(function(a){ return a.bookingNo === b.bookingNo; });
  var pr = prices.filter(function(p){ return p.bookingNo === b.bookingNo; }).slice(-1)[0] || {};
  return [
    b.bookingNo || '',
    b.status || '',
    b.bookingType || '',
    b.departureDate || '',
    b.departureTime || '',
    b.returnDate || '',
    b.returnTime || '',
    b.agentName || '',
    b.supplierName || '',
    ps.map(function(p){return p.passengerName || '';}).filter(Boolean).join('\n'),
    ps.map(function(p){return p.phone || '';}).filter(Boolean).join('\n'),
    b.noOfPax || '',
    b.noOfLuggage || '',
    routeText(b.bookingNo, locations),
    asg.map(function(a){return [a.driverName, a.driverPhone].filter(Boolean).join(' ');}).filter(Boolean).join('\n'),
    asg.map(function(a){return [a.vehiclePlate, a.vehicleType].filter(Boolean).join(' · ');}).filter(Boolean).join('\n'),
    pr.paymentReceived || '',
    pr.paymentMethod || '',
    pr.currency || '',
    pr.selling || '',
    b.driverRemark || '',
    b.internalRemark || ''
  ];
}

function routeText(bookingNo, locations) {
  return locations.filter(function(x){ return x.bookingNo === bookingNo; })
    .sort(function(a,b){ return Number(a.seq || 0) - Number(b.seq || 0); })
    .map(function(x){ return (x.type || '') + ': ' + cleanAddressForSheet(x.fullAddress || ''); })
    .join('\n');
}

function appendMonthlyIndex(ss, rows) {
  var sh = ss.getSheetByName('MONTHLY_REPORT_INDEX') || ss.insertSheet('MONTHLY_REPORT_INDEX');
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,5).setValues([['createdAt','month','reportType','sheetName','rowCount']]).setFontWeight('bold');
  if (rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,5).setValues(rows);
  sh.autoResizeColumns(1,5);
}

function sheetObjects(ss, sheetName) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getRange(1,1,sh.getLastRow(),sh.getLastColumn()).getValues();
  var headers = values.shift().map(function(h){ return String(h || '').trim(); });
  return values.map(function(row){
    var obj = {};
    headers.forEach(function(h,i){ obj[h] = row[i]; });
    return obj;
  });
}

function parseDateOnly(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return v;
  var s = String(v).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(Number(m[1]), Number(m[2])-1, Number(m[3]));
  var d = new Date(s);
  return isNaN(d) ? null : d;
}

function pad2(n){ return String(n).padStart(2,'0'); }
function safeSheetName(name){ return String(name).replace(/[\\\/\?\*\[\]\:]/g,' ').substring(0,99).trim(); }



/********************* ENTERPRISE v2.0 API *********************/
function saveRecord(module, record) { return upsertRecord_(module, record); }
function updateRecord(module, record) { return upsertRecord_(module, record); }
function deleteRecord(module, recordId) {
  var data = readCloudData_();
  var key = moduleKey_(module);
  data[key] = (data[key] || []).filter(function(x){ return String(x.id || x.bookingNo) !== String(recordId); });
  syncAllToSheets(data);
  return {ok:true, action:'delete', module:module, recordId:recordId};
}
function updateDashboard() { var ss=getTmsSpreadsheet_(); var data=readCloudData_(); writeDashboardV2_(ss,data); return {ok:true}; }
function updatePayment(record) { return upsertRecord_('prices', record); }
function updateAssignment(bookingNo, assignment) {
  var data=readCloudData_();
  var b=(data.bookings||[]).filter(function(x){return String(x.bookingNo)===String(bookingNo);})[0];
  if(!b) throw new Error('Booking not found: '+bookingNo);
  b.assignments=b.assignments||[]; b.assignments.push(assignment); b.status='In Progress'; b.updatedAt=new Date().toISOString();
  syncAllToSheets(data); return {ok:true, bookingNo:bookingNo};
}
function upsertRecord_(module, record) {
  var data=readCloudData_(), key=moduleKey_(module), list=data[key]||[];
  var id=String(record.id || record.bookingNo || '');
  var i=list.findIndex(function(x){return String(x.id || x.bookingNo || '')===id;});
  if(i>=0) list[i]=record; else list.push(record); data[key]=list; syncAllToSheets(data);
  return {ok:true, action:i>=0?'update':'save', module:module};
}
function moduleKey_(module){
  var m=String(module||'').toLowerCase();
  var map={booking:'bookings',bookings:'bookings',passenger:'bookings',assignment:'bookings',driver:'drivers',drivers:'drivers',vehicle:'vehicles',vehicles:'vehicles',payment:'prices',payments:'prices',price:'prices',prices:'prices',agent:'agents',agents:'agents',supplier:'suppliers',suppliers:'suppliers'};
  return map[m] || m;
}
function writeDashboardV2_(ss,data){
  var rows=[], today=new Date(); today.setHours(0,0,0,0);
  (data.bookings||[]).forEach(function(b){
    if(String(b.status||'').toLowerCase()==='cancelled') return;
    var dep=parseDateOnly(b.departureDate), ret=parseDateOnly(b.returnDate), end=ret||dep;
    var asg=(b.assignments||[]).some(function(a){return a.driverId&&a.vehicleId;});
    var pr=(data.prices||[]).filter(function(p){return p.bookingId===b.id;}).slice(-1)[0]||{};
    var paid=/^(paid|fully paid|collected)$/i.test(String(pr.paymentReceived||b.paymentStatus||''));
    var days=end?Math.max(0,Math.floor((today.getTime()-new Date(end.getFullYear(),end.getMonth(),end.getDate()).getTime())/86400000)):0;
    var st=String(b.status||'Confirmed'), cat='';
    if(asg&&((dep&&sameDay_(dep,today))||(ret&&sameDay_(ret,today)))) cat="TODAY'S TRIPS";
    else if(dep&&dep>today&&st.toLowerCase()==='in progress') cat='UPCOMING TRIPS';
    else if(!asg&&/^(confirmed|pending)$/i.test(st)) cat='PENDING';
    else if(/trip completed|completed|waiting for payment|take action/i.test(st)&&paid) cat='COMPLETED';
    else if(/trip completed|completed|waiting for payment|take action/i.test(st)&&!paid&&days<=30) cat='WAITING FOR PAYMENT';
    else if(/trip completed|completed|waiting for payment|take action/i.test(st)&&!paid&&days>30) cat='TAKE ACTION';
    rows.push([cat,b.bookingNo,b.bookingType,b.departureDate,b.departureTime,b.returnDate,b.returnTime,st,paid?'Paid':'Unpaid',paid?0:days]);
  });
  rows.sort(function(a,b){ if(a[0]==='TAKE ACTION'&&b[0]==='TAKE ACTION') return Number(b[9])-Number(a[9]); return String(a[3]).localeCompare(String(b[3])); });
  writeSheet(ss,'DASHBOARD',['category','bookingNo','bookingType','departureDate','departureTime','returnDate','returnTime','status','paymentStatus','outstandingDays'],rows);
}
function sameDay_(a,b){return a&&b&&a.getFullYear()===b.getFullYear()&&a.getMonth()===b.getMonth()&&a.getDate()===b.getDate();}
