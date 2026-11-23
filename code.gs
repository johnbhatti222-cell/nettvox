/**
 * ISP Confluence Map - backend (Google Apps Script, bound to the "ISP Confluence Map - Data" sheet)
 *
 * Sheet layout (tab "Data", header in row 1):
 *   Building Name | Area | City | Latitude | Longitude | iON Network | Nettvox | Satellite Netcom | Location Verified
 * Any extra numeric column between Longitude and Location Verified is treated as another ISP.
 *
 * Setup (once):
 *   1. Open the sheet > Extensions > Apps Script, paste this file as Code.gs (replace what is there).
 *   2. Edit the passcode in setPasscode(), select setPasscode in the toolbar, Run, approve the permissions.
 *   3. Deploy > New deployment > Web app: Execute as "Me", Who has access "Anyone".
 *      (When you change this file later: Deploy > Manage deployments > Edit > Version "New version",
 *       so the /exec URL stays the same.)
 *   4. Copy the /exec URL into API_URL near the top of the script in index.html on GitHub.
 */

var SHEET_NAME = 'Data';
var DEMARCATION_SHEET = 'Demarcation';
var MAX_FAILED_LOGINS = 10;           // per 10 minutes, then the API locks for 10 minutes

/* ------------------------------------------------------------------ setup */

function setPasscode() {
  var passcode = 'change-me';         // <-- edit, then Run this function once
  PropertiesService.getScriptProperties().setProperty('PASSCODE', passcode);
  Logger.log('Passcode saved.');
}

/* ------------------------------------------------------------------ web entry points */

// The dashboard itself lives on GitHub Pages. This project deliberately serves no HTML page, so every
// call has to come through doPost and pass the passcode check.
function doGet(e) {
  return json_({ok: true, data: 'ISP Confluence Map API is running. Open the GitHub Pages dashboard.'});
}

// The GitHub Pages copy of index.html calls this with {action, args, key}.
var API_ACTIONS = {
  getData: function (a) { return getData(); },
  saveUploadedData: function (a) { return saveUploadedData(a[0], a[1]); },
  updateBuildingLocation: function (a) { return updateBuildingLocation(a[0], a[1], a[2], a[3], a[4]); },
  autoLocateBatch: function (a) { return autoLocateBatch(a[0], a[1]); },
  saveDemarcation: function (a) { return saveDemarcation(a[0]); }
};

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ok: false, error: 'Bad request'}); }
  var auth = checkKey_(body.key);
  if (auth) return json_({ok: false, code: auth.code, error: auth.error});
  if (!API_ACTIONS.hasOwnProperty(body.action)) return json_({ok: false, error: 'Unknown action'});
  try {
    var data = API_ACTIONS[body.action](body.args || []);
    return json_({ok: true, data: data});
  } catch (err) {
    return json_({ok: false, error: String(err && err.message ? err.message : err)});
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function checkKey_(key) {
  var pass = PropertiesService.getScriptProperties().getProperty('PASSCODE');
  if (!pass) return {code: 'unauthorized', error: 'No passcode is set yet. Run setPasscode() in Apps Script.'};
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('failed_logins') || 0);
  if (fails >= MAX_FAILED_LOGINS) return {code: 'locked', error: 'Too many wrong passcodes. Try again in 10 minutes.'};
  if (String(key || '') !== pass) {
    cache.put('failed_logins', String(fails + 1), 600);
    return {code: 'unauthorized', error: 'Wrong passcode.'};
  }
  return null;
}

/* ------------------------------------------------------------------ sheet helpers */

function sheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, 9).setValues([['Building Name', 'Area', 'City', 'Latitude', 'Longitude',
      'iON Network', 'Nettvox', 'Satellite Netcom', 'Location Verified']]);
  }
  return sh;
}

// Works out which column holds what, from a header row.
function mapHeader_(head) {
  var h = head.map(function (v) { return String(v || '').trim(); });
  var find = function (re) { for (var i = 0; i < h.length; i++) if (re.test(h[i])) return i; return -1; };
  var m = {
    b: find(/^(building|bldg|society|premises|name)/i),
    a: find(/^(area|locality|location$|zone)/i),
    c: find(/^city/i),
    lat: find(/^lat/i),
    lng: find(/^(lng|lon)/i),
    v: find(/verif/i)
  };
  if (m.b === -1) m.b = find(/name|building/i);
  var used = [m.b, m.a, m.c, m.lat, m.lng, m.v];
  var skip = /total|isps? present|remark|address|pin ?code|^s\.? ?no|^sr|^#|^id$|operator|dominant|winner|zone/i;
  m.isps = [];
  h.forEach(function (name, i) { if (name && used.indexOf(i) === -1 && !skip.test(name)) m.isps.push({name: name, col: i}); });
  return m;
}

function verifiedCode_(v) {
  v = String(v || '').trim().toLowerCase();
  if (v === 'yes' || v === 'y' || v === 'true' || v === 'verified') return 'yes';
  if (v === 'auto') return 'auto';
  if (v === 'tried') return 'tried';
  return '';
}
function verifiedLabel_(code) { return {yes: 'Yes', auto: 'Auto', tried: 'Tried'}[code] || ''; }

function num_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  var n = Number(String(v).replace(/,/g, '').trim());
  return isNaN(n) ? NaN : n;
}
function coord_(v) {
  if (v === '' || v === null || v === undefined) return null;
  var n = Number(String(v).trim());
  return isNaN(n) ? null : n;
}
function key_(b, a, c) {
  return [b, a, c].map(function (s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); }).join('|');
}

// Reads the whole Data tab into {isps:[names], rows:[{b,a,c,lat,lng,s:[counts],v}]}
function readAll_() {
  var sh = sheet_();
  var values = sh.getDataRange().getValues();
  if (values.length < 1) return {isps: [], rows: []};
  var m = mapHeader_(values[0]);
  var isps = m.isps.map(function (x) { return x.name; });
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var b = String(row[m.b] || '').trim();
    if (!b) continue;
    rows.push({
      b: b,
      a: m.a > -1 ? String(row[m.a] || '').trim() : '',
      c: m.c > -1 ? String(row[m.c] || '').trim() : '',
      lat: m.lat > -1 ? coord_(row[m.lat]) : null,
      lng: m.lng > -1 ? coord_(row[m.lng]) : null,
      s: m.isps.map(function (x) { var n = num_(row[x.col]); return isNaN(n) ? 0 : n; }),
      v: m.v > -1 ? verifiedCode_(row[m.v]) : ''
    });
  }
  return {isps: isps, rows: rows};
}

function writeAll_(isps, rows) {
  var sh = sheet_();
  var head = ['Building Name', 'Area', 'City', 'Latitude', 'Longitude'].concat(isps).concat(['Location Verified']);
  var out = [head];
  rows.forEach(function (r) {
    out.push([r.b, r.a, r.c, r.lat === null ? '' : r.lat, r.lng === null ? '' : r.lng].concat(r.s).concat([verifiedLabel_(r.v)]));
  });
  sh.clearContents();
  sh.getRange(1, 1, out.length, head.length).setValues(out);
  sh.getRange(2, 4, Math.max(out.length - 1, 1), 2).setNumberFormat('0.000000');
  sh.setFrozenRows(1);
  return out.length - 1;
}

/* ------------------------------------------------------------------ API: read */

function getData() {
  var d = readAll_();
  d.updated = new Date().toISOString();
  return d;
}

/* ------------------------------------------------------------------ API: upload */

/**
 * aoa: array of rows from the uploaded Excel/CSV (first row = header)
 * mode: 'replace' | 'merge'
 */
function saveUploadedData(aoa, mode) {
  if (!aoa || aoa.length < 2) throw new Error('The file has no data rows.');
  var m = mapHeader_(aoa[0]);
  if (m.b === -1) throw new Error('Could not find a "Building Name" column in the file.');
  if (m.a === -1) throw new Error('Could not find an "Area" column in the file.');
  if (!m.isps.length) throw new Error('Could not find any ISP subscriber columns (e.g. iON Network, Nettvox, Satellite Netcom).');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var cur = mode === 'merge' ? readAll_() : {isps: [], rows: []};
    // ISP list = existing ISPs + any new ones in the file (matched case-insensitively)
    var isps = cur.isps.slice();
    var lower = isps.map(function (n) { return n.toLowerCase(); });
    var fileIdx = m.isps.map(function (x) {
      var i = lower.indexOf(x.name.toLowerCase());
      if (i === -1) { isps.push(x.name); lower.push(x.name.toLowerCase()); i = isps.length - 1; }
      return i;
    });
    var rows = cur.rows;
    rows.forEach(function (r) { while (r.s.length < isps.length) r.s.push(0); });
    var index = {};
    rows.forEach(function (r, i) { index[key_(r.b, r.a, r.c)] = i; });

    var added = 0, updated = 0, skipped = [];
    for (var i = 1; i < aoa.length; i++) {
      var row = aoa[i];
      var b = String(row[m.b] || '').trim();
      if (!b) { skipped.push({row: i + 1, reason: 'Building name is empty'}); continue; }
      var counts = isps.map(function () { return 0; }), bad = '';
      m.isps.forEach(function (x, j) {
        var n = num_(row[x.col]);
        if (isNaN(n) || n < 0) bad = 'Subscriber count for ' + x.name + ' is not a valid number ("' + row[x.col] + '")';
        else counts[fileIdx[j]] = Math.round(n);
      });
      if (bad) { skipped.push({row: i + 1, reason: bad}); continue; }
      var lat = m.lat > -1 ? coord_(row[m.lat]) : null, lng = m.lng > -1 ? coord_(row[m.lng]) : null;
      if ((lat !== null && Math.abs(lat) > 90) || (lng !== null && Math.abs(lng) > 180)) {
        skipped.push({row: i + 1, reason: 'Latitude/longitude out of range'}); continue;
      }
      var rec = {
        b: b,
        a: String(row[m.a] || '').trim(),
        c: m.c > -1 ? String(row[m.c] || '').trim() : '',
        lat: lat, lng: lng, s: counts,
        v: m.v > -1 ? verifiedCode_(row[m.v]) : (lat !== null && lng !== null ? 'yes' : '')
      };
      var k = key_(rec.b, rec.a, rec.c);
      if (index.hasOwnProperty(k)) {
        var old = rows[index[k]];
        m.isps.forEach(function (x, j) { old.s[fileIdx[j]] = counts[fileIdx[j]]; });
        if (rec.lat !== null && rec.lng !== null) { old.lat = rec.lat; old.lng = rec.lng; old.v = rec.v || 'yes'; }
        updated++;
      } else {
        index[k] = rows.length; rows.push(rec); added++;
      }
    }
    fillApproxCoords_(rows);
    var written = writeAll_(isps, rows);
    return {written: written, added: added, updated: updated, skipped: skipped};
  } finally {
    lock.releaseLock();
  }
}

// Rows without coordinates get the centre of their area (or city) so they still count; they stay "approximate".
function fillApproxCoords_(rows) {
  var acc = {};
  var add = function (k, r) { if (!acc[k]) acc[k] = [0, 0, 0]; acc[k][0] += r.lat; acc[k][1] += r.lng; acc[k][2]++; };
  rows.forEach(function (r) {
    if (r.lat === null || r.lng === null) return;
    add('a|' + key_('', r.a, r.c), r); add('c|' + String(r.c).toLowerCase(), r);
  });
  rows.forEach(function (r, i) {
    if (r.lat !== null && r.lng !== null) return;
    var c = acc['a|' + key_('', r.a, r.c)] || acc['c|' + String(r.c).toLowerCase()];
    if (!c) return;
    // small deterministic spread so pins in one area do not sit on top of each other
    var ang = i * 2.399963, rad = 0.0015 + (i % 7) * 0.0004;
    r.lat = +(c[0] / c[2] + Math.sin(ang) * rad).toFixed(6);
    r.lng = +(c[1] / c[2] + Math.cos(ang) * rad).toFixed(6);
    r.v = '';
  });
}

/* ------------------------------------------------------------------ API: fix one location */

function updateBuildingLocation(b, a, c, lat, lng) {
  lat = Number(lat); lng = Number(lng);
  if (isNaN(lat) || isNaN(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) throw new Error('Invalid coordinates.');
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = sheet_();
    var values = sh.getDataRange().getValues();
    var m = mapHeader_(values[0]);
    if (m.lat === -1 || m.lng === -1) throw new Error('The sheet has no Latitude/Longitude columns.');
    var vCol = m.v;
    if (vCol === -1) { vCol = values[0].length; sh.getRange(1, vCol + 1).setValue('Location Verified'); }
    var k = key_(b, a, c);
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      if (key_(row[m.b], m.a > -1 ? row[m.a] : '', m.c > -1 ? row[m.c] : '') === k) {
        sh.getRange(r + 1, m.lat + 1).setValue(+lat.toFixed(6));
        sh.getRange(r + 1, m.lng + 1).setValue(+lng.toFixed(6));
        sh.getRange(r + 1, vCol + 1).setValue('Yes');
        return {lat: +lat.toFixed(6), lng: +lng.toFixed(6)};
      }
    }
    throw new Error('Building not found in the sheet: ' + b);
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------------ API: free geocoding (no API key) */

/**
 * Uses Apps Script's built-in Maps geocoder (free, daily quota, no key).
 * Only precise results within 4 km of the area's centre are accepted.
 */
function autoLocateBatch(limit, retryFailed) {
  limit = Math.min(Number(limit) || 100, 200);
  var started = Date.now();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var d = readAll_();
    var rows = d.rows;
    // area centres from trusted points, else from all points
    var cen = {};
    rows.forEach(function (r) {
      if (r.lat === null || r.lng === null) return;
      var k = key_('', r.a, r.c);
      if (!cen[k]) cen[k] = {t: [0, 0, 0], all: [0, 0, 0]};
      var trusted = r.v === 'yes' || r.v === 'auto';
      [cen[k].all].concat(trusted ? [cen[k].t] : []).forEach(function (x) { x[0] += r.lat; x[1] += r.lng; x[2]++; });
    });
    var centre = function (r) {
      var c = cen[key_('', r.a, r.c)]; if (!c) return null;
      var x = c.t[2] >= 3 ? c.t : c.all; return {lat: x[0] / x[2], lng: x[1] / x[2]};
    };
    var nIsp = function (r) { return r.s.filter(function (v) { return v > 0; }).length; };
    var total = function (r) { return r.s.reduce(function (a, b) { return a + b; }, 0); };
    var todo = rows.filter(function (r) { return r.v === '' || (retryFailed && r.v === 'tried'); })
      .sort(function (x, y) { return nIsp(y) - nIsp(x) || total(y) - total(x); });

    var geocoder = Maps.newGeocoder().setRegion('in');
    var processed = 0, located = 0, failed = 0, quotaHit = false;
    var good = ['premise', 'subpremise', 'establishment', 'point_of_interest', 'street_address', 'route'];
    for (var i = 0; i < todo.length && processed < limit; i++) {
      if (Date.now() - started > 270000) break;            // stay under the 6-minute limit
      var r = todo[i], c = centre(r);
      var res;
      try {
        res = geocoder.geocode([r.b, r.a, r.c, 'India'].filter(String).join(', '));
      } catch (err) {
        if (/quota|limit|too many/i.test(String(err))) { quotaHit = true; break; }
        res = null;
      }
      processed++;
      var hit = null;
      if (res && res.status === 'OK') {
        for (var j = 0; j < res.results.length && !hit; j++) {
          var g = res.results[j], loc = g.geometry.location;
          var precise = g.geometry.location_type === 'ROOFTOP' || g.types.some(function (t) { return good.indexOf(t) !== -1; });
          var near = !c || distKm_(c.lat, c.lng, loc.lat, loc.lng) <= 4;
          if (precise && near) hit = loc;
        }
      } else if (res && res.status === 'OVER_QUERY_LIMIT') { quotaHit = true; processed--; break; }
      if (hit) { r.lat = +hit.lat.toFixed(6); r.lng = +hit.lng.toFixed(6); r.v = 'auto'; located++; }
      else { r.v = 'tried'; failed++; }
      Utilities.sleep(60);
    }
    writeAll_(d.isps, rows);
    var remaining = rows.filter(function (r) { return r.v === '' || (retryFailed && r.v === 'tried'); }).length;
    return {processed: processed, located: located, failed: failed, remaining: remaining, quotaHit: quotaHit};
  } finally {
    lock.releaseLock();
  }
}

function distKm_(a, b, c, d) {
  var R = 6371, p = Math.PI / 180;
  var x = Math.sin((c - a) * p / 2), y = Math.sin((d - b) * p / 2);
  return 2 * R * Math.asin(Math.sqrt(x * x + Math.cos(a * p) * Math.cos(c * p) * y * y));
}

/* ------------------------------------------------------------------ API: save the demarcation result */

/**
 * Writes the demarcation computed in the dashboard to a "Demarcation" tab so the field team
 * can filter it in Sheets. rows = array of arrays, first row = header.
 */
function saveDemarcation(rows) {
  if (!rows || rows.length < 2) throw new Error('Nothing to save.');
  var width = rows[0].length;
  rows.forEach(function (r) { if (r.length !== width) throw new Error('Malformed demarcation data.'); });
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(DEMARCATION_SHEET) || ss.insertSheet(DEMARCATION_SHEET);
  sh.clearContents();
  sh.getRange(1, 1, rows.length, width).setValues(rows);
  sh.setFrozenRows(1);
  return {written: rows.length - 1, savedAt: new Date().toISOString()};
}
