/**
 * Mappy bake script — bound to the "Mappy Rome" spreadsheet (one tab per city:
 * Rome, Chicago, …).
 *
 * What it does:
 *  - doPost: each map page's "＋ Add place" form posts here with city=<tab> →
 *    appends a row to that city's tab, then bakes it immediately (Places
 *    lookup → fill columns F–O → download 3 photos @900px + thumb @320px →
 *    commit them to the GitHub repo under that city's photos dir).
 *  - "Mappy → Bake new rows" menu in the sheet: bakes any rows added by hand
 *    (rows with a name but no pid) on whichever city tab is active.
 *
 * One-time setup:
 *  1. Open the sheet → Extensions → Apps Script → paste this file as Code.gs.
 *  2. Project Settings → Script Properties, add:
 *       PLACES_KEY   — API key from the mappymakesmaps project, restricted to
 *                      "Places API (New)" only (no referer restriction).
 *       GITHUB_TOKEN — fine-grained GitHub PAT, repo PBensaid-personal/mappy,
 *                      permission: Contents = Read and write. Nothing else.
 *  3. Deploy → New deployment → Web app: Execute as = Me,
 *     Who has access = Anyone. Copy the /exec URL into APPS_SCRIPT_URL
 *     in rome.html / chicago.html.
 *
 * Adding a city later: add its tab (copy the Rome header row), add an entry
 * to CITIES below, paste the updated file here, then Deploy → Manage
 * deployments → ✎ → Version: New version (the /exec URL stays the same).
 */

// One tab per city in this spreadsheet; each map page posts city=<name>.
// Requests without a city keep working against Rome (older rome.html).
const CITIES = {
  Rome:    { dir: 'photos/rome',    querySuffix: ', Rome, Italy' },
  Chicago: { dir: 'photos/chicago', querySuffix: ', Chicago, IL' },
  'Fillmore & Ventura': { dir: 'photos/fillmore-ventura', querySuffix: ', Ventura County, CA' }
};
const DEFAULT_CITY = 'Rome';
const REPO = 'PBensaid-personal/mappy';
const BRANCH = 'main';

// Sheet columns (1-based): A name, B category, C type, D notes, E home,
// F slug, G pid, H lat, I lng, J rating, K count, L address, M phone,
// N whatsapp, O website
const COL_NAME = 1, COL_SLUG = 6, COL_PID = 7, COL_DAY = 16;

function onOpen(){
  SpreadsheetApp.getUi().createMenu('Mappy')
    .addItem('Bake new rows', 'bakeAll')
    .addToUi();
}

function doPost(e){
  const p = (e && e.parameter) || {};
  if(p.action === 'reorder') return handleReorder(p);
  const out = { ok:false };
  const lock = LockService.getScriptLock();
  lock.tryLock(30000);
  try{
    const city = cityFor(p);
    const name = String(p.name || '').trim();
    if(!name) throw new Error('name is required');
    const sh = sheet(city);
    sh.appendRow([name, p.category || '', p.type || '', p.notes || '',
                  '', '', '', '', '', '', '', '', '', '']);
    out.ok = true;
    out.row = sh.getLastRow();
    try{
      out.baked = bakeRow(sh, out.row, city);
    }catch(bakeErr){
      out.baked = false;
      out.bakeError = String(bakeErr);
    }
  }catch(err){
    out.error = String(err);
  }finally{
    lock.releaseLock();
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// Reorder rows to match an incoming list of {k, d} — k = slug (or name), d = Day
// value to assign. Preserves every column and rewrites the Day column from the
// payload; rows not named keep their relative order and day at the bottom.
function handleReorder(p){
  const out = { ok:false };
  const lock = LockService.getScriptLock();
  lock.tryLock(30000);
  try{
    const order = JSON.parse(p.order || '[]');
    if(!Array.isArray(order) || !order.length) throw new Error('empty order');
    const sh = sheet(cityFor(p));
    const last = sh.getLastRow();
    if(last < 2) throw new Error('no rows');
    const width = sh.getLastColumn();
    const rng = sh.getRange(2, 1, last - 1, width);
    const rows = rng.getValues();
    const keyOf = r => String((r[COL_SLUG - 1] || r[COL_NAME - 1]) || '');
    const byKey = new Map();
    rows.forEach(r => { const k = keyOf(r); if(k && !byKey.has(k)) byKey.set(k, r); });
    const seen = {};
    const newRows = [];
    order.forEach(item => {
      const key = String(item && item.k != null ? item.k : item);
      if(byKey.has(key) && !seen[key]){
        const row = byKey.get(key);
        if(width >= COL_DAY) row[COL_DAY - 1] = (item && item.d != null) ? String(item.d) : '';
        newRows.push(row);
        seen[key] = true;
      }
    });
    rows.forEach(r => { const k = keyOf(r); if(!seen[k]){ newRows.push(r); seen[k] = true; } });
    if(newRows.length !== rows.length) throw new Error('row count mismatch');
    rng.setValues(newRows);
    out.ok = true;
    out.reordered = newRows.length;
  }catch(err){
    out.error = String(err);
  }finally{
    lock.releaseLock();
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// Bakes the tab currently open in the sheet UI — switch to a city tab first.
function bakeAll(){
  const sh = SpreadsheetApp.getActiveSheet();
  const city = sh.getName();
  if(!CITIES[city]){
    SpreadsheetApp.getUi().alert('No city config for tab "' + city + '" — add it to CITIES in Code.gs.');
    return;
  }
  const last = sh.getLastRow();
  let baked = 0, failed = [];
  for(let r = 2; r <= last; r++){
    const hasName = String(sh.getRange(r, COL_NAME).getValue()).trim() !== '';
    const hasPid = String(sh.getRange(r, COL_PID).getValue()).trim() !== '';
    if(!hasName || hasPid) continue;
    try{ if(bakeRow(sh, r, city)) baked++; }
    catch(err){ failed.push(sh.getRange(r, COL_NAME).getValue() + ': ' + err); }
  }
  SpreadsheetApp.getUi().alert(
    'Baked ' + baked + ' ' + city + ' row(s).' + (failed.length ? '\n\nFailed:\n' + failed.join('\n') : ''));
}

function cityFor(p){
  const city = String((p && p.city) || DEFAULT_CITY);
  if(!CITIES[city]) throw new Error('unknown city: ' + city);
  return city;
}

function sheet(city){
  const sh = SpreadsheetApp.getActive().getSheetByName(city);
  if(!sh) throw new Error('no sheet tab named ' + city);
  return sh;
}

function props(){
  return PropertiesService.getScriptProperties();
}

function bakeRow(sh, r, city){
  const KEY = props().getProperty('PLACES_KEY');
  const TOKEN = props().getProperty('GITHUB_TOKEN');
  if(!KEY || !TOKEN) throw new Error('Set PLACES_KEY and GITHUB_TOKEN in Script Properties');
  const cfg = CITIES[city || DEFAULT_CITY];

  const name = String(sh.getRange(r, COL_NAME).getValue()).trim();

  // 1. One-time Places lookup
  const resp = UrlFetchApp.fetch('https://places.googleapis.com/v1/places:searchText', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'X-Goog-Api-Key': KEY,
      'X-Goog-FieldMask': 'places.id,places.displayName,places.location,places.rating,' +
        'places.userRatingCount,places.formattedAddress,places.internationalPhoneNumber,' +
        'places.websiteUri,places.photos'
    },
    payload: JSON.stringify({ textQuery: name + cfg.querySuffix, pageSize: 1 }),
    muteHttpExceptions: true
  });
  if(resp.getResponseCode() >= 300) throw new Error('Places API: ' + resp.getContentText().slice(0, 180));
  const place = (JSON.parse(resp.getContentText()).places || [])[0];
  if(!place) throw new Error('No Google Maps result for "' + name + '"');

  const slug = uniqueSlug(sh, name, r);

  // 2. Photos: same photo fetched at two widths — no resizing needed
  const photoRefs = (place.photos || []).slice(0, 3);
  photoRefs.forEach(function(ph, i){
    ghPutFile(cfg.dir + '/' + slug + '-' + (i + 1) + '.jpg',
      fetchPhoto(ph.name, 900, KEY), TOKEN, 'Add photo ' + (i + 1) + ' for ' + name);
  });
  if(photoRefs[0]){
    ghPutFile(cfg.dir + '/' + slug + '-thumb.jpg',
      fetchPhoto(photoRefs[0].name, 320, KEY), TOKEN, 'Add thumb for ' + name);
  }

  // 3. Write the technical columns back (F–O)
  const phone = place.internationalPhoneNumber || '';
  sh.getRange(r, COL_SLUG, 1, 10).setValues([[
    slug,
    place.id,
    place.location.latitude,
    place.location.longitude,
    place.rating || '',
    place.userRatingCount || '',
    place.formattedAddress || '',
    phone ? "'" + phone : '',   // leading apostrophe: keep the + as text
    '',
    place.websiteUri || ''
  ]]);
  return true;
}

function fetchPhoto(photoName, width, KEY){
  const r = UrlFetchApp.fetch(
    'https://places.googleapis.com/v1/' + photoName + '/media?maxWidthPx=' + width + '&key=' + KEY,
    { followRedirects: true, muteHttpExceptions: true });
  if(r.getResponseCode() >= 300) throw new Error('photo download failed (' + r.getResponseCode() + ')');
  return r.getBlob();
}

function ghPutFile(path, blob, TOKEN, message){
  const url = 'https://api.github.com/repos/' + REPO + '/contents/' + path;
  const headers = { Authorization: 'Bearer ' + TOKEN, Accept: 'application/vnd.github+json' };
  // If the file exists we must pass its sha to overwrite
  let sha = null;
  const head = UrlFetchApp.fetch(url + '?ref=' + BRANCH, { headers: headers, muteHttpExceptions: true });
  if(head.getResponseCode() === 200) sha = JSON.parse(head.getContentText()).sha;
  const payload = { message: message, branch: BRANCH, content: Utilities.base64Encode(blob.getBytes()) };
  if(sha) payload.sha = sha;
  const r = UrlFetchApp.fetch(url, {
    method: 'put', contentType: 'application/json', headers: headers,
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  if(r.getResponseCode() >= 300) throw new Error('GitHub ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 180));
}

function uniqueSlug(sh, name, ownRow){
  let base = String(name).toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 40) || 'place';
  const taken = {};
  const last = sh.getLastRow();
  if(last >= 2){
    sh.getRange(2, COL_SLUG, last - 1, 1).getValues().forEach(function(v, i){
      if(i + 2 !== ownRow && v[0]) taken[String(v[0])] = true;
    });
  }
  let slug = base, n = 2;
  while(taken[slug]) slug = base + '-' + (n++);
  return slug;
}
