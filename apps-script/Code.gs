/**
 * Mappy bake script — bound to the "Mappy Rome" spreadsheet.
 *
 * What it does:
 *  - doPost: the map app's "＋ Add place" form posts here → appends a row,
 *    then bakes it immediately (Places lookup → fill columns F–O → download
 *    3 photos @900px + thumb @320px → commit them to the GitHub repo).
 *  - "Mappy → Bake new rows" menu in the sheet: bakes any rows added by hand
 *    (rows with a name but no pid).
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
 *     in rome.html.
 */

const SHEET_NAME = 'Rome';
const REPO = 'PBensaid-personal/mappy';
const BRANCH = 'main';
const CITY_DIR = 'photos/rome';

// Sheet columns (1-based): A name, B category, C type, D notes, E home,
// F slug, G pid, H lat, I lng, J rating, K count, L address, M phone,
// N whatsapp, O website
const COL_NAME = 1, COL_SLUG = 6, COL_PID = 7;

function onOpen(){
  SpreadsheetApp.getUi().createMenu('Mappy')
    .addItem('Bake new rows', 'bakeAll')
    .addToUi();
}

function doPost(e){
  const out = { ok:false };
  const lock = LockService.getScriptLock();
  lock.tryLock(30000);
  try{
    const p = (e && e.parameter) || {};
    const name = String(p.name || '').trim();
    if(!name) throw new Error('name is required');
    const sh = sheet();
    sh.appendRow([name, p.category || '', p.type || '', p.notes || '',
                  '', '', '', '', '', '', '', '', '', '']);
    out.ok = true;
    out.row = sh.getLastRow();
    try{
      out.baked = bakeRow(sh, out.row);
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

function bakeAll(){
  const sh = sheet();
  const last = sh.getLastRow();
  let baked = 0, failed = [];
  for(let r = 2; r <= last; r++){
    const hasName = String(sh.getRange(r, COL_NAME).getValue()).trim() !== '';
    const hasPid = String(sh.getRange(r, COL_PID).getValue()).trim() !== '';
    if(!hasName || hasPid) continue;
    try{ if(bakeRow(sh, r)) baked++; }
    catch(err){ failed.push(sh.getRange(r, COL_NAME).getValue() + ': ' + err); }
  }
  SpreadsheetApp.getUi().alert(
    'Baked ' + baked + ' row(s).' + (failed.length ? '\n\nFailed:\n' + failed.join('\n') : ''));
}

function sheet(){
  return SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
}

function props(){
  return PropertiesService.getScriptProperties();
}

function bakeRow(sh, r){
  const KEY = props().getProperty('PLACES_KEY');
  const TOKEN = props().getProperty('GITHUB_TOKEN');
  if(!KEY || !TOKEN) throw new Error('Set PLACES_KEY and GITHUB_TOKEN in Script Properties');

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
    payload: JSON.stringify({ textQuery: name + ', Rome, Italy', pageSize: 1 }),
    muteHttpExceptions: true
  });
  if(resp.getResponseCode() >= 300) throw new Error('Places API: ' + resp.getContentText().slice(0, 180));
  const place = (JSON.parse(resp.getContentText()).places || [])[0];
  if(!place) throw new Error('No Google Maps result for "' + name + '"');

  const slug = uniqueSlug(sh, name, r);

  // 2. Photos: same photo fetched at two widths — no resizing needed
  const photoRefs = (place.photos || []).slice(0, 3);
  photoRefs.forEach(function(ph, i){
    ghPutFile(CITY_DIR + '/' + slug + '-' + (i + 1) + '.jpg',
      fetchPhoto(ph.name, 900, KEY), TOKEN, 'Add photo ' + (i + 1) + ' for ' + name);
  });
  if(photoRefs[0]){
    ghPutFile(CITY_DIR + '/' + slug + '-thumb.jpg',
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
