/**
 * CNCOC Service Builder: hymn words, music (MIDI) and full hymnal pages for the hymn viewer.
 *
 * Add this as a second file in the church's Apps Script project (File + > Script, name it Media), then add ONE
 * line to Code.gs, inside the switch in handle_, just below the line for 'pics':
 *
 *       case 'media': return json_(media_(p));
 *
 * Then Deploy > Manage deployments > Edit (pencil) > Version: New version > Deploy.
 *
 *   GET ?action=media&what=words&n=129 -> {words:{lyrics, refrain, credit, descant, midi (base64)}}
 *   GET ?action=media&what=page&n=129  -> {page: base64 GIF of the full hymnal page}
 *
 * It finds its files on the church Drive by name, wherever they are:
 *   hymn-media.json  words and MIDI for all 885 hymns
 *   Hymn Pages       folder(s) holding 0002.gif … 1030.gif
 */
const MEDIA_FILE_NAME = 'hymn-media.json';
const PAGES_FOLDER_NAME = 'Hymn Pages';

function media_(p) {
  const n = Number(p.n);
  if (!(n >= 1 && n <= 9999 && Math.floor(n) === n)) throw new Error('Bad hymn number');
  if (p.what === 'page') return { ok: true, page: hymnPage_(n) };
  return { ok: true, words: hymnMedia_(n) };
}

// Words and MIDI for one hymn. The big file is read once and each hymn is kept in the script cache for 6 hours.
function hymnMedia_(n) {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('media-' + n);
  if (hit) return JSON.parse(hit);
  const all = JSON.parse(fileNamed_(MEDIA_FILE_NAME).getBlob().getDataAsString('UTF-8'));
  const one = all[String(n)];
  if (!one) throw new Error('Hymn ' + n + ' is not in ' + MEDIA_FILE_NAME);
  const keys = Object.keys(all);
  for (let i = 0; i < keys.length; i += 100) { // keep the rest too, so the next hymn is quick
    const batch = {};
    keys.slice(i, i + 100).forEach(function (k) { batch['media-' + k] = JSON.stringify(all[k]); });
    try { cache.putAll(batch, 21600); } catch (err) { /* cache full: fine */ }
  }
  return one;
}

// The full hymnal page (a two-page spread) for one hymn, e.g. 0129.gif in a "Hymn Pages" folder
function hymnPage_(n) {
  const name = ('000' + n).slice(-4) + '.gif';
  const cache = CacheService.getScriptCache(), key = 'pageid-' + n;
  let file = null;
  const id = cache.get(key);
  if (id) { try { file = DriveApp.getFileById(id); if (file.isTrashed()) file = null; } catch (err) { file = null; } }
  if (!file) {
    const it = DriveApp.searchFiles('title = "' + name + '" and trashed = false');
    while (!file && it.hasNext()) {
      const f = it.next(), parents = f.getParents();
      while (parents.hasNext()) { if (parents.next().getName().indexOf(PAGES_FOLDER_NAME) === 0) { file = f; break; } }
    }
    if (!file) throw new Error('The hymnal page for hymn ' + n + ' is not on the church Drive yet (looked for ' + name + ' in a "' + PAGES_FOLDER_NAME + '" folder)');
    cache.put(key, file.getId(), 21600);
  }
  return Utilities.base64Encode(file.getBlob().getBytes());
}

function fileNamed_(name) {
  const cache = CacheService.getScriptCache(), key = 'fileid-' + name;
  const id = cache.get(key);
  if (id) { try { const f = DriveApp.getFileById(id); if (!f.isTrashed()) return f; } catch (err) { /* moved or deleted */ } }
  const it = DriveApp.getFilesByName(name);
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed()) { cache.put(key, f.getId(), 21600); return f; }
  }
  throw new Error(name + ' is not on the church Drive yet');
}

/** Run this once from the editor (select testMedia, click Run) to check the words, music and pages are found. */
function testMedia() {
  const w = hymnMedia_(129);
  Logger.log('Hymn 129 words: %s characters; MIDI: %s bytes', w.lyrics.length, Utilities.base64Decode(w.midi).length);
  try { Logger.log('Hymn 129 page: %s bytes', Utilities.base64Decode(hymnPage_(129)).length); }
  catch (err) { Logger.log('Page: %s', err.message); }
}
