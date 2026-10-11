// ============================================================
// Winsol Presentatievertaler — frontend
// ============================================================
// Al het werk met het PPTX-bestand gebeurt hier, in de browser:
//   1. uitpakken (JSZip), alle alinea's uit de dia's halen;
//   2. enkel de teksten (per alinea, opgesplitst in "runs") naar de Worker;
//   3. vertaalde tekst terugzetten in exact dezelfde runs, taalcode aanpassen;
//   4. opnieuw inpakken in dezelfde volgorde en aanbieden als download.
// Er wordt niets bewaard: het bestand verlaat de browser niet, en de
// vertaling bestaat enkel in het geheugen van dit tabblad.

// URL van de Cloudflare Worker (zie worker.js / README.md).
const API_URL = 'https://pptx-vertaler.gwenn-vanthournout.workers.dev/';

const MAX_FILE_MB = 30;
// Grens op de hoeveelheid tekst per presentatie: dit bepaalt de AI-kost,
// niet de bestandsgrootte (afbeeldingen worden nooit verstuurd).
const MAX_TOTAL_CHARS = 150000;
// Grootte van één aanroep naar de Worker (moet onder de grenzen in worker.js blijven).
const BATCH_MAX_ITEMS = 80;
const BATCH_MAX_CHARS = 8000;
const PARALLEL_CALLS = 4;

const LANGS = {
  nl: { label: 'Nederlands', tag: 'nl-BE', code: 'NL' },
  fr: { label: 'Frans', tag: 'fr-BE', code: 'FR' },
  de: { label: 'Duits', tag: 'de-DE', code: 'DE' },
  en: { label: 'Engels (Brits)', tag: 'en-GB', code: 'EN' },
  es: { label: 'Spaans', tag: 'es-ES', code: 'ES' },
  it: { label: 'Italiaans', tag: 'it-IT', code: 'IT' },
};
// Naam van een taal in de taal van de app.
const langLabel = (code) => t('lang_' + code);
const GLOSSARY_HEADERS = {
  nl: 'Nederlands', fr: 'Frans', de: 'Duits', en: 'Engels', es: 'Spaans', it: 'Italiaans',
};
// Kolomkoppen die we in de woordenlijst-Excel herkennen.
const HEADER_MAP = {
  nederlands: 'nl', dutch: 'nl', nl: 'nl', vlaams: 'nl',
  frans: 'fr', french: 'fr', 'français': 'fr', francais: 'fr', fr: 'fr',
  duits: 'de', german: 'de', deutsch: 'de', de: 'de',
  engels: 'en', english: 'en', en: 'en',
  spaans: 'es', spanish: 'es', 'español': 'es', espanol: 'es', es: 'es',
  italiaans: 'it', italian: 'it', italiano: 'it', italien: 'it', italienisch: 'it', it: 'it',
};

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const PUA_RE = /[-]/gu;
const LETTER_RE = /\p{L}/u;

// ============================================================
// Auth — overgenomen van Feedbackloop
// ============================================================
const TOKEN_KEY = 'pptxvertaler_token';
let authToken = '';
let authRole = '';

function authHeaders() {
  return authToken ? { 'X-Auth-Token': authToken } : {};
}

const $ = (id) => document.getElementById(id);
const loginScreen = $('loginScreen');
const appRoot = $('appRoot');
const loginUsername = $('loginUsername');
const loginPassword = $('loginPassword');
const loginRemember = $('loginRemember');
const loginSubmitBtn = $('loginSubmitBtn');
const loginError = $('loginError');
const topbar = $('topbar');

async function apiRequest(bodyObj) {
  const res = await fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...authHeaders() },
    body: JSON.stringify(bodyObj),
  });
  let data = {};
  try {
    data = await res.json();
  } catch {
    // geen (geldige) JSON-body
  }
  return { ok: res.ok, status: res.status, data };
}

function handleAuthFailure(status) {
  if (status === 401 || status === 403) {
    showLogin(status === 401 ? t('sessionExpired') : t('noAccess'));
    return true;
  }
  return false;
}

function showApp(username, role) {
  authRole = role;
  loginScreen.style.display = 'none';
  appRoot.style.display = 'block';
  topbar.hidden = false;
  $('topbarUser').textContent = username;
  $('topbarRole').textContent = role === 'admin' ? t('roleAdmin') : t('roleUser');
  const isAdmin = role === 'admin';
  $('glossaryCard').hidden = !isAdmin;
  $('usersCard').hidden = !isAdmin;
  loadGlossaryInfo();
}

function showLogin(message) {
  authToken = '';
  authRole = '';
  try {
    localStorage.removeItem(TOKEN_KEY);
    sessionStorage.removeItem(TOKEN_KEY);
  } catch {
    // storage geblokkeerd (privénavigatie) — geen probleem
  }
  appRoot.style.display = 'none';
  loginScreen.style.display = 'flex';
  topbar.hidden = true;
  loginPassword.value = '';
  loginError.textContent = message || '';
}

async function doLogin() {
  const username = loginUsername.value.trim();
  const password = loginPassword.value;
  const remember = loginRemember.checked;
  loginError.textContent = '';
  if (!username || !password) {
    loginError.textContent = t('loginMissing');
    return;
  }
  loginSubmitBtn.disabled = true;
  loginSubmitBtn.textContent = t('loggingIn');
  try {
    const { ok, status, data } = await apiRequest({ mode: 'login', username, password, remember });
    if (!ok) {
      loginError.textContent = status === 401 ? t('loginFailed') : (data.error || t('loginGenericFail'));
      return;
    }
    authToken = data.token;
    try {
      (remember ? localStorage : sessionStorage).setItem(TOKEN_KEY, authToken);
    } catch {
      // storage geblokkeerd — token blijft enkel in het geheugen
    }
    showApp(data.username, data.role);
  } catch {
    loginError.textContent = t('noConnection');
  } finally {
    loginSubmitBtn.disabled = false;
    loginSubmitBtn.textContent = t('loginBtn');
  }
}

loginSubmitBtn.addEventListener('click', doLogin);
loginPassword.addEventListener('keydown', (e) => { if (e.key === 'Enter') doLogin(); });
loginUsername.addEventListener('keydown', (e) => { if (e.key === 'Enter') loginPassword.focus(); });
$('loginPwToggle').addEventListener('click', () => {
  loginPassword.type = loginPassword.type === 'password' ? 'text' : 'password';
});
$('logoutBtn').addEventListener('click', async () => {
  try {
    await apiRequest({ mode: 'logout' });
  } catch {
    // best-effort
  }
  resetTranslation();
  showLogin();
});

async function tryRestoreSession() {
  let token = '';
  try {
    token = localStorage.getItem(TOKEN_KEY) || sessionStorage.getItem(TOKEN_KEY) || '';
  } catch {
    token = '';
  }
  if (!token) {
    showLogin();
    return;
  }
  authToken = token;
  try {
    const { ok, data } = await apiRequest({ mode: 'auth_me' });
    if (ok) showApp(data.username, data.role);
    else showLogin();
  } catch {
    showLogin(t('noConnectionRelogin'));
  }
}

// ============================================================
// UI-hulpjes
// ============================================================

function setStatus(el, msg, kind) {
  el.textContent = msg || '';
  el.className = 'status' + (kind ? ' ' + kind : '');
}

function wireDropzone(zoneId, inputId, onFile) {
  const zone = $(zoneId);
  const input = $(inputId);
  zone.addEventListener('click', () => input.click());
  zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', (e) => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    if (e.dataTransfer.files.length) onFile(e.dataTransfer.files[0]);
  });
  input.addEventListener('change', (e) => {
    if (e.target.files.length) onFile(e.target.files[0]);
    input.value = '';
  });
}

// Bestandsnaam tonen in de (groene) zone; zonder bestand de uitleg tonen.
function setFileLabel(id, name) {
  const el = $(id);
  const zone = el.closest('.drop-zone');
  if (zone) {
    zone.classList.toggle('has-file', !!name);
    const hint = zone.querySelector('.drop-hint');
    if (hint) hint.hidden = !!name;
  }
  el.hidden = !name;
  el.textContent = name || '';
}

function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// ============================================================
// Taalkeuze + woordenlijst-info
// ============================================================
const sourceLang = $('sourceLang');
const targetLang = $('targetLang');
function fillLangSelects() {
  const keep = [sourceLang.value || 'nl', targetLang.value || 'en'];
  for (const sel of [sourceLang, targetLang]) {
    sel.innerHTML = '';
    for (const code of Object.keys(LANGS)) sel.add(new Option(langLabel(code), code));
  }
  sourceLang.value = keep[0];
  targetLang.value = keep[1];
}
fillLangSelects();
let glossaryInfo = null;

async function loadGlossaryInfo() {
  const el = $('glossaryInfo');
  try {
    const { ok, status, data } = await apiRequest({ mode: 'glossary_info' });
    if (handleAuthFailure(status)) return;
    if (!ok) throw new Error(data.error || t('unknownError'));
    glossaryInfo = data;
  } catch (err) {
    glossaryInfo = null;
    setStatus(el, t('glossaryInfoUnavailable', { msg: err.message }), 'err');
    return;
  }
  updateGlossaryInfoText();
  if (authRole === 'admin') updateGlossaryCurrentText();
}

function formatDate(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  return isNaN(d) ? '' : d.toLocaleDateString(currentUiLang() + '-BE', { day: 'numeric', month: 'long', year: 'numeric' });
}

function updateGlossaryInfoText() {
  const el = $('glossaryInfo');
  if (!glossaryInfo || !glossaryInfo.available) {
    setStatus(el, t('glossaryNone'), 'err');
    return;
  }
  const langs = glossaryInfo.languages || [];
  const missing = [sourceLang.value, targetLang.value].filter((l) => !langs.includes(l));
  if (missing.length) {
    setStatus(el, t('glossaryMissingLang', { langs: missing.map(langLabel).join(' / ') }), 'err');
    return;
  }
  setStatus(el, glossaryInfo.updatedAt
    ? t('glossaryInfoDate', { n: glossaryInfo.terms, date: formatDate(glossaryInfo.updatedAt) })
    : t('glossaryInfo', { n: glossaryInfo.terms }));
}
sourceLang.addEventListener('change', () => { updateGlossaryInfoText(); updateTranslateBtn(); });
targetLang.addEventListener('change', () => { updateGlossaryInfoText(); updateTranslateBtn(); });

// ============================================================
// Stap 1: PPTX inlezen
// ============================================================
let currentFile = null;
let resultBlob = null;
let lastResult = null;
let resultName = '';
let busy = false;

wireDropzone('dropzone', 'fileInput', handlePptxFile);

function handlePptxFile(file) {
  const fileStatus = $('fileStatus');
  resetTranslation();
  currentFile = null;
  setFileLabel('fname', file.name);
  if (!/\.pptx$/i.test(file.name)) {
    setStatus(fileStatus, t('onlyPptx'), 'err');
    updateTranslateBtn();
    return;
  }
  if (file.size > MAX_FILE_MB * 1024 * 1024) {
    setStatus(fileStatus, t('fileTooBig', { size: (file.size / 1048576).toFixed(1), max: MAX_FILE_MB }), 'err');
    updateTranslateBtn();
    return;
  }
  currentFile = file;
  setStatus(fileStatus, t('fileReady', { size: (file.size / 1048576).toFixed(1) }));
  // Taal uit de bestandsnaam afleiden (bv. "..._NL.pptx").
  const m = file.name.match(/[_\- ](NL|FR|DE|EN|ES|IT)(?=[_\- .])/i);
  if (m) {
    const code = m[1].toLowerCase();
    sourceLang.value = code;
    if (targetLang.value === code) targetLang.value = code === 'en' ? 'fr' : 'en';
    updateGlossaryInfoText();
  }
  updateTranslateBtn();
}

function updateTranslateBtn() {
  const btn = $('translateBtn');
  const same = sourceLang.value === targetLang.value;
  btn.disabled = busy || !currentFile || same;
  if (same && currentFile) setStatus($('translateStatus'), t('sameLang'), 'err');
  else if (!busy) setStatus($('translateStatus'), '');
}

function resetTranslation() {
  resultBlob = null;
  lastResult = null;
  resultName = '';
  $('resultCard').style.display = 'none';
  $('reportBody').innerHTML = '';
  $('resultStats').innerHTML = '';
  $('progressBar').hidden = true;
  $('progressFill').style.width = '0%';
}

$('newBtn').addEventListener('click', () => {
  resetTranslation();
  currentFile = null;
  setFileLabel('fname', null);
  setStatus($('fileStatus'), '');
  updateTranslateBtn();
});

$('downloadBtn').addEventListener('click', () => {
  if (!resultBlob) return;
  const url = URL.createObjectURL(resultBlob);
  const a = document.createElement('a');
  a.href = url;
  a.download = resultName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
});

// ============================================================
// PPTX-structuur
// ============================================================

function parseXml(str) {
  const doc = new DOMParser().parseFromString(str, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error(t('xmlError'));
  return doc;
}

function serializeXml(doc) {
  let s = new XMLSerializer().serializeToString(doc);
  s = s.replace(/^<\?xml[^>]*\?>\s*/, '');
  return XML_DECL + s;
}

function partNumber(path) {
  const m = path.match(/(\d+)\.xml$/);
  return m ? parseInt(m[1], 10) : 0;
}

function dirOf(path) {
  return path.slice(0, path.lastIndexOf('/') + 1);
}

function resolveTarget(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (dirOf(base) + target).split('/');
  const out = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p !== '.' && p !== '') out.push(p);
  }
  return out.join('/');
}

function relsPathFor(part) {
  return dirOf(part) + '_rels/' + part.slice(part.lastIndexOf('/') + 1) + '.rels';
}

async function readRels(zip, part) {
  const f = zip.file(relsPathFor(part));
  if (!f) return [];
  const doc = parseXml(await f.async('string'));
  return [...doc.getElementsByTagName('Relationship')].map((r) => ({
    id: r.getAttribute('Id'),
    type: r.getAttribute('Type') || '',
    target: r.getAttribute('TargetMode') === 'External' ? null : resolveTarget(part, r.getAttribute('Target') || ''),
  }));
}

// Geeft de dia's in presentatievolgorde terug, plus per dia de onderdelen
// die erbij horen (SmartArt-diagrammen, sprekersnotities).
async function mapSlides(zip) {
  const pres = parseXml(await zip.file('ppt/presentation.xml').async('string'));
  const rels = await readRels(zip, 'ppt/presentation.xml');
  const relById = Object.fromEntries(rels.map((r) => [r.id, r.target]));
  const order = [];
  const ids = pres.getElementsByTagNameNS(NS_P, 'sldId');
  for (const el of ids) {
    const rid = el.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    if (relById[rid] && zip.file(relById[rid])) order.push(relById[rid]);
  }
  const slides = [];
  for (let i = 0; i < order.length; i++) {
    const path = order[i];
    const srels = await readRels(zip, path);
    slides.push({
      path,
      number: i + 1,
      diagrams: srels.filter((r) => r.target && /\/diagrams\/(data|drawing)\d+\.xml$/.test(r.target) && zip.file(r.target)).map((r) => r.target),
      notes: (srels.find((r) => r.target && /notesSlide\d+\.xml$/.test(r.target) && zip.file(r.target)) || {}).target || null,
      layout: (srels.find((r) => r.target && /slideLayout\d+\.xml$/.test(r.target)) || {}).target || null,
    });
  }
  return slides;
}

// Alle alinea's (<a:p>) van een onderdeel, optioneel enkel buiten
// placeholders (voor lay-outs/masters: daar is placeholder-tekst nooit
// zichtbaar op de dia's).
function collectParagraphs(doc, onlyNonPlaceholder) {
  let paras = [...doc.getElementsByTagNameNS(NS_A, 'p')];
  if (onlyNonPlaceholder) {
    paras = paras.filter((p) => {
      let el = p.parentNode;
      while (el && !(el.namespaceURI === NS_P && el.localName === 'sp')) el = el.parentNode;
      if (!el) return true;
      return el.getElementsByTagNameNS(NS_P, 'ph').length === 0;
    });
  }
  return paras;
}

// Splitst een alinea op in segmenten: tekst-runs (a:r), velden (a:fld) en
// regeleindes (a:br), in volgorde.
function paragraphSegments(p) {
  const segs = [];
  for (const c of p.childNodes) {
    if (c.nodeType !== 1 || c.namespaceURI !== NS_A) continue;
    if (c.localName === 'r' || c.localName === 'fld') {
      const t = c.getElementsByTagNameNS(NS_A, 't')[0];
      segs.push({ kind: c.localName, el: c, tEl: t || null, text: t ? t.textContent : '' });
    } else if (c.localName === 'br') {
      segs.push({ kind: 'br', el: c, tEl: null, text: '\n' });
    }
  }
  return segs;
}

function isFixedSegment(seg) {
  // Wordt nooit vertaald: regeleindes, velden (dianummer, datum) en
  // segmenten zonder letters (spaties, tabs, symbolen, cijfers).
  return seg.kind !== 'r' || !LETTER_RE.test(seg.text);
}

// ============================================================
// Stap 3: vertalen
// ============================================================

$('translateBtn').addEventListener('click', runTranslation);

function setProgress(done, total) {
  $('progressBar').hidden = false;
  $('progressFill').style.width = total ? Math.round((done / total) * 100) + '%' : '0%';
}

async function runTranslation() {
  if (!currentFile || busy) return;
  const src = sourceLang.value;
  const tgt = targetLang.value;
  const includeNotes = $('includeNotes').checked;
  const status = $('translateStatus');
  busy = true;
  updateTranslateBtn();
  resetTranslation();
  setProgress(0, 1);

  try {
    // --- 1. Uitpakken en alinea's verzamelen ---
    setStatus(status, t('unpacking'));
    const zip = await JSZip.loadAsync(currentFile);
    if (!zip.file('ppt/presentation.xml')) throw new Error(t('invalidPptx'));
    const slides = await mapSlides(zip);

    // Onderdelen die we behandelen, met het dianummer voor het rapport.
    const parts = []; // {path, slide, kind, nonPh}
    const seen = new Set();
    const addPart = (path, slide, kind, nonPh) => {
      if (!path || seen.has(path) || !zip.file(path)) return;
      seen.add(path);
      parts.push({ path, slide, kind, nonPh });
    };
    for (const s of slides) {
      addPart(s.path, s.number, 'slide', false);
      for (const d of s.diagrams) addPart(d, s.number, 'diagram', false);
      if (includeNotes && s.notes) addPart(s.notes, s.number, 'notes', false);
    }
    // Zichtbare (niet-placeholder) tekst in de gebruikte lay-outs/masters.
    const layouts = [...new Set(slides.map((s) => s.layout).filter(Boolean))];
    for (const l of layouts) addPart(l, null, 'layout', true);
    for (const l of layouts) {
      const lrels = await readRels(zip, l);
      const master = lrels.find((r) => r.target && /slideMaster\d+\.xml$/.test(r.target));
      if (master) addPart(master.target, null, 'master', true);
    }

    const docs = new Map(); // path -> {doc, part, changed}
    const unique = new Map(); // key -> {id, segs(texts), fixed(bool[]), occurrences:[{segs(objs), part}], slide}
    let totalChars = 0;
    let paraCount = 0;

    for (const part of parts) {
      const doc = parseXml(await zip.file(part.path).async('string'));
      docs.set(part.path, { doc, part, changed: false });
      for (const p of collectParagraphs(doc, part.nonPh)) {
        const segs = paragraphSegments(p);
        if (!segs.some((s) => !isFixedSegment(s))) continue;
        const texts = segs.map((s) => s.text);
        const key = JSON.stringify(texts);
        paraCount++;
        if (!unique.has(key)) {
          unique.set(key, {
            id: String(unique.size + 1),
            texts,
            fixed: segs.map(isFixedSegment),
            occurrences: [],
            slide: part.slide,
          });
          totalChars += texts.join('').length;
        }
        unique.get(key).occurrences.push({ segs, path: part.path });
      }
    }

    if (!unique.size) throw new Error(t('noText'));
    if (totalChars > MAX_TOTAL_CHARS) {
      throw new Error(t('tooMuchText', { n: totalChars.toLocaleString('nl-BE'), max: MAX_TOTAL_CHARS.toLocaleString('nl-BE') }));
    }

    // --- 2. In delen naar de Worker sturen ---
    const items = [...unique.values()].sort((a, b) => (a.slide ?? 9999) - (b.slide ?? 9999));
    const batches = makeBatches(items);
    setStatus(status, t('batchesDone', { done: 0, total: batches.length }));
    const results = new Map(); // id -> segs
    const usage = { inputTokens: 0, outputTokens: 0 };
    const glossaryTerms = new Map(); // src -> tgt
    let done = 0;

    let failedBatches = 0;
    let lastBatchError = null;
    const runBatch = async (batch) => {
      let r;
      try {
        r = await translateBatch(batch, src, tgt);
      } catch (err) {
        if (err.authFailure) throw err;
        // Deze alinea's komen hieronder nog eens aan bod (tweede kans).
        failedBatches++;
        lastBatchError = err;
        done++;
        setProgress(done, batches.length);
        return;
      }
      for (const t of r.translations) results.set(String(t.id), t.segs);
      usage.inputTokens += r.usage.inputTokens || 0;
      usage.outputTokens += r.usage.outputTokens || 0;
      for (const g of r.glossaryTerms || []) glossaryTerms.set(g.src, g.tgt);
      done++;
      setProgress(done, batches.length);
      setStatus(status, t('batchesDone', { done, total: batches.length }));
    };
    await runPool(batches, runBatch, PARALLEL_CALLS);
    if (failedBatches === batches.length) throw lastBatchError || new Error(t('translateFailed'));

    // Items met een ontbrekende of foute vertaling (verkeerd aantal
    // segmenten) krijgen één tweede kans, in kleine delen.
    const retry = items.filter((it) => !validTranslation(it, results.get(it.id)));
    if (retry.length) {
      setStatus(status, t('retrying', { n: retry.length }));
      const rb = makeBatches(retry, 15);
      await runPool(rb, async (batch) => {
        try {
          const r = await translateBatch(batch, src, tgt);
          for (const t of r.translations) {
            const it = batch.find((b) => b.id === String(t.id));
            if (it && validTranslation(it, t.segs)) results.set(it.id, t.segs);
            else if (it && !results.has(it.id)) results.set(it.id, t.segs);
          }
          usage.inputTokens += r.usage.inputTokens || 0;
          usage.outputTokens += r.usage.outputTokens || 0;
        } catch (err) {
          if (err.authFailure) throw err;
          // tweede poging mislukt: deze alinea's blijven onvertaald (zie rapport)
        }
      }, PARALLEL_CALLS);
    }

    // --- 3. Terugzetten ---
    setStatus(status, t('writingBack'));
    const report = { simplified: [], untranslated: [], longer: [], symbols: [] };
    const tgtTag = LANGS[tgt].tag;

    for (const it of items) {
      const out = finalizeSegments(it, results.get(it.id), report);
      if (!out) continue;
      for (const occ of it.occurrences) {
        occ.segs.forEach((seg, i) => {
          if (seg.kind === 'r' && seg.tEl && seg.tEl.textContent !== out[i]) {
            seg.tEl.textContent = out[i];
            docs.get(occ.path).changed = true;
          }
        });
      }
      const before = it.texts.join('');
      const after = out.join('');
      if (after.length > before.length * 1.15 && after.length - before.length >= 8) {
        report.longer.push({ slide: it.slide, before, after });
      }
    }

    // Taalcode van de tekst aanpassen (spellingcontrole in PowerPoint).
    for (const { doc, part } of docs.values()) {
      if (part.kind === 'layout' || part.kind === 'master') continue;
      if (setLanguage(doc, src, tgtTag)) docs.get(part.path).changed = true;
    }

    // --- 4. Opnieuw inpakken (zelfde volgorde, enkel gewijzigde onderdelen) ---
    setStatus(status, t('packing'));
    for (const [path, entry] of docs) {
      if (entry.changed) zip.file(path, serializeXml(entry.doc), { createFolders: false });
    }
    resultBlob = await zip.generateAsync({
      type: 'blob',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    });
    resultName = outputName(currentFile.name, src, tgt);

    // --- 5. Rapport ---
    const glossaryCheck = checkGlossary(items, results, glossaryTerms, report);
    lastResult = { fileName: currentFile.name, slideCount: slides.length, paraCount, usage, report, glossaryCheck, src, tgt, includeNotes };
    renderResult(lastResult, true);
    setStatus(status, t('done'), 'ok');
  } catch (err) {
    if (!err.authFailure) setStatus(status, err.message || String(err), 'err');
    $('progressBar').hidden = true;
  } finally {
    busy = false;
    updateTranslateBtn();
    if (resultBlob) setStatus(status, t('done'), 'ok');
  }
}

function makeBatches(items, maxItems = BATCH_MAX_ITEMS) {
  const batches = [];
  let cur = [];
  let chars = 0;
  for (const it of items) {
    const len = it.texts.join('').length;
    if (cur.length && (cur.length >= maxItems || chars + len > BATCH_MAX_CHARS)) {
      batches.push(cur);
      cur = [];
      chars = 0;
    }
    cur.push(it);
    chars += len;
  }
  if (cur.length) batches.push(cur);
  return batches;
}

async function runPool(list, fn, size) {
  let i = 0;
  let failure = null;
  const workers = Array.from({ length: Math.min(size, list.length) }, async () => {
    while (i < list.length && !failure) {
      const item = list[i++];
      try {
        await fn(item);
      } catch (err) {
        failure = err;
      }
    }
  });
  await Promise.all(workers);
  if (failure) throw failure;
}

async function translateBatch(batch, src, tgt) {
  const payload = {
    mode: 'translate',
    source: src,
    target: tgt,
    items: batch.map((it) => ({ id: it.id, slide: it.slide ?? '', segs: it.texts })),
  };
  let lastErr = null;
  for (let attempt = 0; attempt < 4; attempt++) {
    if (attempt) await new Promise((r) => setTimeout(r, 1500 * 2 ** attempt));
    let res;
    try {
      res = await apiRequest(payload);
    } catch {
      lastErr = new Error(t('noConnection'));
      continue;
    }
    if (handleAuthFailure(res.status)) {
      const e = new Error(t('sessionExpired'));
      e.authFailure = true;
      throw e;
    }
    if (res.ok) {
      return {
        translations: Array.isArray(res.data.translations) ? res.data.translations : [],
        usage: res.data.usage || {},
        glossaryTerms: res.data.glossaryTerms || [],
      };
    }
    lastErr = new Error(res.data.error || t('translateFailedCode', { code: res.status }));
    if (!res.data.retryable) break;
  }
  throw lastErr;
}

function validTranslation(it, segs) {
  return Array.isArray(segs) && segs.length === it.texts.length;
}

// Zet het antwoord van de AI om naar de definitieve tekst per segment, met
// de vangnetten die de opmaak beschermen.
function finalizeSegments(it, segs, report) {
  const orig = it.texts;
  if (!Array.isArray(segs) || !segs.some((s) => String(s || '').trim())) {
    report.untranslated.push({ slide: it.slide, before: orig.join('') });
    return null;
  }
  let out;
  if (segs.length === orig.length) {
    out = orig.map((o, i) => (it.fixed[i] ? o : keepEdges(o, String(segs[i] ?? ''))));
  } else {
    // Ander aantal segmenten: volledige vertaling in het eerste
    // tekstsegment, de andere tekstsegmenten leeg. Tekst klopt, enkel de
    // opmaak binnen de alinea (bv. één vet woord) gaat verloren.
    const textual = segs.map((s) => String(s ?? '')).join(' ').replace(PUA_RE, '');
    const first = it.fixed.indexOf(false);
    out = orig.map((o, i) => (it.fixed[i] ? o : i === first ? keepEdges(o, textual.replace(/\s+/g, ' ').trim()) : ''));
    report.simplified.push({ slide: it.slide, before: orig.join(''), after: out.join('') });
  }
  // Symbool-tekens (pijltjes, vinkjes uit Wingdings) moeten behouden blijven.
  out = out.map((s, i) => {
    const need = (orig[i].match(PUA_RE) || []).join('');
    const have = (s.match(PUA_RE) || []).join('');
    if (need && need !== have) {
      report.symbols.push({ slide: it.slide, before: orig.join('') });
      return orig[i];
    }
    return s;
  });
  return out;
}

// Behoud de spaties/tabs vooraan en achteraan het oorspronkelijke segment.
function keepEdges(orig, translated) {
  const lead = orig.match(/^[\s ]*/)[0];
  const trail = orig.match(/[\s ]*$/)[0];
  const core = translated.replace(/^[\s ]+|[\s ]+$/g, '');
  if (!core) return '';
  return lead + core + trail;
}

function setLanguage(doc, src, tgtTag) {
  let changed = false;
  for (const name of ['rPr', 'endParaRPr', 'defRPr']) {
    for (const el of doc.getElementsByTagNameNS(NS_A, name)) {
      const lang = el.getAttribute('lang');
      if (lang && lang.toLowerCase().split('-')[0] === src && lang !== tgtTag) {
        el.setAttribute('lang', tgtTag);
        changed = true;
      }
    }
  }
  return changed;
}

function outputName(name, src, tgt) {
  const base = name.replace(/\.pptx$/i, '');
  const re = new RegExp(`([_\\- ])${LANGS[src].code}$`, 'i');
  const out = re.test(base) ? base.replace(re, `$1${LANGS[tgt].code}`) : `${base}_${LANGS[tgt].code}`;
  return out + '.pptx';
}

// Controleert of elke woordenlijst-term die in de brontekst voorkomt, ook in
// de vertaling terechtkwam.
function checkGlossary(items, results, terms, report) {
  const applied = [];
  const missing = [];
  for (const [src, tgt] of terms) {
    const s = src.toLowerCase();
    const sStem = s.length > 5 ? s.replace(/(en|es|e|s|n)$/u, '') : s;
    const t = tgt.toLowerCase();
    const tStem = t.length > 5 ? t.replace(/(en|es|e|s|n)$/u, '') : t;
    let found = 0;
    let ok = 0;
    for (const it of items) {
      const before = it.texts.join('').toLowerCase();
      if (!before.includes(s) && !before.includes(sStem)) continue;
      const out = results.get(it.id);
      if (!Array.isArray(out)) continue;
      found++;
      const after = out.join('').toLowerCase();
      if (after.includes(t) || after.includes(tStem)) ok++;
      else missing.push({ slide: it.slide, src, tgt, after: out.join('') });
    }
    if (found) applied.push({ src, tgt, found, ok });
  }
  report.glossaryMissing = missing;
  return applied;
}

function renderResult(r, scroll) {
  const card = $('resultCard');
  $('resultSub').textContent = t('resultSub', { file: r.fileName, src: langLabel(r.src), tgt: langLabel(r.tgt) });
  $('resultStats').innerHTML = [
    [r.slideCount, t('statSlides')],
    [r.paraCount, t('statParas')],
    [r.glossaryCheck.length, t('statTerms')],
    [Math.round((r.usage.inputTokens + r.usage.outputTokens) / 1000) + 'K', t('statTokens')],
  ].map(([n, l]) => `<div class="stat"><b>${esc(n)}</b><span>${esc(l)}</span></div>`).join('');

  const rep = r.report;
  // Symbool-tekens (Wingdings-pijltjes) zijn buiten PowerPoint onleesbaar.
  const show = (txt) => esc(String(txt).replace(PUA_RE, '').replace(/\s+/g, ' ').trim());
  const slideLbl = (n) => esc(n ? t('slideN', { n }) : t('layoutLbl'));
  const title = (key) => `<h3 class="part-subtitle">${esc(t(key))}</h3>`;
  const sub = (key) => `<p class="part-sub">${esc(t(key))}</p>`;
  const list = (rows, render) => (rows.length
    ? `<ul class="report-list">${rows.map(render).join('')}</ul>`
    : `<p class="report-empty">${esc(t('none'))}</p>`);
  let html = '';

  html += title('repLongerTitle') + sub('repLongerSub');
  html += list(rep.longer, (x) => `<li><span class="ranked-cat">${slideLbl(x.slide)}</span><span class="txt">${show(x.after)}<div class="orig">${show(x.before)}</div></span><span class="pill issue">+${Math.round((x.after.length / x.before.length - 1) * 100)}%</span></li>`);

  if (rep.glossaryMissing && rep.glossaryMissing.length) {
    html += title('repGlossTitle');
    html += list(rep.glossaryMissing, (x) => `<li><span class="ranked-cat">${slideLbl(x.slide)}</span><span class="txt">${show(x.after)}<div class="orig">${esc(t('repExpected', { src: x.src, tgt: x.tgt }))}</div></span><span class="pill issue">${esc(t('pillTerm'))}</span></li>`);
  }

  if (rep.simplified.length) {
    html += title('repSimplTitle') + sub('repSimplSub');
    html += list(rep.simplified, (x) => `<li><span class="ranked-cat">${slideLbl(x.slide)}</span><span class="txt">${show(x.after)}<div class="orig">${show(x.before)}</div></span></li>`);
  }

  const notDone = rep.untranslated.concat(rep.symbols);
  if (notDone.length) {
    html += title('repNotTitle') + sub('repNotSub');
    html += list(notDone, (x) => `<li><span class="ranked-cat">${slideLbl(x.slide)}</span><span class="txt">${show(x.before)}</span><span class="pill negative">${esc(t('pillOriginal'))}</span></li>`);
  }

  if (r.glossaryCheck.length) {
    html += title('repTermsTitle');
    html += '<div class="term-chips">' + r.glossaryCheck
      .map((g) => `<span class="term-chip">${esc(g.src)} → ${esc(g.tgt)}</span>`).join('') + '</div>';
  }

  html += title('repSkippedTitle') + sub(r.includeNotes ? 'repSkippedNotes' : 'repSkippedNoNotes');

  $('reportBody').innerHTML = html;
  card.style.display = 'block';
  if (scroll) card.scrollIntoView({ behavior: 'smooth', block: 'start' });
}


// ============================================================
// Admin: woordenlijst beheren
// ============================================================
let parsedGlossary = null;
wireDropzone('glossaryDropzone', 'glossaryFileInput', handleGlossaryFile);

function updateGlossaryCurrentText() {
  const el = $('glossaryCurrent');
  if (!glossaryInfo || !glossaryInfo.available) {
    setStatus(el, t('glossaryNoneAdmin'), 'err');
    return;
  }
  const langs = (glossaryInfo.languages || []).map(langLabel).join(', ');
  setStatus(el, glossaryInfo.fileName && glossaryInfo.updatedAt
    ? t('glossaryCurrentFile', { n: glossaryInfo.terms, langs, file: glossaryInfo.fileName, date: formatDate(glossaryInfo.updatedAt) })
    : t('glossaryCurrent', { n: glossaryInfo.terms, langs }));
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        resolve(XLSX.read(e.target.result, { type: 'array' }));
      } catch {
        reject(new Error(t('excelReadFail')));
      }
    };
    reader.onerror = () => reject(new Error(t('fileReadFail')));
    reader.readAsArrayBuffer(file);
  });
}

async function handleGlossaryFile(file) {
  const st = $('glossaryStatus');
  setFileLabel('glossaryFname', file.name);
  $('glossaryUpdateBtn').disabled = true;
  parsedGlossary = null;
  setStatus(st, t('reading'));
  let wb;
  try {
    wb = await readWorkbook(file);
  } catch (err) {
    setStatus(st, err.message, 'err');
    return;
  }
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
  // Koprij zoeken: eerste rij met minstens twee herkende taalnamen.
  let headerIdx = -1;
  let colMap = {};
  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    const map = {};
    rows[i].forEach((cell, c) => {
      const lang = HEADER_MAP[String(cell).trim().toLowerCase()];
      if (lang && !(lang in map)) map[lang] = c;
    });
    if (Object.keys(map).length >= 2) {
      headerIdx = i;
      colMap = map;
      break;
    }
  }
  if (headerIdx < 0) {
    setStatus(st, t('noHeader'), 'err');
    return;
  }
  const languages = Object.keys(LANGS).filter((l) => l in colMap);
  const clean = (v) => String(v ?? '').replace(/\s+/g, ' ').trim();
  const data = rows.slice(headerIdx + 1)
    .map((r) => Object.fromEntries(languages.map((l) => [l, clean(r[colMap[l]])])))
    .filter((o) => languages.filter((l) => o[l]).length >= 2);
  if (!data.length) {
    setStatus(st, t('noRows'), 'err');
    return;
  }
  // Waarschuwingen: dubbele termen met verschillende vertaling, lege cellen.
  const warnings = [];
  const byKey = new Map();
  for (const o of data) {
    const k = (o[languages[0]] || '').toLowerCase();
    if (!k) continue;
    if (byKey.has(k)) {
      const prev = byKey.get(k);
      const diff = languages.filter((l) => prev[l] && o[l] && prev[l] !== o[l]);
      if (diff.length) warnings.push(t('dupWarn', { term: o[languages[0]], langs: diff.map(langLabel).join(', ') }));
    } else {
      byKey.set(k, o);
    }
  }
  const empty = data.reduce((n, o) => n + languages.filter((l) => !o[l]).length, 0);
  if (empty) warnings.push(t('emptyWarn', { n: empty }));
  parsedGlossary = { languages, rows: data, fileName: file.name };
  setStatus(st, [
    t('glossaryParsed', { n: data.length, langs: languages.map(langLabel).join(', ') }),
    warnings.length ? t('warnPrefix', { list: warnings.join('; ') }) : '',
    t('confirmClick', { btn: t('glossaryUpdateBtn') }),
  ].filter(Boolean).join(' '), warnings.length ? 'err' : '');
  $('glossaryUpdateBtn').disabled = false;
}

$('glossaryUpdateBtn').addEventListener('click', async () => {
  if (!parsedGlossary) return;
  const st = $('glossaryStatus');
  $('glossaryUpdateBtn').disabled = true;
  setStatus(st, t('updating'));
  try {
    const { ok, status, data } = await apiRequest({ mode: 'glossary_put', ...parsedGlossary });
    if (handleAuthFailure(status)) return;
    if (!ok) {
      setStatus(st, data.error || t('updateFailed'), 'err');
      $('glossaryUpdateBtn').disabled = false;
      return;
    }
    setStatus(st, t('termsSaved', { n: data.bewaard }), 'ok');
    parsedGlossary = null;
    await loadGlossaryInfo();
  } catch (err) {
    setStatus(st, t('updateFailedMsg', { msg: err.message }), 'err');
    $('glossaryUpdateBtn').disabled = false;
  }
});

$('glossaryExportBtn').addEventListener('click', async () => {
  const st = $('glossaryStatus');
  try {
    const { ok, status, data } = await apiRequest({ mode: 'glossary_get' });
    if (handleAuthFailure(status)) return;
    if (!ok) throw new Error(data.error || t('unknownError'));
    const g = data.glossary;
    if (!g || !g.rows || !g.rows.length) {
      setStatus(st, t('noGlossaryYet'), 'err');
      return;
    }
    const aoa = [g.languages.map((l) => GLOSSARY_HEADERS[l] || l)]
      .concat(g.rows.map((r) => g.languages.map((l) => r[l] || '')));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = g.languages.map(() => ({ wch: 28 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Woordenlijst');
    XLSX.writeFile(wb, 'Woordenlijst_huidig.xlsx');
  } catch (err) {
    setStatus(st, t('downloadFailed', { msg: err.message }), 'err');
  }
});

// ============================================================
// Admin: gebruikers beheren — overgenomen van Feedbackloop
// ============================================================
let parsedUserRows = [];
wireDropzone('usersDropzone', 'usersFileInput', handleUsersFile);

async function handleUsersFile(file) {
  const st = $('usersStatus');
  setFileLabel('usersFname', file.name);
  setStatus(st, t('reading'));
  $('usersUpdateBtn').disabled = true;
  parsedUserRows = [];
  let wb;
  try {
    wb = await readWorkbook(file);
  } catch (err) {
    setStatus(st, err.message, 'err');
    return;
  }
  const sheetName = wb.SheetNames[0];
  const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '', raw: false });
  const rows = raw
    .map((r) => {
      const n = {};
      for (const k in r) n[k.trim().toLowerCase()] = String(r[k] ?? '').trim();
      return { username: n.user || '', password: n.pw || '', role: (n.role || '').toLowerCase() };
    })
    .filter((r) => r.username || r.password);
  if (!rows.length) {
    setStatus(st, t('usersNoRows'), 'err');
    return;
  }
  parsedUserRows = rows;
  setStatus(st, t('usersParsed', { n: rows.length, sheet: sheetName }) + ' ' + t('confirmClick', { btn: t('usersUpdateBtn') }));
  $('usersUpdateBtn').disabled = false;
}

$('usersUpdateBtn').addEventListener('click', async () => {
  if (!parsedUserRows.length) return;
  const st = $('usersStatus');
  $('usersUpdateBtn').disabled = true;
  setStatus(st, t('updating'));
  try {
    const { ok, status, data } = await apiRequest({ mode: 'admin_upsert_users', users: parsedUserRows });
    if (handleAuthFailure(status)) return;
    if (!ok) {
      setStatus(st, (data && data.error) || t('updateFailed'), 'err');
      $('usersUpdateBtn').disabled = false;
      return;
    }
    let msg = t('usersUpdated', { u: data.updated, t: data.total });
    if (data.errors && data.errors.length) msg += ' ' + t('errorsPrefix', { list: data.errors.join('; ') });
    setStatus(st, msg, data.errors && data.errors.length ? 'err' : 'ok');
    parsedUserRows = [];
    if (!(data.errors && data.errors.length)) {
    }
  } catch (err) {
    setStatus(st, t('updateFailedMsg', { msg: err.message }), 'err');
    $('usersUpdateBtn').disabled = false;
  }
});

// ============================================================
// Taal van de app (zoals Winsol-Socrates)
// ============================================================
function refreshUiLanguage() {
  applyI18n();
  fillUiLangSelect($('uiLangLogin'));
  fillUiLangSelect($('uiLangApp'));
  fillLangSelects();
  $('loginError').textContent = '';
  if (authRole) {
    $('topbarRole').textContent = authRole === 'admin' ? t('roleAdmin') : t('roleUser');
    if (glossaryInfo) updateGlossaryInfoText();
    if (authRole === 'admin' && glossaryInfo) updateGlossaryCurrentText();
  }
  if (lastResult) {
    renderResult(lastResult, false);
  }
  if (currentFile && !busy && !resultBlob) {
    const fileStatus = $('fileStatus');
    if (!fileStatus.classList.contains('err')) setStatus(fileStatus, t('fileReady', { size: (currentFile.size / 1048576).toFixed(1) }));
  }
}

for (const id of ['uiLangLogin', 'uiLangApp']) {
  $(id).addEventListener('change', (e) => {
    const v = e.target.value;
    uiLangOverride = v;
    saveUiLangSetting(v);
    refreshUiLanguage();
  });
}

refreshUiLanguage();
tryRestoreSession();
