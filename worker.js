// ============================================================
// Winsol Presentatievertaler — Cloudflare Worker
// ============================================================
//
// Zelfde patroon als Feedbackloop: de frontend (GitHub Pages) doet al het
// werk met het PPTX-bestand zelf (uitpakken, tekst eruit halen, terugzetten,
// opnieuw inpakken) in de browser. Deze Worker krijgt ENKEL losse
// dia-teksten binnen, voegt de woordenlijst toe en vraagt Claude om ze te
// vertalen.
//
// Er wordt GEEN presentatie-inhoud bewaard: geen KV-writes met tekst, geen
// logging van tekst. In de KV staan enkel:
//   user:<username>    -> { username, passwordHash, role }
//   session:<token>    -> { username, role }            (met vervaldatum)
//   glossary           -> { languages, rows, fileName, updatedAt, updatedBy }
//
// Secrets (npx wrangler secret put ...):
//   ANTHROPIC_API_KEY  — API-sleutel van console.anthropic.com
//   PASSWORD_PEPPER    — willekeurige lange tekst, voor de wachtwoord-hashes
// Optionele variabele (wrangler.toml [vars]):
//   CLAUDE_MODEL       — standaard claude-sonnet-4-6

// Enkel de eigen webpagina mag de Worker vanuit een browser aanspreken.
// Ander adres (andere repo, eigen domein, lokaal testen)? Zet de variabele
// ALLOWED_ORIGIN in wrangler.toml [vars] of in .dev.vars (bv. "*" om lokaal te testen).
const DEFAULT_ALLOWED_ORIGIN = 'https://winsol-belgie.github.io';
let ALLOWED_ORIGIN = DEFAULT_ALLOWED_ORIGIN;
const DEFAULT_MODEL = 'claude-sonnet-4-6';

// Grenzen per aanroep (de frontend splitst grotere presentaties zelf op in
// meerdere aanroepen). Houdt elke prompt klein en snel, en begrenst de kost.
const MAX_ITEMS_PER_CALL = 120;
const MAX_CHARS_PER_CALL = 12000;
const MAX_GLOSSARY_TERMS_PER_CALL = 250;

const GLOSSARY_KEY = 'glossary';

// Ondersteunde talen. "variant" wordt letterlijk aan Claude gegeven.
const LANGS = {
  nl: { name: 'Dutch', variant: 'Dutch as written in Flanders (Belgium)' },
  fr: { name: 'French', variant: 'French as written in Belgium (standard French spelling)' },
  de: { name: 'German', variant: 'standard German' },
  en: { name: 'English', variant: 'British English (UK spelling: -ise, colour, programme, etc.)' },
  es: { name: 'Spanish', variant: 'Spanish as written in Spain' },
  it: { name: 'Italian', variant: 'standard Italian as written in Italy' },
};

const TRANSLATE_TOOL = {
  name: 'return_translations',
  description: 'Geef de vertaling van elk item terug, met exact evenveel segmenten als het origineel.',
  input_schema: {
    type: 'object',
    properties: {
      translations: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            id: { type: 'string' },
            segs: { type: 'array', items: { type: 'string' } },
          },
          required: ['id', 'segs'],
        },
      },
    },
    required: ['translations'],
  },
};

function systemPrompt() {
  return [
    'You translate the text of a PowerPoint presentation for Winsol, a Belgian manufacturer of sun protection (screens, roller shutters, awnings, pergolas, windows and doors).',
    '',
    'INPUT: a list of items in slide order. Each item is ONE paragraph, split into "segs" (text runs). Every run carries its own formatting (bold, colour, underline, hyperlink...).',
    '',
    'OUTPUT RULES — follow strictly:',
    '1. Return every item id exactly once, through the return_translations tool.',
    '2. Return EXACTLY the same number of segs as the input item. Segment i of the output keeps the formatting of segment i of the input, so put each translated word in the segment that plays the same role (e.g. an underlined word stays in the underlined segment). A segment may become an empty string only when that is unavoidable.',
    '3. Copy unchanged: segments that contain only whitespace, tabs or "\\n"; placeholders of the form ⟦S1⟧, ⟦S2⟧ (they stand for symbols such as arrows and ticks: keep every placeholder exactly once, at the same position relative to the text, e.g. "⟦S1⟧ Wordt bekeken" -> "⟦S1⟧ Being reviewed"); leading and trailing spaces and tabs of every segment; numbers, dates, units, percentages, currency amounts; e-mail addresses and URLs; product names, brand names, model codes and abbreviations (e.g. SO!, Z!P, WinCube, SolarFix, Somfy, QC, R&D, SPOC).',
    '4. Text that is already in the target language, or that is a name/code, is returned unchanged.',
    '5. The GLOSSARY is binding: when a glossary source term occurs (also inside a compound word, a plural or another inflection), use the glossary target term, adapted to the grammar (plural, compound). Do not translate glossary terms any other way.',
    '6. The slide layout cannot change. Keep each translation as short as the original or shorter; never exceed the original length by more than 10% when a shorter correct phrasing exists. Prefer concise, natural presentation style (titles stay titles, bullets stay bullets).',
    '7. Use the other items for context (same slide = same topic). Keep terminology consistent across the whole presentation.',
    '8. Do not add explanations, quotes or notes. Only the translated text.',
  ].join('\n');
}

export default {
  async fetch(request, env) {
    ALLOWED_ORIGIN = env.ALLOWED_ORIGIN || DEFAULT_ALLOWED_ORIGIN;
    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: corsHeaders() });
    }
    if (request.method !== 'POST') {
      return jsonResponse({ error: 'Method not allowed' }, 405);
    }

    let body;
    try {
      body = await request.json();
    } catch {
      return jsonResponse({ error: 'Ongeldige request body.' }, 400);
    }

    // --- Auth-routes: geen AI-aanroep. ---
    if (body.mode === 'login') return handleLogin(body, env, request);
    if (body.mode === 'logout') return handleLogout(request, env);
    if (body.mode === 'auth_me') {
      const session = await getSession(request, env);
      return session
        ? jsonResponse({ username: session.username, role: session.role })
        : jsonResponse({ error: 'Niet ingelogd of sessie verlopen.' }, 401);
    }

    // --- Alles hieronder vereist een geldige, ingelogde sessie. ---
    const session = await getSession(request, env);
    if (!session) {
      return jsonResponse({ error: 'Niet ingelogd of sessie verlopen.' }, 401);
    }

    if (body.mode === 'glossary_info') {
      const g = await env.VERTALER_KV.get(GLOSSARY_KEY, 'json');
      return jsonResponse({
        available: !!(g && Array.isArray(g.rows) && g.rows.length),
        terms: g && Array.isArray(g.rows) ? g.rows.length : 0,
        languages: g ? g.languages : [],
        fileName: g ? g.fileName : '',
        updatedAt: g ? g.updatedAt : '',
      });
    }

    if (body.mode === 'translate') {
      return handleTranslate(body, env);
    }

    // --- Admin-only ---
    if (session.role !== 'admin') {
      return jsonResponse({ error: 'Enkel toegankelijk voor admins.' }, 403);
    }

    if (body.mode === 'glossary_get') {
      const g = await env.VERTALER_KV.get(GLOSSARY_KEY, 'json');
      return jsonResponse({ glossary: g || null });
    }
    if (body.mode === 'glossary_put') {
      return handleGlossaryPut(body, env, session);
    }
    if (body.mode === 'admin_upsert_users') {
      return handleAdminUpsertUsers(body, env, session);
    }

    return jsonResponse({ error: 'Onbekende actie.' }, 400);
  },
};

// ============================================================
// Vertalen
// ============================================================

async function handleTranslate(body, env) {
  if (!env.ANTHROPIC_API_KEY) {
    return jsonResponse({ error: 'Geen API-sleutel geconfigureerd (wrangler secret put ANTHROPIC_API_KEY).' }, 500);
  }
  const source = String(body.source || '');
  const target = String(body.target || '');
  if (!LANGS[source] || !LANGS[target]) {
    return jsonResponse({ error: 'Onbekende bron- of doeltaal.' }, 400);
  }
  if (source === target) {
    return jsonResponse({ error: 'Bron- en doeltaal zijn dezelfde.' }, 400);
  }

  const items = Array.isArray(body.items) ? body.items : [];
  if (!items.length) return jsonResponse({ error: 'Geen tekst ontvangen.' }, 400);
  if (items.length > MAX_ITEMS_PER_CALL) {
    return jsonResponse({ error: `Te veel alinea's in één aanroep (max ${MAX_ITEMS_PER_CALL}).` }, 413);
  }
  const clean = [];
  let chars = 0;
  for (const it of items) {
    const segs = Array.isArray(it.segs) ? it.segs.map((s) => String(s ?? '')) : [];
    const id = String(it.id ?? '');
    if (!id || !segs.length) continue;
    chars += segs.join('').length;
    clean.push({ id, slide: it.slide ?? '', segs });
  }
  if (chars > MAX_CHARS_PER_CALL) {
    return jsonResponse({ error: `Te veel tekst in één aanroep (max ${MAX_CHARS_PER_CALL} tekens).` }, 413);
  }

  // Enkel de woordenlijst-termen meesturen die in deze teksten voorkomen.
  const glossary = await env.VERTALER_KV.get(GLOSSARY_KEY, 'json');
  const terms = relevantTerms(glossary, source, target, clean);

  const userContent = [
    `Translate from ${LANGS[source].name} into ${LANGS[target].variant}.`,
    '',
    'GLOSSARY (source => target, binding):',
    terms.length ? terms.map((t) => `${t.src} => ${t.tgt}`).join('\n') : '(no glossary terms apply to these items)',
    '',
    'ITEMS (JSON):',
    JSON.stringify(clean),
  ].join('\n');

  const model = env.CLAUDE_MODEL || DEFAULT_MODEL;
  let apiRes;
  try {
    apiRes = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'x-api-key': env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model,
        max_tokens: 16000,
        temperature: 0,
        system: [{ type: 'text', text: systemPrompt(), cache_control: { type: 'ephemeral' } }],
        tools: [TRANSLATE_TOOL],
        tool_choice: { type: 'tool', name: TRANSLATE_TOOL.name },
        messages: [{ role: 'user', content: userContent }],
      }),
    });
  } catch (err) {
    return jsonResponse({ error: 'AI niet bereikbaar: ' + err.message, retryable: true }, 502);
  }

  if (!apiRes.ok) {
    const retryable = apiRes.status === 429 || apiRes.status === 529 || apiRes.status >= 500;
    let msg = `AI-fout (${apiRes.status}).`;
    try {
      const e = await apiRes.json();
      if (e && e.error && e.error.message) msg += ' ' + e.error.message;
    } catch {
      // geen JSON
    }
    return jsonResponse({ error: msg, retryable }, retryable ? 503 : 502);
  }

  const data = await apiRes.json();
  const toolUse = (data.content || []).find((c) => c.type === 'tool_use');
  const translations = toolUse && toolUse.input && Array.isArray(toolUse.input.translations)
    ? toolUse.input.translations
    : [];

  return jsonResponse({
    translations,
    truncated: data.stop_reason === 'max_tokens',
    glossaryTerms: terms,
    usage: {
      inputTokens: data.usage?.input_tokens || 0,
      outputTokens: data.usage?.output_tokens || 0,
      cacheReadTokens: data.usage?.cache_read_input_tokens || 0,
    },
  });
}

// Zoekt per woordenlijst-rij of de bronterm (of een stam ervan, voor
// meervouden en samenstellingen) in de teksten voorkomt.
function relevantTerms(glossary, source, target, items) {
  if (!glossary || !Array.isArray(glossary.rows)) return [];
  const hay = items.map((i) => i.segs.join('')).join('\n').toLowerCase();
  const out = [];
  const seen = new Set();
  for (const row of glossary.rows) {
    const src = String(row[source] || '').trim();
    const tgt = String(row[target] || '').trim();
    if (!src || !tgt) continue;
    const key = src.toLowerCase();
    if (seen.has(key + '=>' + tgt.toLowerCase())) continue;
    if (hay.includes(key) || (key.length > 5 && hay.includes(stem(key)))) {
      seen.add(key + '=>' + tgt.toLowerCase());
      out.push({ src, tgt });
      if (out.length >= MAX_GLOSSARY_TERMS_PER_CALL) break;
    }
  }
  return out;
}

function stem(word) {
  // Eenvoudige stam: laatste meervouds-/buigingsuitgang eraf.
  return word.replace(/(en|es|e|s|n)$/u, '');
}

// ============================================================
// Woordenlijst (admin)
// ============================================================
// De frontend leest de Excel al in de browser in en stuurt
// { languages: ['nl','fr',...], rows: [{nl:'...', fr:'...'}, ...], fileName }.
// De volledige lijst wordt vervangen (geen samenvoeging).

async function handleGlossaryPut(body, env, session) {
  const languages = (Array.isArray(body.languages) ? body.languages : []).filter((l) => LANGS[l]);
  const rows = Array.isArray(body.rows) ? body.rows : [];
  if (languages.length < 2) {
    return jsonResponse({ error: 'Minstens twee taalkolommen nodig (bv. Nederlands en Engels).' }, 400);
  }
  const cleanRows = rows
    .map((r) => {
      const o = {};
      for (const l of languages) o[l] = String((r && r[l]) || '').replace(/\s+/g, ' ').trim();
      return o;
    })
    .filter((o) => languages.filter((l) => o[l]).length >= 2);
  if (!cleanRows.length) {
    return jsonResponse({ error: 'Geen bruikbare rijen (elke rij heeft minstens twee talen nodig).' }, 400);
  }
  const record = {
    languages,
    rows: cleanRows,
    fileName: String(body.fileName || '').slice(0, 200),
    updatedAt: new Date().toISOString(),
    updatedBy: session.username,
  };
  await env.VERTALER_KV.put(GLOSSARY_KEY, JSON.stringify(record));
  return jsonResponse({ bewaard: cleanRows.length, languages });
}

// ============================================================
// Auth — overgenomen van Feedbackloop (users en sessies in Workers KV).
// ============================================================
// Wachtwoorden worden nooit leesbaar bewaard: sha256(wachtwoord + PASSWORD_PEPPER).

const SESSION_TTL_REMEMBER = 60 * 60 * 24 * 30; // 30 dagen ("blijf aangemeld")
const SESSION_TTL_DEFAULT = 60 * 60 * 12; // 12 uur

async function sha256Hex(str) {
  const data = new TextEncoder().encode(str);
  const hashBuffer = await crypto.subtle.digest('SHA-256', data);
  return [...new Uint8Array(hashBuffer)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

function randomToken() {
  const bytes = crypto.getRandomValues(new Uint8Array(32));
  return [...bytes].map((b) => b.toString(16).padStart(2, '0')).join('');
}

// Beveiliging van de login:
// - wachtwoorden enkel als hash (sha256 + geheime PASSWORD_PEPPER);
// - na 5 mislukte pogingen voor dezelfde gebruikersnaam, of 20 vanaf hetzelfde
//   IP-adres, wordt aanmelden 15 minuten geblokkeerd;
// - een sessie is enkel geldig zolang de gebruiker nog bestaat; rol-wijzigingen
//   en verwijderingen gelden dus meteen, ook voor wie al aangemeld is;
// - wachtwoorden zijn minstens 8 tekens.
const MAX_FAILS_USER = 5;
const MAX_FAILS_IP = 20;
const FAIL_WINDOW = 60 * 15; // seconden
const MIN_PASSWORD_LENGTH = 8;

function timingSafeEqual(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string' || a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

async function getSession(request, env) {
  const token = request.headers.get('X-Auth-Token') || '';
  if (!token) return null;
  const raw = await env.VERTALER_KV.get(`session:${token}`);
  if (!raw) return null;
  let session;
  try {
    session = JSON.parse(raw);
  } catch {
    return null;
  }
  // Bestaat de gebruiker nog? Rol altijd uit het actuele gebruikersrecord.
  const user = await env.VERTALER_KV.get(`user:${session.username}`, 'json');
  if (!user) {
    await env.VERTALER_KV.delete(`session:${token}`);
    return null;
  }
  return { username: user.username, role: user.role };
}

async function readFails(env, key) {
  const v = await env.VERTALER_KV.get(key);
  return v ? parseInt(v, 10) || 0 : 0;
}

async function addFail(env, key) {
  const n = (await readFails(env, key)) + 1;
  await env.VERTALER_KV.put(key, String(n), { expirationTtl: FAIL_WINDOW });
  return n;
}

async function handleLogin(body, env, request) {
  const username = (body.username || '').trim();
  const password = body.password || '';
  const remember = !!body.remember;
  if (!username || !password) {
    return jsonResponse({ error: 'Gebruikersnaam en wachtwoord verplicht.' }, 400);
  }
  if (!env.PASSWORD_PEPPER) {
    return jsonResponse({ error: 'PASSWORD_PEPPER ontbreekt (wrangler secret put PASSWORD_PEPPER).' }, 500);
  }

  const ip = request.headers.get('CF-Connecting-IP') || 'onbekend';
  const failUserKey = `fail:user:${username.toLowerCase()}`;
  const failIpKey = `fail:ip:${ip}`;
  const [failsUser, failsIp] = await Promise.all([readFails(env, failUserKey), readFails(env, failIpKey)]);
  if (failsUser >= MAX_FAILS_USER || failsIp >= MAX_FAILS_IP) {
    return jsonResponse({ error: 'Te veel mislukte pogingen. Probeer het binnen 15 minuten opnieuw.', code: 'too_many_attempts' }, 429);
  }

  const fail = async () => {
    await Promise.all([addFail(env, failUserKey), addFail(env, failIpKey)]);
    return jsonResponse({ error: 'Onbekende gebruikersnaam of fout wachtwoord.' }, 401);
  };

  const user = await env.VERTALER_KV.get(`user:${username}`, 'json');
  // Ook bij een onbekende gebruiker een hash berekenen, zodat de responstijd
  // niet verraadt of een gebruikersnaam bestaat.
  const hash = await sha256Hex(password + env.PASSWORD_PEPPER);
  if (!user || !timingSafeEqual(hash, user.passwordHash)) return fail();

  if (failsUser) await env.VERTALER_KV.delete(failUserKey);
  const token = randomToken();
  const ttl = remember ? SESSION_TTL_REMEMBER : SESSION_TTL_DEFAULT;
  await env.VERTALER_KV.put(
    `session:${token}`,
    JSON.stringify({ username: user.username, role: user.role }),
    { expirationTtl: ttl }
  );
  return jsonResponse({ token, username: user.username, role: user.role });
}

async function handleLogout(request, env) {
  const token = request.headers.get('X-Auth-Token') || '';
  if (token) await env.VERTALER_KV.delete(`session:${token}`);
  return jsonResponse({ ok: true });
}

// Admin laadt een Users-Excel op (kolommen USER/PW/ROLE). Per rij:
//   ROLE admin/user  -> gebruiker toevoegen of bijwerken (wachtwoord verplicht);
//   ROLE delete      -> gebruiker verwijderen (wachtwoord niet nodig).
// Gebruikers die niet in het bestand staan, blijven ongewijzigd.
const DELETE_ROLES = ['delete', 'verwijderen', 'verwijder', 'supprimer', 'löschen', 'eliminar', 'elimina'];

async function handleAdminUpsertUsers(body, env, session) {
  if (!env.PASSWORD_PEPPER) {
    return jsonResponse({ error: 'PASSWORD_PEPPER ontbreekt (wrangler secret put PASSWORD_PEPPER).' }, 500);
  }
  const users = Array.isArray(body.users) ? body.users : [];
  if (!users.length) return jsonResponse({ error: 'Geen gebruikers ontvangen.' }, 400);

  let updated = 0;
  let deleted = 0;
  const errors = [];
  for (const raw of users) {
    const username = (raw.username || '').trim();
    const password = (raw.password || '').trim();
    const role = (raw.role || '').trim().toLowerCase();
    if (!username) {
      errors.push('(geen gebruikersnaam) — overgeslagen.');
      continue;
    }
    if (DELETE_ROLES.includes(role)) {
      if (username === session.username) {
        errors.push(`${username}: je kan jezelf niet verwijderen — overgeslagen.`);
        continue;
      }
      await env.VERTALER_KV.delete(`user:${username}`);
      deleted++;
      continue;
    }
    if (role !== 'admin' && role !== 'user') {
      errors.push(`${username}: onbekende rol "${raw.role || ''}" (verwacht admin, user of delete) — overgeslagen.`);
      continue;
    }
    if (password.length < MIN_PASSWORD_LENGTH) {
      errors.push(`${username}: wachtwoord ontbreekt of is korter dan ${MIN_PASSWORD_LENGTH} tekens — overgeslagen.`);
      continue;
    }
    if (username === session.username && role !== 'admin') {
      errors.push(`${username}: je kan je eigen admin-rechten niet afnemen — overgeslagen.`);
      continue;
    }
    try {
      const passwordHash = await sha256Hex(password + env.PASSWORD_PEPPER);
      await env.VERTALER_KV.put(`user:${username}`, JSON.stringify({ username, passwordHash, role }));
      updated++;
    } catch (err) {
      errors.push(`${username}: ${err.message}`);
    }
  }
  return jsonResponse({ updated, deleted, total: users.length, errors });
}

// ============================================================
// Helpers
// ============================================================

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': ALLOWED_ORIGIN,
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, X-Auth-Token',
  };
}

function jsonResponse(obj, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { 'content-type': 'application/json', ...corsHeaders() },
  });
}
