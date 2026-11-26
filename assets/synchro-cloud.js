/*!
 * MEP dashboards · Synchro cloud data layer            assets/synchro-cloud.js
 *
 * Reads live SYNCHRO 4D data for the dashboards from Bentley's iTwin Platform
 * Schedules API (api.bentley.com): the same cloud project the SYNCHRO 4D Power BI
 * connector reads. Each person signs in with their own Bentley account (OAuth 2.0
 * authorisation code with PKCE, the flow Bentley prescribes for browser apps).
 * Access is read only: nothing is ever written back to Synchro.
 *
 * It also parses the manual Excel export, so the dashboards keep working offline.
 */
(function (global) {
  'use strict';

  const VERSION = '1.2.0';
  const SCRIPT_SRC = (document.currentScript && document.currentScript.src) || '';
  const ROOT = SCRIPT_SRC ? new URL('../', SCRIPT_SRC) : new URL('./', location.href);                 // the dashboards folder
  const IS_HTTP = /^https?:$/.test(location.protocol);
  const ACCEPT = 'application/vnd.bentley.itwin-platform.v1+json';
  const IN_FRAME = global.parent !== global;

  // Opened as a downloaded file (for example straight from SharePoint)? Bentley sign-in cannot
  // return to a file, so hop to the team's web address for the same page, where it can.
  // Published copies carry their data inside and are left alone.
  (function hopToWebAddress() {
    const web = (global.MEP_DASHBOARD_CONFIG && global.MEP_DASHBOARD_CONFIG.webUrl) || '';
    if (IS_HTTP || !web || document.getElementById('mep-snapshot')) return;
    try {
      const page = decodeURIComponent(location.pathname.split('/').pop() || '');
      const base = web.endsWith('/') ? web : web + '/';
      const target = new URL(/\.html?$/i.test(page) ? encodeURIComponent(page) : '', base);
      target.hash = location.hash;
      location.replace(target.href);
    } catch (e) { /* stay on the file; the Excel route still works */ }
  })();

  const KEYS = {
    user: 'mep.sc.config.v1',
    token: 'mep.sc.token.v1',
    pending: 'mep.sc.pending.',
    authError: 'mep.sc.authError',
    reopen: 'mep.sc.reopenSettings',
    reopenPublish: 'mep.sc.reopenPublish',
    autoSignIn: 'mep.sc.autoSignIn',
    publish: 'mep.sc.publish.v1',
    staticData: 'mep.sc.static.v1',
    dataset: 'mep.sc.dataset.v1',
    me: 'mep.sc.me.v1',
    complete: 'mep.completeStatuses.v1',
    snapshots: 'mep.snapshots.v1',
    workbook: 'cable-tray-workbook-cache'          // same key the original dashboard used
  };

  const DEFAULTS = {
    clientId: '', projectId: '', scheduleId: '', projectLabel: '',
    statusCategoryId: '', lengthFieldId: 'auto', lengthUnit: 'auto', objectSource: 'auto',
    scopeMode: 'status', scopeValue: '', unrecordedStatus: 'Not Started', linkProgramme: true,
    refreshMinutes: 10, fullReloadHours: 6, staleAfterDays: 3, latestUrl: '', webUrl: '',
    apiBase: 'https://api.bentley.com', imsBase: 'https://ims.bentley.com', scopes: 'itwin-platform'
  };

  /* ------------------------------------------------------------------ storage */
  const store = {
    get(key, fallback = null) {
      try { const raw = localStorage.getItem(key); return raw == null ? fallback : JSON.parse(raw); } catch (e) { return fallback; }
    },
    set(key, value) {
      try { localStorage.setItem(key, JSON.stringify(value)); return true; } catch (e) { return false; }
    },
    remove(key) { try { localStorage.removeItem(key); } catch (e) { /* ignore */ } }
  };
  const session = {
    get(key) { try { return sessionStorage.getItem(key); } catch (e) { return null; } },
    set(key, v) { try { sessionStorage.setItem(key, v); } catch (e) { /* ignore */ } },
    take(key) { const v = session.get(key); try { sessionStorage.removeItem(key); } catch (e) { /* ignore */ } return v; },
    remove(key) { try { sessionStorage.removeItem(key); } catch (e) { /* ignore */ } }
  };

  function config() {
    return Object.assign({}, DEFAULTS, global.MEP_DASHBOARD_CONFIG || {}, store.get(KEYS.user, {}));
  }
  function saveUserConfig(patch) {
    store.set(KEYS.user, Object.assign(store.get(KEYS.user, {}), patch));
  }
  function isCloudConfigured(cfg = config()) {
    return Boolean(cfg.clientId && cfg.scheduleId);
  }
  function mode() {
    if (document.getElementById('mep-snapshot')) return 'snapshot';
    return IS_HTTP && isCloudConfigured() ? 'cloud' : 'offline';
  }
  function redirectUri() { return ROOT.href.split('#')[0].split('?')[0]; }

  /* ------------------------------------------------------------------ formatting */
  const numberFormats = {};
  function nf(dp) {
    return numberFormats[dp] || (numberFormats[dp] = new Intl.NumberFormat('en-GB', { maximumFractionDigits: dp }));
  }
  const toDate = (v) => (v instanceof Date ? v : new Date(v));
  const fmt = {
    num(v, dp = 1) { return Number.isFinite(v) ? nf(dp).format(v) : '-'; },
    int(v) { return Number.isFinite(v) ? nf(0).format(Math.round(v)) : '-'; },
    pct(v, dp = 1) { return Number.isFinite(v) ? `${nf(dp).format(v)}%` : '-'; },
    date(v) { return v ? toDate(v).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }) : '-'; },
    shortDate(v) { return v ? toDate(v).toLocaleDateString('en-GB', { day: 'numeric', month: 'short' }) : '-'; },
    time(v) { return v ? toDate(v).toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' }) : '-'; },
    dateTime(v) { return v ? `${fmt.date(v)}, ${fmt.time(v)}` : '-'; },
    ago(v) {
      if (!v) return '-';
      const s = Math.round((Date.now() - toDate(v).getTime()) / 1000);
      if (s < 60) return 'just now';
      if (s < 3600) return `${Math.round(s / 60)} min ago`;
      if (s < 86400) return `${Math.round(s / 3600)} h ago`;
      const d = Math.round(s / 86400);
      return `${d} day${d === 1 ? '' : 's'} ago`;
    }
  };
  const clean = (v) => (v == null ? '' : String(v).trim());
  const escapeHtml = (v) => clean(v).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  /* ------------------------------------------------------------------ status rules */
  function statusTone(name) {
    const v = clean(name).toLowerCase();
    if (/incorrect|defect|reject|rework|snag|fail|damage|wrong|ncr|non.?conform/.test(v)) return 'defect';
    if (!v || /not\s*started|no status|unspecified|outstanding|to do|unrecorded/.test(v)) return 'todo';
    if (/installed|complete|done|finished|signed.?off|approved|accepted|handed/.test(v)) return 'complete';
    if (/set\s*out|ready|marked|released|in progress|started|partial|first fix|1st fix|delivered|on site/.test(v)) return 'ready';
    return 'other';
  }
  // The original dashboard's rule for which statuses count as installed by default.
  function looksComplete(status) {
    const value = clean(status).toLowerCase();
    if (/\b(partial|pending|progress|incomplete|incorrect|not|hold|no|n\/a)\b/.test(value)) return false;
    return /\b(installed|complete|completed|done|yes|true)\b|100\s*%/.test(value);
  }
  const TONE_ORDER = { todo: 0, ready: 1, complete: 2, defect: 3, other: 4 };
  function sortStatuses(names) {
    return [...new Set(names)].sort((a, b) => (TONE_ORDER[statusTone(a)] - TONE_ORDER[statusTone(b)]) || a.localeCompare(b));
  }

  // Which statuses count as complete: one choice shared by both dashboards.
  const completion = {
    get(labels) {
      const snap = readSnapshot();
      if (snap && snap.meta) return new Set(labels.filter((l) => (snap.meta.complete || []).includes(l)));
      const saved = store.get(KEYS.complete, null);
      const result = new Set();
      const known = saved && Array.isArray(saved.known) ? saved.known : [];
      const chosen = saved && Array.isArray(saved.complete) ? saved.complete : [];
      for (const label of labels) {
        if (known.includes(label)) { if (chosen.includes(label)) result.add(label); }
        else if (looksComplete(label)) result.add(label);
      }
      return result;
    },
    set(completeSet, labels) {
      if (readSnapshot()) return;
      const saved = store.get(KEYS.complete, {}) || {};
      const known = new Set([...(saved.known || []), ...labels]);
      const complete = new Set((saved.complete || []).filter((l) => !labels.includes(l)));
      for (const l of completeSet) complete.add(l);
      store.set(KEYS.complete, { known: [...known], complete: [...complete] });
    }
  };

  /* ------------------------------------------------------------------ names and object types */
  function splitName(name) {
    const text = clean(name);
    const m = text.match(/^(.*?)\s*\[([^\]]+)\]\s*$/);
    const base = m ? m[1] : text;
    const parts = base.split(' : ');
    return { family: clean(parts[0] || base), type: clean(parts.slice(1).join(' : ')), ref: m ? m[2] : '' };
  }
  function prettyFamily(family) {
    return clean(family).replace(/_/g, ' ').replace(/\s+/g, ' ');
  }
  const FITTING_WORDS = ['bend', 'tee', 'reducer', 'riser', 'cross', 'union', 'elbow', 'junction', 'coupler', 'wye',
    'transition', 'offset', 'fitting', 'fittings', 'connector', 'adaptor', 'adapter', 'drop'];
  const DUCT_WORDS = ['trunking', 'salamandre', 'duct', 'ducts', 'dw144', 'tophat', 'conduit', 'disttrunking'];
  const TRAY_WORDS = ['tray', 'trays', 'ladder', 'channel', 'basket', 'swifts', 'unistrut'];
  // Built-in rules for HS2 MEP containment names. Matches the Excel export's "Object"
  // column for about 98% of elements; used only when Synchro has no category field.
  function classifyByName(name) {
    const { family } = splitName(name);
    const tokens = family.replace(/([a-z])([A-Z])/g, '$1 $2').toLowerCase().split(/[^a-z0-9]+/).filter(Boolean);
    const has = (list) => tokens.some((t) => list.includes(t));
    const fitting = has(FITTING_WORDS);
    const duct = has(DUCT_WORDS);
    const tray = has(TRAY_WORDS);
    const kind = tray ? 'Cable Tray' : duct ? 'Cable Duct' : 'Containment';
    return fitting ? `${kind} Fitting` : kind;
  }

  /* ------------------------------------------------------------------ Excel export */
  function parseLength(value) {
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    const text = clean(value).replace(/,/g, '');
    const match = text.match(/-?\d+(?:\.\d+)?/);
    return match ? Number(match[0]) : null;
  }
  function isHeaderRow(row) {
    return /length/i.test(clean(row[1])) && /(install|status)/i.test(clean(row[4]));
  }
  function detectUnit(header) {
    const text = header ? clean(header[1]) : '';
    const unit = text.match(/\(([^)]+)\)|\b(mm|cm|m|ft|feet|in|inch|inches)\b/i);
    return unit ? (unit[1] || unit[2]).trim() : 'metres';
  }
  function shortUnit(unit) {
    const u = clean(unit).toLowerCase();
    if (/^(m|metres?|meters?)$/.test(u)) return 'm';
    if (/^(mm|millimetres?|millimeters?)$/.test(u)) return 'mm';
    if (/^(ft|feet|foot)$/.test(u)) return 'ft';
    return unit || 'm';
  }
  // Same column rules as the original dashboard: A element, B length, C quantity, D object, E status.
  function parseSheet(rawRows) {
    const headerIndex = rawRows.findIndex(isHeaderRow);
    const start = headerIndex >= 0 ? headerIndex + 1 : 0;
    const unit = detectUnit(headerIndex >= 0 ? rawRows[headerIndex] : null);
    const rows = [];
    for (let index = start; index < rawRows.length; index += 1) {
      const row = rawRows[index] || [];
      const object = clean(row[3]);
      const isFitting = /fitting/i.test(object);
      const sourceLength = parseLength(row[1]);
      if (!object && sourceLength === null) continue;
      if (!isFitting && sourceLength === null) continue;
      const element = clean(row[0]);
      const names = splitName(element);
      rows.push({
        id: `x${index + 1}`, rowNumber: index + 1, element, item: object, family: names.family, type: names.type, ref: names.ref,
        quantity: clean(row[2]), length: isFitting ? 0 : sourceLength, fittingCount: isFitting ? 1 : 0, isFitting,
        status: clean(row[4]) || 'Unspecified', statusColor: null, statusDate: null, recorded: true
      });
    }
    return { rows, unit };
  }
  function datasetFromSheet(aoa, filename, sheetName, sheetNames, modifiedAt) {
    const { rows, unit } = parseSheet(aoa || []);
    const fetchedAt = new Date().toISOString();
    return {
      version: 1, source: 'excel', label: `${filename} / ${sheetName}`, filename, sheetName, sheetNames: sheetNames || [sheetName],
      fetchedAt, modifiedAt: modifiedAt || null, asAt: modifiedAt || fetchedAt, lastSync: null, unit, unitShort: shortUnit(unit), rows,
      statuses: sortStatuses(rows.map((r) => r.status)).map((name) => ({ name, color: null })),
      history: [], programmeLinked: false, notes: []
    };
  }
  function datasetFromCache(cache) {
    const name = cache.sheetNames.find((n) => parseSheet(cache.sheets[n] || []).rows.length) || cache.sheetNames[0];
    return datasetFromSheet(cache.sheets[name], cache.filename, name, cache.sheetNames, cache.modifiedAt);
  }
  function workbookToCache(workbook, filename, lastModified) {
    const sheets = {};
    for (const name of workbook.SheetNames) {
      sheets[name] = global.XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '', raw: true });
    }
    const modifiedAt = Number.isFinite(lastModified) && lastModified > 0 ? new Date(lastModified).toISOString() : null;
    const payload = { filename, sheetNames: workbook.SheetNames, sheets, savedAt: new Date().toISOString(), modifiedAt };
    store.set(KEYS.workbook, payload);
    return payload;
  }
  function cachedWorkbook() {
    const cached = store.get(KEYS.workbook, null);
    return cached && Array.isArray(cached.sheetNames) && cached.sheets ? cached : null;
  }

  /* ------------------------------------------------------------------ sign-in (OAuth 2.0 + PKCE) */
  function b64url(bytes) {
    let bin = '';
    const arr = new Uint8Array(bytes);
    for (let i = 0; i < arr.length; i += 1) bin += String.fromCharCode(arr[i]);
    return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }
  function randomString(n = 48) {
    const a = new Uint8Array(n);
    crypto.getRandomValues(a);
    return b64url(a);
  }
  async function pkcePair() {
    const verifier = randomString(48);
    const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier));
    return { verifier, challenge: b64url(digest) };
  }
  function prunePending() {
    try {
      for (let i = localStorage.length - 1; i >= 0; i -= 1) {
        const key = localStorage.key(i);
        if (key && key.startsWith(KEYS.pending)) {
          const v = store.get(key, null);
          if (!v || Date.now() - (v.created || 0) > 15 * 60e3) localStorage.removeItem(key);
        }
      }
    } catch (e) { /* ignore */ }
  }
  async function authorizeUrl({ prompt, returnTo }) {
    const cfg = config();
    if (!IS_HTTP) throw new Error('Open the dashboards with Start Dashboards.bat to sign in to Bentley.');
    if (!cfg.clientId) throw new Error('Add the Bentley app client ID in Synchro settings first.');
    prunePending();
    const { verifier, challenge } = await pkcePair();
    const state = randomString(24);
    store.set(KEYS.pending + state, { verifier, returnTo: returnTo || null, created: Date.now(), silent: prompt === 'none' });
    const q = new URLSearchParams({
      response_type: 'code', client_id: cfg.clientId, redirect_uri: redirectUri(), scope: cfg.scopes,
      state, code_challenge: challenge, code_challenge_method: 'S256'
    });
    if (prompt) q.set('prompt', prompt);
    return `${cfg.imsBase}/connect/authorize?${q}`;
  }
  async function signIn(returnTo = location.href, prompt) {
    location.assign(await authorizeUrl({ returnTo, prompt }));
  }
  function readToken() {
    const t = store.get(KEYS.token, null);
    return t && t.accessToken && t.expiresAt > Date.now() + 60e3 ? t : null;
  }
  function saveToken(json) {
    store.set(KEYS.token, { accessToken: json.access_token, expiresAt: Date.now() + (Number(json.expires_in) || 3600) * 1000, scope: json.scope || '' });
    session.remove(KEYS.autoSignIn);
  }
  function signOut() {
    store.remove(KEYS.token);
    store.remove(KEYS.me);
  }
  async function completeAuth(search) {
    const cfg = config();
    const p = new URLSearchParams(search);
    const state = p.get('state') || '';
    const pending = store.get(KEYS.pending + state, null);
    store.remove(KEYS.pending + state);
    if (!pending) throw new Error('That sign-in response did not match a request from this browser. Please sign in again.');
    if (p.get('error')) {
      const err = new Error(p.get('error_description') || p.get('error'));
      err.code = p.get('error');
      err.returnTo = pending.returnTo;
      throw err;
    }
    const body = new URLSearchParams({
      grant_type: 'authorization_code', code: p.get('code') || '', client_id: cfg.clientId,
      redirect_uri: redirectUri(), code_verifier: pending.verifier
    }).toString();
    // A form-encoded POST with no custom headers is a CORS "simple" request, which Bentley's
    // sign-in server answers for browser apps.
    let res;
    try {
      res = await fetch(`${cfg.imsBase}/connect/token`, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body });
    } catch (e) {
      // Bentley's sign-in server only lets a web address read the token once that address is
      // on its allow list, which can lag behind adding it as a redirect URI on the app.
      const err = new Error(`Bentley accepted the sign-in but has not yet allowed this web address (${location.origin}) to collect it. If the address was only just added to the app's redirect URIs, try again later.`);
      err.returnTo = pending.returnTo;
      throw err;
    }
    const json = await res.json().catch(() => ({}));
    if (!res.ok || !json.access_token) {
      const err = new Error(json.error_description || json.error || `Bentley sign-in failed (${res.status}).`);
      err.returnTo = pending.returnTo;
      throw err;
    }
    saveToken(json);
    store.remove(KEYS.me);
    return { returnTo: pending.returnTo };
  }
  async function handleRedirect() {
    if (!IS_HTTP) return false;
    const params = new URLSearchParams(location.search);
    if (!params.get('state') || !(params.has('code') || params.has('error'))) return false;
    if (IN_FRAME) {
      try { global.parent.postMessage({ type: 'mep-sc-auth', search: location.search }, location.origin); } catch (e) { /* ignore */ }
      return true;
    }
    let target = redirectUri();
    try {
      const result = await completeAuth(location.search);
      target = result.returnTo || target;
    } catch (e) {
      session.set(KEYS.authError, e.message);
      target = e.returnTo || target;
    }
    location.replace(target);
    return new Promise(() => {});   // the page is navigating away
  }
  let renewing = null;
  function silentRenew(timeoutMs = 12000) {
    if (renewing) return renewing;
    renewing = new Promise((resolve, reject) => {
      let done = false;
      let frame = null;
      let timer = null;
      const finish = (err) => {
        if (done) return;
        done = true;
        clearTimeout(timer);
        global.removeEventListener('message', onMessage);
        if (frame) frame.remove();
        renewing = null;
        if (err) reject(err); else resolve(true);
      };
      const onMessage = (ev) => {
        if (ev.origin !== location.origin || !ev.data || ev.data.type !== 'mep-sc-auth') return;
        completeAuth(ev.data.search).then(() => finish(), finish);
      };
      global.addEventListener('message', onMessage);
      timer = setTimeout(() => finish(new Error('Bentley sign-in needs a refresh.')), timeoutMs);
      authorizeUrl({ prompt: 'none' }).then((url) => {
        frame = document.createElement('iframe');
        frame.hidden = true;
        frame.setAttribute('aria-hidden', 'true');
        frame.title = 'Bentley sign-in refresh';
        frame.src = url;
        document.body.appendChild(frame);
      }, finish);
    });
    return renewing;
  }
  class AuthRequiredError extends Error {
    constructor(message) { super(message || 'Sign in to Bentley to load live Synchro data.'); this.name = 'AuthRequiredError'; this.authRequired = true; }
  }
  async function getToken() {
    const t = readToken();
    if (t) return t.accessToken;
    if (!IS_HTTP || !config().clientId) throw new AuthRequiredError();
    try {
      await silentRenew();
    } catch (e) {
      throw new AuthRequiredError();
    }
    const renewed = readToken();
    if (!renewed) throw new AuthRequiredError();
    return renewed.accessToken;
  }
  function isSignedIn() { return Boolean(readToken()); }

  /* ------------------------------------------------------------------ Schedules API client */
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  function friendlyApiError(status, body, url) {
    const code = body && body.error && body.error.code;
    const message = body && body.error && body.error.message;
    let err;
    if (status === 401) err = new AuthRequiredError('Your Bentley sign-in has expired. Sign in again to continue.');
    else if (status === 403) {
      err = new Error(`Your Bentley account does not have access to this project or schedule (403${code ? `, ${code}` : ''}).`);
      err.noAccess = true;
    } else if (status === 404 && /schedule/i.test(code || '')) err = new Error('Schedule not found. Check the project uses scheduling services and the schedule chosen in Synchro settings.');
    else if (status === 404) err = new Error(`Not found: ${message || url}`);
    else err = new Error(`Bentley returned ${status}${message ? `: ${message}` : ''}.`);
    err.status = status;
    return err;
  }
  async function api(pathOrUrl, attempt = 0) {
    const cfg = config();
    const url = /^https?:/i.test(pathOrUrl) ? pathOrUrl : cfg.apiBase + pathOrUrl;
    if (!url.toLowerCase().startsWith(cfg.apiBase.toLowerCase())) throw new Error('Refusing to send the Bentley token outside the API host.');
    const token = await getToken();
    let res;
    try {
      res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: ACCEPT }, cache: 'no-store' });
    } catch (e) {
      throw new Error('Could not reach Bentley\'s cloud. Check the network connection.');
    }
    if (res.status === 401 && attempt === 0) {
      store.remove(KEYS.token);
      return api(pathOrUrl, attempt + 1);
    }
    if ((res.status === 429 || res.status >= 500) && attempt < 3) {
      const wait = Math.min(30, Number(res.headers.get('retry-after')) || 2 ** attempt * 2);
      await sleep(wait * 1000);
      return api(pathOrUrl, attempt + 1);
    }
    if (!res.ok) throw friendlyApiError(res.status, await res.json().catch(() => null), url);
    const json = await res.json();
    json.__continuation = res.headers.get('continuation-token');
    return json;
  }
  function query(params) {
    return Object.entries(params)
      .filter(([, v]) => v !== undefined && v !== null && v !== '')
      .map(([k, v]) => `${k}=${encodeURIComponent(v)}`).join('&');
  }
  async function getAll(path, key, params = {}, onPage) {
    const cfg = config();
    const first = `${cfg.apiBase}${path}${path.includes('?') ? '&' : '?'}${query(Object.assign({ $top: 10000 }, params))}`;
    const items = [];
    let url = first;
    let lastToken = null;
    for (let page = 0; url; page += 1) {
      if (page > 400) throw new Error(`Too many pages while reading ${key}.`);
      const json = await api(url);
      const batch = json[key] || [];
      for (const item of batch) items.push(item);
      if (onPage) onPage(items.length);
      const next = json._links && json._links.next && json._links.next.href;
      if (next) url = next;
      else if (json.__continuation && json.__continuation !== lastToken && batch.length) {
        lastToken = json.__continuation;
        url = `${first}&$continuationToken=${encodeURIComponent(lastToken)}`;
      } else url = null;
    }
    return items.filter((it) => !it.deleted);
  }
  async function mapLimit(list, limit, fn) {
    const out = new Array(list.length);
    let next = 0;
    const workers = Array.from({ length: Math.min(limit, list.length) }, async () => {
      while (next < list.length) {
        const k = next;
        next += 1;
        out[k] = await fn(list[k], k);
      }
    });
    await Promise.all(workers);
    return out;
  }
  async function whoAmI() {
    const cached = store.get(KEYS.me, null);
    if (cached && cached.token === (readToken() || {}).accessToken) return cached.user;
    const json = await api('/users/me');
    const user = json.user || {};
    store.set(KEYS.me, { token: (readToken() || {}).accessToken, user });
    return user;
  }
  const listProjects = (search) => api(`/itwins/?${query({ subClass: 'Project', $top: 50, $search: search || undefined })}`).then((j) => j.iTwins || []);
  const listSchedules = (projectId) => getAll('/schedules', 'schedules', { iTwinId: projectId, $top: 100 });
  const scheduleBase = (id) => `/schedules/${encodeURIComponent(id)}`;

  /* ------------------------------------------------------------------ cloud dataset */
  function normColour(c) {
    const v = clean(c);
    return /^#[0-9a-f]{3,8}$/i.test(v) ? v.slice(0, 7) : null;
  }
  function pickCategory(categories, id) {
    return categories.find((c) => c.id === id) || categories.find((c) => /install/i.test(c.name)) || categories[0] || null;
  }
  function resolveLengthField(fields, setting) {
    if (setting === 'none') return null;
    const usable = fields.filter((f) => ['Entity3d', 'Resource'].includes(f.category) && ['Float', 'Integer', 'String'].includes(f.type));
    if (setting && setting !== 'auto') return usable.find((f) => f.id === setting) || null;
    const rank = (f) => (f.category === 'Entity3d' ? 0 : 1) + (f.type === 'String' ? 2 : 0);
    const exact = usable.filter((f) => /^length$/i.test(clean(f.name))).sort((a, b) => rank(a) - rank(b));
    const loose = usable.filter((f) => /length/i.test(f.name) && !/unit/i.test(f.name)).sort((a, b) => rank(a) - rank(b));
    return exact[0] || loose[0] || null;
  }
  function resolveObjectSource(fields, setting) {
    if (setting === 'rules' || setting === 'parent') return { kind: setting };
    const usable = fields.filter((f) => ['Entity3d', 'Resource'].includes(f.category) && f.type === 'String');
    if (setting && setting.startsWith('field:')) {
      const f = usable.find((x) => x.id === setting.slice(6));
      return f ? { kind: 'field', field: f } : { kind: 'rules' };
    }
    const f = usable.find((x) => /^(category|object|object type|revit category|element category)$/i.test(clean(x.name)));
    return f ? { kind: 'field', field: f } : { kind: 'rules' };
  }
  function configSignature(cfg) {
    return JSON.stringify(['scheduleId', 'statusCategoryId', 'lengthFieldId', 'lengthUnit', 'objectSource', 'scopeMode', 'scopeValue', 'linkProgramme', 'apiBase'].map((k) => cfg[k]));
  }
  function nameMatcher(text) {
    const t = clean(text);
    if (!t) return () => true;
    const m = t.match(/^\/(.+)\/([a-z]*)$/i);
    if (m) {
      try { const re = new RegExp(m[1], m[2].includes('i') ? m[2] : `${m[2]}i`); return (s) => re.test(s); } catch (e) { /* fall through */ }
    }
    const parts = t.split('|').map((p) => p.trim().toLowerCase()).filter(Boolean);
    return (s) => parts.some((p) => s.toLowerCase().includes(p));
  }
  function histCompare(a, b) {
    return (Date.parse(a.date) || 0) - (Date.parse(b.date) || 0) || (Date.parse(a.recordDate) || 0) - (Date.parse(b.recordDate) || 0);
  }

  async function fieldValuesByResource(S, field, entityLinks, progress) {
    const out = new Map();
    if (!field) return out;
    if (field.category === 'Resource') {
      const values = await getAll(`${S}/resources/user-field-values`, 'userFieldValues', { userFieldId: field.id }, (n) => progress(`Reading ${field.name} (${fmt.int(n)})`));
      for (const v of values) out.set(v.resourceId, [v.value]);
      return out;
    }
    const values = await getAll(`${S}/entities-3d/user-field-values`, 'userFieldValues', { userFieldId: field.id }, (n) => progress(`Reading ${field.name} (${fmt.int(n)})`));
    const byEntity = new Map(values.map((v) => [v.entity3dId, v.value]));
    for (const [resourceId, entityIds] of entityLinks) {
      const vals = entityIds.map((id) => byEntity.get(id)).filter((v) => v != null && v !== '');
      if (vals.length) out.set(resourceId, vals);
    }
    return out;
  }

  async function loadStatic(cfg, progress) {
    const S = scheduleBase(cfg.scheduleId);
    const notes = [];
    progress('Opening the Synchro schedule');
    const [scheduleJson, userFields, categories, sets, items] = await Promise.all([
      api(S), getAll(`${S}/user-fields`, 'userFields'), getAll(`${S}/resource-status-categories`, 'resourceStatusCategories'),
      getAll(`${S}/resource-status-sets`, 'resourceStatusSets'), getAll(`${S}/resource-status-items`, 'resourceStatusItems')
    ]);
    const schedule = scheduleJson.schedule || scheduleJson;
    const category = pickCategory(categories, cfg.statusCategoryId);
    if (!category) notes.push('This schedule has no resource status categories, so every item shows the default status.');
    const setIds = new Set(sets.filter((s) => category && s.statusCategoryId === category.id).map((s) => s.id));
    const statusItems = items.filter((i) => setIds.has(i.statusSetId)).map((i, order) => ({ id: i.id, name: clean(i.name), color: normColour(i.color), order }));

    // Resources in scope
    const scopeMode = cfg.scopeMode === 'auto' || !cfg.scopeMode ? 'status' : cfg.scopeMode;
    let all = null;
    let candidates;
    const readAll = async () => all || (all = await getAll(`${S}/resources`, 'resources', {}, (n) => progress(`Reading resources (${fmt.int(n)})`)));
    if (scopeMode === 'group' && cfg.scopeValue) {
      candidates = await getAll(`${S}/resource-groups/${encodeURIComponent(cfg.scopeValue)}/resources`, 'resources', {}, (n) => progress(`Reading resource group (${fmt.int(n)})`));
    } else {
      candidates = await readAll();
    }
    const objectSource = resolveObjectSource(userFields, cfg.objectSource);
    if (objectSource.kind === 'parent') await readAll();
    const byId = new Map((all || candidates).map((r) => [r.id, r]));
    const hasChildren = new Set((all || candidates).filter((r) => r.parentId).map((r) => r.parentId));
    const isLeaf = (r) => !hasChildren.has(r.id);
    let scope;
    if (scopeMode === 'descendants') {
      const root = cfg.scopeValue;
      const children = new Map();
      for (const r of all) if (r.parentId) (children.get(r.parentId) || children.set(r.parentId, []).get(r.parentId)).push(r);
      scope = [];
      const stack = [...(children.get(root) || [])];
      while (stack.length) {
        const r = stack.pop();
        if (children.has(r.id)) stack.push(...children.get(r.id)); else scope.push(r);
      }
      if (!root) notes.push('Choose a resource folder in Synchro settings.');
    } else if (scopeMode === 'name') {
      const test = nameMatcher(cfg.scopeValue);
      scope = candidates.filter((r) => isLeaf(r) && test(clean(r.name)));
    } else if (scopeMode === 'status') {
      progress('Finding items with an installation status');
      const latest = category ? await getAll(`${S}/resources/status-history`, 'statusHistory', { categoryIds: category.id, isLatest: 'true' }, (n) => progress(`Finding items with a status (${fmt.int(n)})`)) : [];
      const ids = new Set(latest.map((h) => h.resourceId));
      scope = candidates.filter((r) => ids.has(r.id));
    } else {
      scope = candidates.filter(isLeaf);
    }
    scope.sort((a, b) => clean(a.name).localeCompare(clean(b.name), 'en-GB', { numeric: true }));

    // 3D object links, only when a 3D user field is needed
    const lengthField = resolveLengthField(userFields, cfg.lengthFieldId);
    const needEntities = [lengthField, objectSource.field].some((f) => f && f.category === 'Entity3d');
    const entityLinks = new Map();
    if (needEntities && scope.length) {
      const scopeIds = new Set(scope.map((r) => r.id));
      const links = await getAll(`${S}/resources/entity-3d-assignments`, 'entity3dAssignments', {}, (n) => progress(`Linking 3D objects (${fmt.int(n)})`));
      for (const l of links) if (scopeIds.has(l.resourceId)) (entityLinks.get(l.resourceId) || entityLinks.set(l.resourceId, []).get(l.resourceId)).push(l.entity3dId);
    }
    const lengthValues = await fieldValuesByResource(S, lengthField, entityLinks, progress);
    const objectValues = objectSource.kind === 'field' ? await fieldValuesByResource(S, objectSource.field, entityLinks, progress) : null;
    if (!lengthField && cfg.lengthFieldId !== 'none') notes.push('No length field found, so lengths count as zero. Choose the length field in Synchro settings.');

    // Length unit
    const parsed = new Map();
    for (const [id, vals] of lengthValues) {
      let sum = 0; let seen = false; let unitHint = null;
      for (const v of vals) {
        const n = parseLength(v);
        if (n !== null) { sum += n; seen = true; }
        if (/mm\b/i.test(String(v))) unitHint = 'mm'; else if (/ft\b|'/.test(String(v))) unitHint = 'ft';
      }
      if (seen) parsed.set(id, { value: sum, unitHint });
    }
    let unit = cfg.lengthUnit;
    if (!unit || unit === 'auto') {
      const nums = [...parsed.values()].map((p) => p.value).filter((n) => n > 0).sort((a, b) => a - b);
      const median = nums.length ? nums[Math.floor(nums.length / 2)] : 0;
      unit = median > 100 ? 'mm' : 'm';
    }
    const toMetres = (p) => {
      if (!p) return null;
      const u = p.unitHint || unit;
      return u === 'mm' ? p.value / 1000 : u === 'ft' ? p.value * 0.3048 : p.value;
    };

    const rows = scope.map((r) => {
      const element = clean(r.name);
      const names = splitName(element);
      let item;
      if (objectSource.kind === 'field') item = clean((objectValues.get(r.id) || [])[0]) || classifyByName(element);
      else if (objectSource.kind === 'parent') item = clean((byId.get(r.parentId) || {}).name) || classifyByName(element);
      else item = classifyByName(element);
      const metres = toMetres(parsed.get(r.id));
      const isFitting = /fitting/i.test(item) || (metres === null && /fitting/i.test(classifyByName(element)));
      return {
        id: r.id, element, item, family: names.family, type: names.type, ref: names.ref || clean(r.userDefinedId),
        quantity: '1', length: isFitting ? 0 : (metres || 0), noLength: !isFitting && metres === null,
        fittingCount: isFitting ? 1 : 0, isFitting
      };
    });
    const missing = rows.filter((r) => r.noLength).length;
    if (lengthField && missing) notes.push(`${fmt.int(missing)} tray or duct items have no ${lengthField.name} value, so they add no length.`);

    // Programme link: task dates for the planned curve
    let programmeLinked = false;
    if (cfg.linkProgramme && rows.length) {
      try {
        const scopeIds = new Set(rows.map((r) => r.id));
        const assignments = await getAll(`${S}/tasks/resource-assignments`, 'resourceAssignments', {}, (n) => progress(`Linking to the programme (${fmt.int(n)} assignments)`));
        const byResource = new Map();
        const taskIds = new Set();
        for (const a of assignments) {
          if (!scopeIds.has(a.resourceId)) continue;
          (byResource.get(a.resourceId) || byResource.set(a.resourceId, []).get(a.resourceId)).push(a.taskId);
          taskIds.add(a.taskId);
        }
        let tasks = [];
        if (taskIds.size && taskIds.size <= 400) {
          tasks = await mapLimit([...taskIds], 6, (id) => api(`${S}/tasks/${encodeURIComponent(id)}`).then((j) => j.task || j).catch(() => null));
        } else if (taskIds.size) {
          tasks = await getAll(`${S}/tasks`, 'tasks', {}, (n) => progress(`Reading programme tasks (${fmt.int(n)})`));
        }
        const taskById = new Map(tasks.filter(Boolean).map((t) => [t.id, t]));
        for (const row of rows) {
          const linked = (byResource.get(row.id) || []).map((id) => taskById.get(id)).filter(Boolean);
          if (!linked.length) continue;
          const starts = linked.map((t) => Date.parse(t.plannedStart)).filter(Number.isFinite);
          const finishes = linked.map((t) => Date.parse(t.plannedFinish)).filter(Number.isFinite);
          if (starts.length) row.plannedStart = new Date(Math.min(...starts)).toISOString();
          if (finishes.length) {
            const last = Math.max(...finishes);
            row.plannedFinish = new Date(last).toISOString();
            const t = linked.find((x) => Date.parse(x.plannedFinish) === last) || linked[0];
            row.task = [clean(t.userDefinedId), clean(t.name)].filter(Boolean).join('  ');
          }
          programmeLinked = true;
        }
        if (!programmeLinked) notes.push('No programme tasks are linked to these items yet.');
      } catch (e) {
        if (e.authRequired) throw e;
        notes.push(`Programme link unavailable: ${e.message}`);
      }
    }

    return {
      schedule: { id: schedule.id || cfg.scheduleId, name: clean(schedule.name) || 'Synchro schedule', type: schedule.type || '' },
      category: category ? { id: category.id, name: clean(category.name) } : null,
      statusItems, rows, programmeLinked, notes, unit,
      mapping: {
        length: lengthField ? `${lengthField.name} (${lengthField.category === 'Entity3d' ? '3D object' : 'resource'} field)` : 'none',
        object: objectSource.kind === 'field' ? `${objectSource.field.name} field` : objectSource.kind === 'parent' ? 'parent resource' : 'element name rules',
        scope: scopeMode
      }
    };
  }

  async function lastSuccessfulSync(S) {
    try {
      const json = await api(`${S}/sync-operations?$top=20`);
      const ok = (json.syncOperations || [])
        .filter((o) => /succeeded|partialsuccess/i.test(o.state || '') && o.completionTime)
        .sort((a, b) => Date.parse(b.completionTime) - Date.parse(a.completionTime))[0];
      return ok ? ok.completionTime : null;
    } catch (e) {
      if (e.authRequired) throw e;
      return null;
    }
  }

  async function loadDataset(opts = {}) {
    const cfg = Object.assign(config(), opts.config || {});
    const progress = opts.onProgress || (() => {});
    if (!cfg.scheduleId) throw new Error('Choose a Synchro schedule in Synchro settings first.');
    const S = scheduleBase(cfg.scheduleId);
    const sig = configSignature(cfg);
    const cached = opts.noCache ? null : store.get(KEYS.staticData, null);
    let base = cached && cached.sig === sig && Date.now() - cached.at < cfg.fullReloadHours * 3600e3 && !opts.force ? cached : null;
    if (!base) {
      base = await loadStatic(cfg, progress);
      base.sig = sig;
      base.at = Date.now();
      if (!opts.noCache && !store.set(KEYS.staticData, base)) store.remove(KEYS.staticData);
    }
    progress('Reading installation statuses');
    const withHistory = opts.history !== false;
    const [hist, lastSync] = await Promise.all([
      base.category ? getAll(`${S}/resources/status-history`, 'statusHistory', Object.assign({ categoryIds: base.category.id }, withHistory ? {} : { isLatest: 'true' }), (n) => progress(`Reading statuses (${fmt.int(n)})`)) : [],
      lastSuccessfulSync(S)
    ]);
    const rowIds = new Set(base.rows.map((r) => r.id));
    if (base.mapping.scope === 'status' && !opts.retried && hist.some((h) => !rowIds.has(h.resourceId))) {
      // New items have been given a status since the last full read: refresh the item list.
      return loadDataset(Object.assign({}, opts, { force: true, retried: true }));
    }
    const itemById = new Map(base.statusItems.map((i) => [i.id, i]));
    const latest = new Map();
    for (const h of hist) {
      if (!rowIds.has(h.resourceId)) continue;
      const prev = latest.get(h.resourceId);
      if (!prev || histCompare(h, prev) > 0) latest.set(h.resourceId, h);
    }
    const unrecorded = clean(cfg.unrecordedStatus) || 'Not Started';
    const unrecordedItem = base.statusItems.find((i) => i.name.toLowerCase() === unrecorded.toLowerCase());
    const rows = base.rows.map((r, index) => {
      const h = latest.get(r.id);
      const item = h && itemById.get(h.statusItemId);
      return Object.assign({}, r, {
        rowNumber: index + 1,
        status: h ? (item ? item.name : 'Unknown status') : unrecorded,
        statusColor: h ? (item ? item.color : null) : (unrecordedItem ? unrecordedItem.color : null),
        statusDate: h ? h.date : null,
        recorded: Boolean(h),
        critical: Boolean(h && h.isCritical),
        note: h ? clean(h.note) : ''
      });
    });
    const statuses = base.statusItems.map((i) => ({ name: i.name, color: i.color }));
    for (const name of sortStatuses(rows.map((r) => r.status))) if (!statuses.some((s) => s.name === name)) statuses.push({ name, color: null });
    const history = withHistory
      ? hist.filter((h) => rowIds.has(h.resourceId)).sort(histCompare)
        .map((h) => ({ id: h.resourceId, date: h.date, status: (itemById.get(h.statusItemId) || {}).name || 'Unknown status', critical: Boolean(h.isCritical) }))
      : [];
    const readAt = new Date().toISOString();
    const dataset = {
      version: 1, source: 'cloud', label: base.schedule.name, scheduleId: base.schedule.id, scheduleType: base.schedule.type,
      projectId: cfg.projectId, projectLabel: cfg.projectLabel, category: base.category ? base.category.name : '',
      fetchedAt: readAt, asAt: readAt, staticAt: new Date(base.at || Date.now()).toISOString(), lastSync,
      unit: 'metres', unitShort: 'm', rows, statuses, history, hasHistory: withHistory,
      programmeLinked: base.programmeLinked, notes: base.notes, mapping: base.mapping
    };
    if (!opts.noCache) cacheDataset(dataset);
    return dataset;
  }
  function cacheDataset(dataset) {
    if (store.set(KEYS.dataset, dataset)) return;
    store.set(KEYS.dataset, Object.assign({}, dataset, { history: [], hasHistory: false, historyTrimmed: true }));
  }
  function cachedDataset() {
    const d = store.get(KEYS.dataset, null);
    const cfg = config();
    return d && d.source === 'cloud' && d.scheduleId === cfg.scheduleId ? d : null;
  }

  /* ------------------------------------------------------------------ snapshots (trend without history) */
  function sourceKey(dataset) {
    if (dataset.source === 'cloud') return `cloud:${dataset.scheduleId}`;
    // One series per export, whatever date or copy number gets added to the file name.
    const name = clean(dataset.filename || dataset.label).replace(/^cached copy:\s*/i, '').replace(/\.[a-z0-9]+$/i, '')
      .toLowerCase().replace(/[^a-z]+/g, ' ').trim();
    return `excel:${name}`;
  }
  const snapshots = {
    record(dataset) {
      if (!dataset || dataset.snapshot || !dataset.rows || !dataset.rows.length) return;
      const day = new Date().toLocaleDateString('en-CA');
      const items = {};
      for (const r of dataset.rows) {
        const byStatus = items[r.item || 'Unspecified'] || (items[r.item || 'Unspecified'] = {});
        const t = byStatus[r.status] || (byStatus[r.status] = { n: 0, len: 0, fit: 0 });
        t.n += 1; t.len += r.length || 0; t.fit += r.fittingCount || 0;
      }
      const src = sourceKey(dataset);
      const all = (store.get(KEYS.snapshots, []) || []).filter((s) => !(s.src === src && s.day === day));
      all.push({ src, day, at: new Date().toISOString(), items });
      store.set(KEYS.snapshots, all.slice(-600));
    },
    list(dataset) {
      if (dataset && Array.isArray(dataset.series)) return dataset.series;
      const src = sourceKey(dataset);
      return (store.get(KEYS.snapshots, []) || []).filter((s) => s.src === src).sort((a, b) => a.day.localeCompare(b.day));
    }
  };

  /* ------------------------------------------------------------------ auto refresh */
  function autoRefresh(fn, onTick) {
    let due = Date.now() + config().refreshMinutes * 60e3;
    let running = false;
    let stopped = false;
    const tick = async () => {
      if (stopped) return;
      if (onTick) onTick(Math.max(0, due - Date.now()), running);
      if (Date.now() >= due && !running && document.visibilityState === 'visible') {
        running = true;
        try { await fn(); } catch (e) { /* the caller reports errors */ }
        running = false;
        due = Date.now() + Math.max(2, Number(config().refreshMinutes) || 10) * 60e3;
        if (onTick) onTick(due - Date.now(), false);
      }
    };
    const timer = setInterval(tick, 15000);
    document.addEventListener('visibilitychange', tick);
    tick();
    return {
      reset() { due = Date.now() + Math.max(2, Number(config().refreshMinutes) || 10) * 60e3; tick(); },
      stop() { stopped = true; clearInterval(timer); document.removeEventListener('visibilitychange', tick); }
    };
  }

  /* ------------------------------------------------------------------ settings dialog */
  const DIALOG_CSS = `
  .sc-dialog{--_ink:var(--sc-ink,#18211f);--_muted:var(--sc-muted,#5b6763);--_accent:var(--sc-accent,#1e3775);--_paper:var(--sc-paper,#f6f7f5);--_line:var(--sc-line,#d9dfda);--_font:var(--sc-font,"Segoe UI",system-ui,sans-serif);--_mono:var(--sc-mono,Consolas,monospace);--_bad:var(--sc-danger,#b3261e);--_ok:var(--sc-ok,#2e7d32);
    width:min(820px,95vw);max-height:92vh;padding:0;border:0;border-radius:4px;color:var(--_ink);background:#fff;font-family:var(--_font);box-shadow:0 24px 80px rgba(10,20,40,.28)}
  .sc-dialog::backdrop{background:rgba(12,20,38,.46)}
  .sc-shell{display:flex;flex-direction:column;max-height:92vh;margin:0}
  .sc-head{display:flex;align-items:flex-start;justify-content:space-between;gap:16px;padding:20px 24px 14px;border-bottom:1px solid var(--_line)}
  .sc-kicker{margin:0 0 4px;font:600 11px var(--_mono);letter-spacing:.08em;text-transform:uppercase;color:var(--_accent)}
  .sc-head h2{margin:0;font-size:20px;font-weight:700;letter-spacing:-.01em}
  .sc-x{border:0;background:none;font-size:26px;line-height:1;color:var(--_muted);cursor:pointer;padding:2px 6px}
  .sc-body{overflow:auto;padding:6px 24px 10px}
  .sc-sec{padding:16px 0;border-bottom:1px solid var(--_line)}
  .sc-sec:last-of-type{border-bottom:0}
  .sc-sec h3{display:flex;align-items:center;gap:10px;margin:0 0 12px;font-size:14px;font-weight:700}
  .sc-sec h3 span{display:inline-grid;place-items:center;width:22px;height:22px;border-radius:50%;background:var(--_accent);color:#fff;font:600 11px var(--_mono)}
  .sc-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px 16px}
  .sc-field{display:flex;flex-direction:column;gap:5px;font-size:12px;color:var(--_muted);min-width:0}
  .sc-field>span{font-weight:600}
  .sc-wide{grid-column:1 / -1}
  .sc-field input,.sc-field select{min-height:36px;border:1px solid var(--_line);border-radius:3px;padding:0 10px;font:13px var(--_font);color:var(--_ink);background:#fff;min-width:0}
  .sc-field input[name=clientId],.sc-field input[name=projectId]{font-family:var(--_mono);font-size:12px}
  .sc-field input:focus,.sc-field select:focus{outline:2px solid color-mix(in srgb,var(--_accent) 40%,transparent);outline-offset:0;border-color:var(--_accent)}
  .sc-check{display:flex;align-items:center;gap:9px;font-size:13px;color:var(--_ink)}
  .sc-check input{width:16px;height:16px;accent-color:var(--_accent)}
  .sc-auth{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin-top:12px}
  .sc-who{flex:1;min-width:200px;font-size:13px}
  .sc-who b{font-weight:700}
  .sc-help{margin:10px 0 0;font-size:12px;color:var(--_muted);line-height:1.5}
  .sc-help code,.sc-warn code{font:12px var(--_mono);background:var(--_paper);padding:1px 5px;border:1px solid var(--_line)}
  .sc-warn{margin:14px 0 2px;padding:12px 14px;background:#fff6e5;border:1px solid #f0d49b;font-size:13px;line-height:1.5}
  .sc-btn{min-height:36px;padding:0 14px;border:1px solid var(--_accent);border-radius:3px;background:var(--_accent);color:#fff;font:600 12px var(--_font);cursor:pointer}
  .sc-btn:hover{filter:brightness(1.12)}
  .sc-btn:disabled{opacity:.45;cursor:not-allowed;filter:none}
  .sc-ghost{background:#fff;color:var(--_accent);border-color:var(--_line)}
  .sc-ghost:hover{border-color:var(--_accent);filter:none}
  .sc-link{border:0;background:none;color:var(--_muted);font:12px var(--_font);text-decoration:underline;cursor:pointer;padding:0}
  .sc-foot{display:flex;flex-wrap:wrap;align-items:center;gap:10px;padding:14px 24px;border-top:1px solid var(--_line);background:var(--_paper)}
  .sc-spacer{flex:1}
  .sc-msg{min-height:18px;margin:8px 0 6px;font-size:12px;color:var(--_muted)}
  .sc-msg.is-bad{color:var(--_bad)} .sc-msg.is-ok{color:var(--_ok)}
  .sc-preview h3{margin-bottom:8px}
  .sc-facts{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin:0 0 12px}
  .sc-facts div{padding:10px 12px;background:var(--_paper);border:1px solid var(--_line)}
  .sc-facts b{display:block;font:600 18px var(--_mono);color:var(--_ink)}
  .sc-facts span{font-size:11px;color:var(--_muted)}
  .sc-chips{display:flex;flex-wrap:wrap;gap:6px;margin:6px 0 10px}
  .sc-chips span{font-size:11px;padding:4px 8px;border:1px solid var(--_line);background:#fff}
  .sc-table{width:100%;border-collapse:collapse;font-size:12px}
  .sc-table th{text-align:left;font:600 10px var(--_mono);text-transform:uppercase;color:var(--_muted);padding:6px;border-bottom:1px solid var(--_line)}
  .sc-table td{padding:6px;border-bottom:1px solid #eef0ee;max-width:320px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .sc-notes{margin:8px 0 0;padding-left:18px;font-size:12px;color:#8a5a00}
  .sc-lead{margin:16px 0 2px;font-size:14px;line-height:1.55}
  .sc-source{padding:11px 13px;background:var(--_paper);border:1px solid var(--_line);font-size:13px;line-height:1.55}
  .sc-files{list-style:none;margin:0 0 10px;padding:0;display:grid;gap:6px}
  .sc-files li{display:flex;justify-content:space-between;gap:12px;padding:9px 12px;border:1px solid var(--_line);background:var(--_paper);font-size:13px}
  .sc-files span{font:12px var(--_mono);color:var(--_muted)}
  .sc-btn[hidden]{display:none}
  @media (max-width:640px){.sc-grid,.sc-facts{grid-template-columns:1fr}.sc-head,.sc-body,.sc-foot{padding-left:16px;padding-right:16px}}`;
  function injectStyles() {
    if (document.getElementById('sc-dialog-css')) return;
    const style = document.createElement('style');
    style.id = 'sc-dialog-css';
    style.textContent = DIALOG_CSS;
    document.head.appendChild(style);
  }
  function option(value, label, selected) {
    const o = document.createElement('option');
    o.value = value;
    o.textContent = label;
    if (selected) o.selected = true;
    return o;
  }
  function download(filename, text, type = 'text/javascript') {
    const link = document.createElement('a');
    link.href = URL.createObjectURL(new Blob([text], { type: `${type};charset=utf-8` }));
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    setTimeout(() => { URL.revokeObjectURL(link.href); link.remove(); }, 500);
  }

  let settingsSavedHandler = null;
  function openSettings(options = {}) {
    injectStyles();
    const onSaved = options.onSaved || settingsSavedHandler;
    document.querySelectorAll('dialog.sc-dialog').forEach((d) => d.remove());
    const cfg = config();
    const dlg = document.createElement('dialog');
    dlg.className = 'sc-dialog';
    dlg.setAttribute('aria-labelledby', 'sc-title');
    dlg.innerHTML = `
    <form method="dialog" class="sc-shell">
      <header class="sc-head">
        <div><p class="sc-kicker">Data source</p><h2 id="sc-title">Synchro cloud connection</h2></div>
        <button type="button" class="sc-x" data-act="close" aria-label="Close">&times;</button>
      </header>
      <div class="sc-body">
        ${IS_HTTP ? '' : '<p class="sc-warn">These dashboards are open straight from a file, so Bentley sign-in cannot return to them. Start them with <code>Start Dashboards.bat</code> to use the live connection. Loading an Excel export still works here.</p>'}
        <section class="sc-sec">
          <h3><span>1</span>Bentley sign-in</h3>
          <div class="sc-grid">
            <label class="sc-field sc-wide"><span>App client ID (from developer.bentley.com, My Apps)</span><input name="clientId" autocomplete="off" spellcheck="false" placeholder="spa-xxxxxxxxxxxxxxxxxxxxxxxx"></label>
          </div>
          <div class="sc-auth"><span class="sc-who" data-who>Not signed in</span>
            <button type="button" class="sc-btn" data-act="signin">Sign in with Bentley</button>
            <button type="button" class="sc-btn sc-ghost" data-act="signout">Sign out</button></div>
          <p class="sc-help">One-time setup: register a single-page app with redirect URI <code data-redirect></code>. The steps are in README.md beside the dashboards.</p>
        </section>
        <section class="sc-sec">
          <h3><span>2</span>Project and schedule</h3>
          <div class="sc-grid">
            <label class="sc-field"><span>Find a project</span><input name="projectSearch" placeholder="Type part of the project name" autocomplete="off"></label>
            <label class="sc-field"><span>Your projects</span><select name="projectPick"><option value="">Sign in to list projects</option></select></label>
            <label class="sc-field sc-wide"><span>Project ID (fills in when you choose a project above)</span><input name="projectId" spellcheck="false" autocomplete="off" placeholder="00000000-0000-0000-0000-000000000000"></label>
            <label class="sc-field sc-wide"><span>Synchro schedule</span><select name="scheduleId"><option value="">Choose a project first</option></select></label>
          </div>
        </section>
        <section class="sc-sec">
          <h3><span>3</span>What to read</h3>
          <div class="sc-grid">
            <label class="sc-field"><span>Installation status category</span><select name="statusCategoryId"><option value="">Detect</option></select></label>
            <label class="sc-field"><span>Status to show when none is recorded</span><input name="unrecordedStatus"></label>
            <label class="sc-field"><span>Length field</span><select name="lengthFieldId"><option value="auto">Detect a field called Length</option><option value="none">No length (count items only)</option></select></label>
            <label class="sc-field"><span>Length unit in Synchro</span><select name="lengthUnit"><option value="auto">Detect</option><option value="m">Metres</option><option value="mm">Millimetres</option><option value="ft">Feet</option></select></label>
            <label class="sc-field sc-wide"><span>Object type (tray, duct, fitting) comes from</span><select name="objectSource"><option value="auto">Detect a Category field, else element-name rules</option><option value="rules">Element name rules (tray, duct and fitting words)</option><option value="parent">Name of the parent resource</option></select></label>
            <label class="sc-field"><span>Items in scope</span><select name="scopeMode">
              <option value="status">Resources with an installation status</option>
              <option value="descendants">Everything inside a resource folder</option>
              <option value="group">A resource group</option>
              <option value="name">Resources whose name matches</option>
              <option value="all">Every resource in the schedule</option></select></label>
            <label class="sc-field" data-scope="descendants group"><span data-scope-label>Folder</span><select name="scopeSelect"></select></label>
            <label class="sc-field" data-scope="name"><span>Name contains (use | for or, or /regex/)</span><input name="scopeText" placeholder="Cable|Trunking|Ladder"></label>
            <label class="sc-check sc-wide"><input type="checkbox" name="linkProgramme"><span>Link to programme task dates (planned curve and behind-programme list)</span></label>
            <label class="sc-field"><span>Refresh statuses every (minutes)</span><input type="number" min="2" max="240" step="1" name="refreshMinutes"></label>
            <label class="sc-field"><span>Project label on the dashboards</span><input name="projectLabel" placeholder="HS2 Euston"></label>
          </div>
        </section>
        <section class="sc-sec sc-preview" data-preview hidden></section>
        <p class="sc-msg" data-msg role="status" aria-live="polite"></p>
      </div>
      <footer class="sc-foot">
        <button type="button" class="sc-link" data-act="reset">Reset to team defaults</button>
        <span class="sc-spacer"></span>
        <button type="button" class="sc-btn sc-ghost" data-act="export">Download team config</button>
        <button type="button" class="sc-btn sc-ghost" data-act="preview">Preview data</button>
        <button type="button" class="sc-btn" data-act="save">Save and load</button>
      </footer>
    </form>`;
    document.body.appendChild(dlg);
    const $ = (sel) => dlg.querySelector(sel);
    const f = (name) => dlg.querySelector(`[name="${name}"]`);
    const msgEl = $('[data-msg]');
    const say = (text, kind) => { msgEl.textContent = text || ''; msgEl.className = `sc-msg${kind ? ` is-${kind}` : ''}`; };
    const state = { folders: null, groups: null, busy: false };
    $('[data-redirect]').textContent = redirectUri();

    f('clientId').value = cfg.clientId;
    f('projectId').value = cfg.projectId;
    f('unrecordedStatus').value = cfg.unrecordedStatus;
    f('lengthUnit').value = cfg.lengthUnit || 'auto';
    f('scopeMode').value = cfg.scopeMode === 'auto' ? 'status' : cfg.scopeMode;
    f('scopeText').value = cfg.scopeMode === 'name' ? cfg.scopeValue : '';
    f('linkProgramme').checked = cfg.linkProgramme !== false;
    f('refreshMinutes').value = cfg.refreshMinutes;
    f('projectLabel').value = cfg.projectLabel || '';
    const keepOption = (select, value, label) => { if (value && ![...select.options].some((o) => o.value === value)) select.append(option(value, label, true)); select.value = value || select.value; };
    keepOption(f('scheduleId'), cfg.scheduleId, `Current schedule (${cfg.scheduleId})`);
    keepOption(f('statusCategoryId'), cfg.statusCategoryId, 'Current category');
    keepOption(f('lengthFieldId'), cfg.lengthFieldId, 'Current length field');
    keepOption(f('objectSource'), cfg.objectSource, 'Current object field');

    function collect() {
      const scopeMode = f('scopeMode').value;
      return {
        clientId: clean(f('clientId').value), projectId: clean(f('projectId').value), scheduleId: f('scheduleId').value,
        statusCategoryId: f('statusCategoryId').value, unrecordedStatus: clean(f('unrecordedStatus').value) || 'Not Started',
        lengthFieldId: f('lengthFieldId').value || 'auto', lengthUnit: f('lengthUnit').value || 'auto', objectSource: f('objectSource').value || 'auto',
        scopeMode, scopeValue: scopeMode === 'name' ? clean(f('scopeText').value) : (['descendants', 'group'].includes(scopeMode) ? f('scopeSelect').value : ''),
        linkProgramme: f('linkProgramme').checked, refreshMinutes: Math.max(2, Number(f('refreshMinutes').value) || 10),
        projectLabel: clean(f('projectLabel').value)
      };
    }
    function syncScopeUI() {
      const m = f('scopeMode').value;
      dlg.querySelectorAll('[data-scope]').forEach((elx) => { elx.hidden = !elx.dataset.scope.split(' ').includes(m); });
      $('[data-scope-label]').textContent = m === 'group' ? 'Resource group' : 'Resource folder';
    }
    async function fillScope() {
      const m = f('scopeMode').value;
      if (m === 'group') await fillGroups(); else if (m === 'descendants') await fillFolders();
    }
    async function guard(task) {
      if (state.busy) return;
      state.busy = true;
      dlg.querySelectorAll('.sc-foot button, [data-act]').forEach((b) => { b.disabled = true; });
      try { await task(); } catch (e) {
        say(e.message, 'bad');
        if (e.authRequired) renderWho();
      } finally {
        state.busy = false;
        dlg.querySelectorAll('.sc-foot button, [data-act]').forEach((b) => { b.disabled = false; });
        renderWho();
      }
    }
    async function renderWho() {
      const who = $('[data-who]');
      const signedIn = isSignedIn();
      $('[data-act="signout"]').hidden = !signedIn;
      $('[data-act="signin"]').hidden = signedIn;
      $('[data-act="signin"]').disabled = !IS_HTTP;
      if (!signedIn) { who.textContent = IS_HTTP ? 'Not signed in.' : 'Sign-in needs the launcher (see above).'; return; }
      who.innerHTML = 'Signed in.';
      try {
        const me = await whoAmI();
        who.innerHTML = `Signed in as <b>${escapeHtml(me.displayName || me.email || 'Bentley user')}</b>${me.organizationName ? `, ${escapeHtml(me.organizationName)}` : ''}`;
      } catch (e) { /* keep the short label */ }
    }
    async function fillProjects(search) {
      const sel = f('projectPick');
      const projects = await listProjects(search);
      const current = clean(f('projectId').value);
      sel.innerHTML = '';
      sel.append(option('', projects.length ? 'Choose a project' : 'No projects found'));
      for (const p of projects) sel.append(option(p.id, `${p.displayName}${p.number ? `  (${p.number})` : ''}`, p.id === current));
    }
    async function fillSchedules() {
      const projectId = clean(f('projectId').value);
      const sel = f('scheduleId');
      if (!projectId) return;
      const schedules = await listSchedules(projectId);
      const current = sel.value || cfg.scheduleId;
      sel.innerHTML = '';
      if (!schedules.length) { sel.append(option('', 'No Synchro schedules in this project')); return; }
      const preferred = schedules.find((s) => s.id === current) || schedules.find((s) => s.type === 'v6.5') || schedules[0];
      for (const s of schedules) sel.append(option(s.id, `${s.name}  (${s.type === 'v6.5' ? 'SYNCHRO 4D Pro' : s.type})`, s.id === preferred.id));
      await fillScheduleOptions();
    }
    async function fillScheduleOptions() {
      const sid = f('scheduleId').value;
      if (!sid) return;
      const S = scheduleBase(sid);
      say('Reading status categories and user fields...');
      const [categories, fields] = await Promise.all([
        getAll(`${S}/resource-status-categories`, 'resourceStatusCategories'), getAll(`${S}/user-fields`, 'userFields')
      ]);
      const cat = f('statusCategoryId');
      const wantCat = cat.value || cfg.statusCategoryId;
      cat.innerHTML = '';
      cat.append(option('', categories.length ? 'Detect (a category named Installation)' : 'No status categories in this schedule'));
      for (const c of categories) cat.append(option(c.id, c.name, c.id === wantCat));
      const lengthSel = f('lengthFieldId');
      const wantLen = lengthSel.value || cfg.lengthFieldId;
      lengthSel.innerHTML = '';
      lengthSel.append(option('auto', 'Detect a field called Length'));
      lengthSel.append(option('none', 'No length (count items only)'));
      const label = (x) => `${x.name}  (${x.category === 'Entity3d' ? '3D object' : x.category.toLowerCase()} field, ${x.type.toLowerCase()})`;
      for (const x of fields.filter((u) => ['Entity3d', 'Resource'].includes(u.category) && ['Float', 'Integer', 'String'].includes(u.type))) lengthSel.append(option(x.id, label(x)));
      lengthSel.value = [...lengthSel.options].some((o) => o.value === wantLen) ? wantLen : 'auto';
      const objSel = f('objectSource');
      const wantObj = objSel.value || cfg.objectSource;
      objSel.innerHTML = '';
      objSel.append(option('auto', 'Detect a Category field, else element-name rules'));
      objSel.append(option('rules', 'Element name rules (tray, duct and fitting words)'));
      objSel.append(option('parent', 'Name of the parent resource'));
      for (const x of fields.filter((u) => ['Entity3d', 'Resource'].includes(u.category) && u.type === 'String')) objSel.append(option(`field:${x.id}`, `User field: ${label(x)}`));
      objSel.value = [...objSel.options].some((o) => o.value === wantObj) ? wantObj : 'auto';
      state.groups = null;
      state.folders = null;
      say(`${categories.length} status categor${categories.length === 1 ? 'y' : 'ies'} and ${fields.length} user field${fields.length === 1 ? '' : 's'} found.`, 'ok');
      syncScopeUI();
      await fillScope();
    }
    async function fillGroups() {
      const sel = f('scopeSelect');
      if (!isSignedIn() || !f('scheduleId').value) { sel.innerHTML = ''; sel.append(option(cfg.scopeMode === 'group' ? cfg.scopeValue : '', 'Sign in to list resource groups')); return; }
      if (!state.groups) state.groups = await getAll(`${scheduleBase(f('scheduleId').value)}/resource-groups`, 'resourceGroups');
      const want = cfg.scopeMode === 'group' ? cfg.scopeValue : sel.value;
      sel.innerHTML = '';
      sel.append(option('', state.groups.length ? 'Choose a group' : 'No resource groups in this schedule'));
      for (const g of state.groups.sort((a, b) => clean(a.name).localeCompare(clean(b.name)))) sel.append(option(g.id, g.name, g.id === want));
    }
    async function fillFolders() {
      const sel = f('scopeSelect');
      if (!isSignedIn() || !f('scheduleId').value) { sel.innerHTML = ''; sel.append(option(cfg.scopeMode === 'descendants' ? cfg.scopeValue : '', 'Sign in to list folders')); return; }
      if (!state.folders) {
        say('Reading the resource tree...');
        const all = await getAll(`${scheduleBase(f('scheduleId').value)}/resources`, 'resources', {}, (n) => say(`Reading the resource tree (${fmt.int(n)})...`));
        const children = new Map();
        for (const r of all) if (r.parentId) (children.get(r.parentId) || children.set(r.parentId, []).get(r.parentId)).push(r);
        const leafCount = new Map();
        const count = (id) => {
          if (leafCount.has(id)) return leafCount.get(id);
          let n = 0;
          for (const c of children.get(id) || []) n += children.has(c.id) ? count(c.id) : 1;
          leafCount.set(id, n);
          return n;
        };
        const byId = new Map(all.map((r) => [r.id, r]));
        const depth = (r) => { let d = 0; let p = r; while (p && p.parentId && d < 30) { p = byId.get(p.parentId); d += 1; } return d; };
        state.folders = all.filter((r) => children.has(r.id)).map((r) => ({ id: r.id, name: clean(r.name), leaves: count(r.id), depth: depth(r) }))
          .sort((a, b) => a.depth - b.depth || b.leaves - a.leaves).slice(0, 400);
        say(`${fmt.int(all.length)} resources read.`, 'ok');
      }
      const want = cfg.scopeMode === 'descendants' ? cfg.scopeValue : sel.value;
      sel.innerHTML = '';
      sel.append(option('', state.folders.length ? 'Choose a folder' : 'This schedule has no resource folders'));
      for (const x of state.folders) sel.append(option(x.id, `${'  '.repeat(Math.min(x.depth, 6))}${x.name}  (${fmt.int(x.leaves)} items)`, x.id === want));
    }
    function renderPreview(d) {
      const box = $('[data-preview]');
      const len = d.rows.reduce((s, r) => s + (r.length || 0), 0);
      const fittings = d.rows.reduce((s, r) => s + (r.fittingCount || 0), 0);
      const count = (key) => { const m = new Map(); for (const r of d.rows) m.set(r[key], (m.get(r[key]) || 0) + 1); return [...m.entries()].sort((a, b) => b[1] - a[1]); };
      box.hidden = false;
      box.innerHTML = `<h3><span>&#10003;</span>Preview: ${escapeHtml(d.label)}</h3>
        <div class="sc-facts"><div><b>${fmt.int(d.rows.length)}</b><span>items in scope</span></div><div><b>${fmt.num(len, 1)} m</b><span>tray and duct length</span></div><div><b>${fmt.int(fittings)}</b><span>fittings</span></div></div>
        <div class="sc-chips">${count('status').map(([k, v]) => `<span>${escapeHtml(k)}: ${fmt.int(v)}</span>`).join('')}</div>
        <div class="sc-chips">${count('item').slice(0, 12).map(([k, v]) => `<span>${escapeHtml(k)}: ${fmt.int(v)}</span>`).join('')}</div>
        <table class="sc-table"><thead><tr><th>Element</th><th>Object</th><th>Length (m)</th><th>Status</th></tr></thead><tbody>
        ${d.rows.slice(0, 6).map((r) => `<tr><td title="${escapeHtml(r.element)}">${escapeHtml(r.element)}</td><td>${escapeHtml(r.item)}</td><td>${r.isFitting ? 'fitting' : fmt.num(r.length, 2)}</td><td>${escapeHtml(r.status)}</td></tr>`).join('')}
        </tbody></table>
        <p class="sc-help">Length from: ${escapeHtml(d.mapping.length)}. Object type from: ${escapeHtml(d.mapping.object)}. Status category: ${escapeHtml(d.category || 'none')}. Programme: ${d.programmeLinked ? 'linked' : 'not linked'}.${d.lastSync ? ` Last cloud sync ${escapeHtml(fmt.dateTime(d.lastSync))}.` : ''}</p>
        ${d.notes.length ? `<ul class="sc-notes">${d.notes.map((n) => `<li>${escapeHtml(n)}</li>`).join('')}</ul>` : ''}`;
    }
    function teamConfigText(values) {
      const keys = ['clientId', 'webUrl', 'projectId', 'scheduleId', 'projectLabel', 'statusCategoryId', 'lengthFieldId', 'lengthUnit', 'objectSource', 'scopeMode', 'scopeValue', 'unrecordedStatus', 'linkProgramme', 'refreshMinutes', 'fullReloadHours'];
      const merged = Object.assign({}, config(), values);
      const body = keys.map((k) => `  ${k}: ${JSON.stringify(merged[k])}`).join(',\n');
      return `/*\n * MEP dashboards: team defaults. Saved from Synchro settings on ${fmt.dateTime(new Date())}.\n * Save this file as assets/dashboard-config.js beside the dashboards so everyone\n * who opens them from that folder starts connected.\n */\nwindow.MEP_DASHBOARD_CONFIG = {\n${body}\n};\n`;
    }

    dlg.addEventListener('click', (ev) => {
      const act = ev.target.closest('[data-act]');
      if (!act) return;
      const which = act.dataset.act;
      if (which === 'close') {
        dlg.close('cancel');
      } else if (which === 'signin') {
        const clientId = clean(f('clientId').value);
        if (!clientId) { say('Paste the app client ID first.', 'bad'); f('clientId').focus(); return; }
        saveUserConfig({ clientId, projectId: clean(f('projectId').value) });
        session.set(KEYS.reopen, '1');
        signIn(location.href).catch((e) => say(e.message, 'bad'));
      } else if (which === 'signout') {
        signOut();
        renderWho();
        say('Signed out of the dashboards.', 'ok');
      } else if (which === 'preview') {
        guard(async () => {
          const values = collect();
          if (!values.scheduleId) throw new Error('Choose a schedule first.');
          saveUserConfig({ clientId: values.clientId });
          const d = await loadDataset({ config: values, noCache: true, history: false, onProgress: (t) => say(`${t}...`) });
          renderPreview(d);
          say(`Preview ready: ${fmt.int(d.rows.length)} items. Save to use these settings.`, 'ok');
        });
      } else if (which === 'save') {
        const values = collect();
        saveUserConfig(values);
        dlg.close('saved');
        if (onSaved) onSaved(values);
      } else if (which === 'export') {
        download('dashboard-config.js', teamConfigText(collect()));
        say('Saved dashboard-config.js to Downloads. Copy it into the assets folder (replace the old one).', 'ok');
      } else if (which === 'reset') {
        store.remove(KEYS.user);
        dlg.close('reset');
        openSettings(options);
      }
    });
    f('projectPick').addEventListener('change', () => {
      if (!f('projectPick').value) return;
      f('projectId').value = f('projectPick').value;
      guard(fillSchedules);
    });
    f('projectId').addEventListener('change', () => guard(fillSchedules));
    f('scheduleId').addEventListener('change', () => guard(fillScheduleOptions));
    f('scopeMode').addEventListener('change', () => { syncScopeUI(); guard(fillScope); });
    $('form').addEventListener('submit', (ev) => ev.preventDefault());
    let searchTimer;
    f('projectSearch').addEventListener('input', () => {
      clearTimeout(searchTimer);
      searchTimer = setTimeout(() => guard(() => fillProjects(clean(f('projectSearch').value))), 350);
    });
    dlg.addEventListener('close', () => setTimeout(() => dlg.remove(), 0));
    dlg.showModal();
    syncScopeUI();
    renderWho();
    if (IS_HTTP && config().clientId) {
      say('Checking Bentley sign-in...');
      guard(async () => {
        try { await getToken(); } catch (e) { say('Sign in with Bentley to choose the project and schedule.'); return; }
        await renderWho();
        await fillProjects('');
        if (clean(f('projectId').value)) await fillSchedules();
        say('Connected. Preview the data or save.', 'ok');
      });
    } else if (!IS_HTTP) {
      say('Live connection needs the launcher. You can still fill in the IDs and download a team config.');
    } else {
      say('Paste the Bentley app client ID, then sign in.');
    }
    return dlg;
  }

  /* ------------------------------------------------------------------ published copies */
  // A published copy is one dashboard page with its data, the shared code and the publisher's
  // choices sealed into a single HTML file. It opens anywhere (a SharePoint download, Teams,
  // email) with no Synchro licence, no sign-in and no network, and stays exactly as published.
  const PUBLISH_FILES = [
    { key: 'progress', source: '01 MEP Installation Progress.html', target: 'MEP Installation Progress.html' },
    { key: 'insights', source: '02 Containment Installation Insights.html', target: 'Containment Installation Insights.html' }
  ];
  const SNAPSHOT_FORMAT = 'mep-dashboard-snapshot';
  const TARGET_KEY = 'mep.insights.target';
  const SCRIPT_END = '</' + 'script>';
  const COMMENT_OPEN = '<' + '!--';
  const COMMENT_CLOSE = '--' + '>';
  const toMinutes = (iso) => { const t = Date.parse(iso); return Number.isFinite(t) ? Math.round(t / 60000) : 0; };
  const fromMinutes = (m) => (m ? new Date(m * 60000).toISOString() : null);
  const round3 = (v) => Math.round((Number(v) || 0) * 1000) / 1000;
  const httpsOnly = (url) => (/^https:\/\/[^\s"'<>]+$/i.test(clean(url)) ? clean(url) : '');

  // Rows and history travel as short arrays with lookup lists, which keeps a copy of a few
  // thousand elements and their whole status history well under a megabyte.
  function packSnapshot(d, meta) {
    const lists = { status: [], item: [], task: [] };
    const seen = { status: new Map(), item: new Map(), task: new Map() };
    const ix = (kind, value) => {
      const key = clean(value);
      if (!seen[kind].has(key)) { seen[kind].set(key, lists[kind].length); lists[kind].push(key); }
      return seen[kind].get(key);
    };
    const colours = {};
    for (const s of d.statuses || []) { ix('status', s.name); if (s.color) colours[clean(s.name)] = s.color; }
    const rowIndex = new Map();
    const rows = d.rows.map((r, i) => {
      rowIndex.set(r.id, i);
      if (r.statusColor && !colours[clean(r.status)]) colours[clean(r.status)] = r.statusColor;
      const flags = (r.recorded ? 1 : 0) | (r.critical ? 2 : 0) | (r.noLength ? 4 : 0) | (r.isFitting ? 8 : 0);
      const qty = clean(r.quantity);
      return [r.element, ix('item', r.item), r.isFitting ? 0 : round3(r.length), ix('status', r.status), toMinutes(r.statusDate), flags,
        clean(r.note), toMinutes(r.plannedStart), toMinutes(r.plannedFinish), r.task ? ix('task', r.task) : -1,
        qty === '1' ? 0 : qty, r.rowNumber && r.rowNumber !== i + 1 ? r.rowNumber : 0];
    });
    const history = (d.history || []).filter((h) => rowIndex.has(h.id))
      .map((h) => [rowIndex.get(h.id), toMinutes(h.date), ix('status', h.status), h.critical ? 1 : 0]);
    const dataset = {};
    for (const key of ['source', 'label', 'filename', 'sheetName', 'scheduleType', 'projectLabel', 'category', 'fetchedAt', 'modifiedAt', 'asAt',
      'lastSync', 'unit', 'unitShort', 'hasHistory', 'programmeLinked', 'notes', 'mapping', 'series']) {
      if (d[key] !== undefined && d[key] !== null) dataset[key] = d[key];
    }
    return { format: SNAPSHOT_FORMAT, v: 1, meta, dataset, statuses: lists.status, colours, items: lists.item, tasks: lists.task, rows, history };
  }
  function unpackSnapshot(p) {
    if (!p || p.format !== SNAPSHOT_FORMAT || !Array.isArray(p.rows)) throw new Error('this file does not contain readable dashboard data');
    const meta = p.meta || {};
    const st = p.statuses || [];
    const items = p.items || [];
    const tasks = p.tasks || [];
    const colours = p.colours || {};
    const d = Object.assign({ version: 1, notes: [], history: [] }, p.dataset);
    d.rows = p.rows.map((a, i) => {
      const [element, item, length, status, statusMin, flags, note, startMin, finishMin, task, qty, rowNumber] = a;
      const names = splitName(element);
      const isFitting = Boolean(flags & 8);
      const name = st[status] || 'Unspecified';
      const row = {
        id: `s${i}`, rowNumber: rowNumber || i + 1, element, item: items[item] || 'Unspecified', family: names.family, type: names.type, ref: names.ref,
        quantity: qty === 0 ? '1' : clean(qty), length: isFitting ? 0 : Number(length) || 0, noLength: Boolean(flags & 4),
        fittingCount: isFitting ? 1 : 0, isFitting, status: name, statusColor: colours[name] || null, statusDate: fromMinutes(statusMin),
        recorded: Boolean(flags & 1), critical: Boolean(flags & 2), note: note || ''
      };
      if (startMin) row.plannedStart = fromMinutes(startMin);
      if (finishMin) row.plannedFinish = fromMinutes(finishMin);
      if (task >= 0 && tasks[task]) row.task = tasks[task];
      return row;
    });
    d.history = (p.history || []).map(([row, min, status, critical]) => ({ id: `s${row}`, date: fromMinutes(min), status: st[status] || 'Unknown status', critical: Boolean(critical) }));
    d.statuses = st.map((name) => ({ name, color: colours[name] || null }));
    d.asAt = meta.asAt || d.asAt || d.fetchedAt;
    d.snapshot = meta;
    return { dataset: d, meta };
  }

  let snapshotCache;
  function readSnapshot() {
    if (snapshotCache !== undefined) return snapshotCache;
    const el = document.getElementById('mep-snapshot');
    if (!el) { snapshotCache = null; return null; }
    try {
      snapshotCache = unpackSnapshot(JSON.parse(el.textContent));
    } catch (e) {
      snapshotCache = { error: e.message || 'unreadable data', meta: null, dataset: null };
    }
    return snapshotCache;
  }
  function describeSnapshot(snap) {
    const meta = (snap && snap.meta) || {};
    const d = (snap && snap.dataset) || {};
    const asAt = meta.asAt || d.asAt;
    const parsed = Date.parse(asAt);
    const ageDays = Number.isFinite(parsed) ? Math.max(0, (Date.now() - parsed) / 864e5) : 0;
    const stale = ageDays > (Number(meta.staleAfterDays) || 3);
    const from = d.source === 'cloud' ? 'Synchro' : `the Excel export${d.filename ? ` ${d.filename}` : ''}`;
    let html = `Published copy: data as at <b>${escapeHtml(fmt.dateTime(asAt))}</b> from ${escapeHtml(from)}`;
    if (meta.publishedBy) html += `, published by ${escapeHtml(meta.publishedBy)}`;
    html += '. It stays as published until the next update.';
    const days = Math.floor(ageDays);
    const link = httpsOnly(meta.latestUrl);
    const staleHtml = stale
      ? `This copy is ${days} day${days === 1 ? '' : 's'} old. ${link ? `<a href="${escapeHtml(link)}" target="_blank" rel="noopener noreferrer">Open the latest on SharePoint</a>.` : 'The latest version is on SharePoint.'}`
      : '';
    return { html, staleHtml, stale, ageDays, asAt };
  }
  // Key figures as plain HTML, for viewers that do not run scripts (some phone previews).
  function noscriptSummary(d, completeSet, meta) {
    let len = 0; let fit = 0; let doneLen = 0; let doneFit = 0;
    const byStatus = new Map();
    for (const r of d.rows) {
      len += r.length || 0;
      fit += r.fittingCount || 0;
      if (completeSet.has(r.status)) { doneLen += r.length || 0; doneFit += r.fittingCount || 0; }
      byStatus.set(r.status, (byStatus.get(r.status) || 0) + 1);
    }
    const total = len + fit;
    const done = doneLen + doneFit;
    const pct = total ? (done / total) * 100 : 0;
    const lines = sortStatuses([...byStatus.keys()]).map((s) => `<li>${escapeHtml(s)}: ${fmt.int(byStatus.get(s))} items</li>`).join('');
    return `<noscript><style>body > :not(noscript) { display: none !important; }</style><div style="max-width:760px;margin:24px auto;padding:20px 24px;background:#fff;border:1px solid #ccc;font:15px/1.5 'Segoe UI',Arial,sans-serif;color:#222">
<h2 style="margin:0 0 6px;font-size:20px">${escapeHtml(meta.projectLabel || 'MEP')} cable containment: published copy</h2>
<p style="margin:0 0 10px">Data as at ${escapeHtml(fmt.dateTime(meta.asAt))}${meta.publishedBy ? `, published by ${escapeHtml(meta.publishedBy)}` : ''}.</p>
<p style="margin:0 0 10px"><b>${fmt.num(pct, 1)}%</b> of the weighted scope is installed: ${fmt.num(done, 1)} of ${fmt.num(total, 1)} points (${fmt.num(doneLen, 1)} m of ${fmt.num(len, 1)} m, ${fmt.int(doneFit)} of ${fmt.int(fit)} fittings).</p>
<ul style="margin:0 0 10px">${lines}</ul>
<p style="margin:0;color:#555">Open this file in Edge or Chrome for the full interactive dashboard.</p></div></noscript>`;
  }
  function buildPublishedHtml(source, core, payload, meta, noscript, note) {
    let html = String(source);
    const coreTag = /<script src="assets\/synchro-cloud\.js"><\/script>/i;
    if (!coreTag.test(html)) throw new Error('A dashboard page has changed shape and cannot be published (it no longer loads assets/synchro-cloud.js).');
    html = html.replace(/[ \t]*<script[^>]*\bsrc="https:\/\/cdn\.jsdelivr\.net\/npm\/xlsx[^"]*"[^>]*><\/script>[ \t]*\r?\n?/i, () => '');
    html = html.replace(/[ \t]*<script src="assets\/dashboard-config\.js"><\/script>[ \t]*\r?\n?/i, () => '');
    for (const file of PUBLISH_FILES) html = html.split(encodeURI(file.source)).join(encodeURI(file.target));
    html = html.replace(/<title>([\s\S]*?)<\/title>/i, (m, t) => `<title>${t} · ${escapeHtml(fmt.date(meta.asAt))}</title>`);
    html = html.replace(/<!doctype html>/i, (m) => `${m}\n${COMMENT_OPEN} ${note.replace(/-{2,}/g, '-')} ${COMMENT_CLOSE}`);
    html = html.replace(/<body([^>]*)>/i, (m) => `${m}\n${noscript}`);
    const json = JSON.stringify(payload).replace(/</g, '\\u003c');
    const inlineCore = String(core).replace(/<\/script/gi, '<\\/script');
    html = html.replace(coreTag, () => `<script type="application/json" id="mep-snapshot">${json}${SCRIPT_END}\n  <script>\n${inlineCore}\n  ${SCRIPT_END}`);
    return html;
  }

  // The page templates come from the launcher, or from assets/publish-kit.js when the
  // dashboards are opened straight from a file.
  function loadKit() {
    if (global.MEP_PUBLISH_KIT) return Promise.resolve(global.MEP_PUBLISH_KIT);
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = new URL('assets/publish-kit.js', ROOT).href;
      s.onload = () => (global.MEP_PUBLISH_KIT ? resolve(global.MEP_PUBLISH_KIT) : reject(new Error('assets/publish-kit.js is empty.')));
      s.onerror = () => reject(new Error('Could not load assets/publish-kit.js, which publishing needs when the dashboards are opened from a file.'));
      document.head.appendChild(s);
    });
  }
  async function publishSources() {
    if (IS_HTTP) {
      const get = async (path) => {
        const res = await fetch(new URL(path, ROOT), { cache: 'no-store' });
        if (!res.ok) throw new Error(`Could not read ${decodeURI(path)} (${res.status}).`);
        return res.text();
      };
      const [core, ...pages] = await Promise.all([get('assets/synchro-cloud.js'), ...PUBLISH_FILES.map((file) => get(encodeURI(file.source)))]);
      return { core, pages };
    }
    const kit = await loadKit();
    if (kit.version !== VERSION) throw new Error('assets/publish-kit.js is out of date. Publish through Start Dashboards.bat, or rebuild the kit (see README).');
    return { core: kit.core, pages: PUBLISH_FILES.map((file) => kit.pages[file.key]) };
  }

  const idb = {
    db: null,
    open() {
      if (this.db) return Promise.resolve(this.db);
      return new Promise((resolve, reject) => {
        const req = indexedDB.open('mep-dashboards', 1);
        req.onupgradeneeded = () => req.result.createObjectStore('kv');
        req.onsuccess = () => { this.db = req.result; resolve(req.result); };
        req.onerror = () => reject(req.error);
      });
    },
    async get(key) {
      const db = await this.open();
      return new Promise((resolve, reject) => {
        const req = db.transaction('kv').objectStore('kv').get(key);
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
    },
    async set(key, value) {
      const db = await this.open();
      return new Promise((resolve, reject) => {
        const tx = db.transaction('kv', 'readwrite');
        tx.objectStore('kv').put(value, key);
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      });
    }
  };
  async function dirHasFile(dir, name) {
    try { await dir.getFileHandle(name); return true; } catch (e) { return false; }
  }
  async function ensureWritable(handle) {
    const access = { mode: 'readwrite' };
    if ((await handle.queryPermission(access)) === 'granted') return true;
    return (await handle.requestPermission(access)) === 'granted';
  }
  async function writeFiles(dir, files) {
    if (await dirHasFile(dir, 'Start Dashboards.bat')) throw new Error("That is the dashboards' own folder. Choose the SharePoint folder instead.");
    for (const file of files) {
      const handle = await dir.getFileHandle(file.name, { create: true });
      const writable = await handle.createWritable();
      await writable.write(new Blob([file.text], { type: 'text/html' }));
      await writable.close();
    }
  }
  function downloadFiles(files) {
    files.forEach((file, i) => setTimeout(() => download(file.name, file.text, 'text/html'), i * 700));
  }
  function lastPublish() { return store.get(KEYS.publish, null); }

  async function publishNow(opts = {}) {
    if (readSnapshot()) throw new Error('This is already a published copy. Publish from the dashboards folder instead.');
    const progress = opts.onProgress || (() => {});
    const cfg = config();
    let d = typeof opts.dataset === 'function' ? opts.dataset() : opts.dataset;
    if (mode() === 'cloud' && (!d || d.source === 'cloud')) {
      progress('Reading the latest data from Synchro');
      d = await loadDataset({ history: true, onProgress: progress });
    }
    if (!d) {
      const cache = cachedWorkbook();
      if (cache) d = datasetFromCache(cache);
    }
    if (!d || !d.rows || !d.rows.length) throw new Error('There is no data to publish yet. Load an Excel export or connect to Synchro first.');
    if (d.source === 'excel') snapshots.record(d);
    const labels = sortStatuses(d.rows.map((r) => r.status));
    const completeSet = completion.get(labels);
    let target = '';
    try { target = localStorage.getItem(TARGET_KEY) || ''; } catch (e) { /* ignore */ }
    const meta = {
      asAt: d.asAt || d.fetchedAt, publishedAt: new Date().toISOString(), publishedBy: clean(opts.by), latestUrl: httpsOnly(opts.latestUrl),
      staleAfterDays: Number(cfg.staleAfterDays) || 3, complete: [...completeSet], target: /^\d{4}-\d{2}-\d{2}$/.test(target) ? target : '',
      projectLabel: clean(cfg.projectLabel || d.projectLabel), generator: `MEP dashboards ${VERSION}`
    };
    progress('Packing the data');
    const payload = packSnapshot(Object.assign({}, d, { series: d.source === 'excel' ? snapshots.list(d) : undefined }), meta);
    progress('Building the dashboard files');
    const sources = await publishSources();
    const noscript = noscriptSummary(d, completeSet, meta);
    const note = `Published copy of the MEP installation dashboards. Data as at ${fmt.dateTime(meta.asAt)}; published ${fmt.dateTime(meta.publishedAt)}${meta.publishedBy ? ` by ${meta.publishedBy}` : ''}. Made with the Publish for SharePoint button: publish again rather than editing this file.`;
    const files = PUBLISH_FILES.map((file, i) => ({ name: file.target, text: buildPublishedHtml(sources.pages[i], sources.core, payload, meta, noscript, note) }));
    if (opts.dirHandle) {
      progress('Saving the files');
      await writeFiles(opts.dirHandle, files);
    } else if (opts.download !== false) {
      downloadFiles(files);
    }
    const result = {
      at: meta.publishedAt, asAt: meta.asAt, by: meta.publishedBy, items: d.rows.length, folder: opts.dirHandle ? opts.dirHandle.name : null,
      files: files.map((x) => ({ name: x.name, bytes: new Blob([x.text]).size }))
    };
    if (opts.record !== false) store.set(KEYS.publish, Object.assign(store.get(KEYS.publish, {}) || {}, { last: result }));
    return Object.assign({ dataset: d, meta, html: opts.returnHtml ? files : undefined }, result);
  }

  let publishHooks = null;
  function openPublish(options = {}) {
    if (readSnapshot()) return null;
    injectStyles();
    const hooks = Object.assign({}, publishHooks || {}, options);
    document.querySelectorAll('dialog.sc-dialog').forEach((x) => x.remove());
    const cfg = config();
    const prefs = store.get(KEYS.publish, {}) || {};
    const canPick = typeof global.showDirectoryPicker === 'function';
    const getDataset = () => (typeof hooks.dataset === 'function' ? hooks.dataset() : hooks.dataset) || null;
    const dlg = document.createElement('dialog');
    dlg.className = 'sc-dialog';
    dlg.setAttribute('aria-labelledby', 'sc-pub-title');
    dlg.innerHTML = `
    <form method="dialog" class="sc-shell">
      <header class="sc-head">
        <div><p class="sc-kicker">Leadership copies</p><h2 id="sc-pub-title">Publish for SharePoint</h2></div>
        <button type="button" class="sc-x" data-act="close" aria-label="Close">&times;</button>
      </header>
      <div class="sc-body">
        <p class="sc-lead">Saves both dashboards as files with the data sealed inside. Anyone can open them from SharePoint, Teams or email, with no Synchro licence and no sign-in. They show the data as it is now and stay that way until the next publish.</p>
        <section class="sc-sec">
          <h3><span>1</span>Data</h3>
          <div class="sc-source" data-source></div>
        </section>
        <section class="sc-sec">
          <h3><span>2</span>Shown on the copies</h3>
          <div class="sc-grid">
            <label class="sc-field"><span>Published by</span><input name="by" autocomplete="name" placeholder="Your name"></label>
            <label class="sc-field"><span>SharePoint folder link (optional)</span><input name="latestUrl" spellcheck="false" autocomplete="off" placeholder="https://...sharepoint.com/..."></label>
          </div>
          <p class="sc-help">When a copy is more than ${Number(cfg.staleAfterDays) || 3} days old it says so, and points to this link for the latest version.</p>
        </section>
        <section class="sc-sec">
          <h3><span>3</span>Save to</h3>
          ${canPick ? `<div class="sc-auth"><span class="sc-who" data-folder>No folder chosen yet.</span>
            <button type="button" class="sc-btn sc-ghost" data-act="pick">Choose folder</button></div>
          <p class="sc-help">Choose the SharePoint folder as synced to this PC by OneDrive (in SharePoint, use Sync or Add shortcut to My files). Each publish replaces the two files there, and SharePoint keeps the earlier versions.</p>`
          : '<p class="sc-help">This browser will download the two files. Upload both to the SharePoint folder, replacing the old copies.</p>'}
          <p class="sc-help" data-last></p>
        </section>
        <section class="sc-sec" data-result hidden></section>
        <p class="sc-msg" data-msg role="status" aria-live="polite"></p>
      </div>
      <footer class="sc-foot">
        <button type="button" class="sc-btn sc-ghost" data-act="signin" hidden>Sign in with Bentley</button>
        <span class="sc-spacer"></span>
        ${canPick ? '<button type="button" class="sc-btn sc-ghost" data-act="download">Download instead</button>' : ''}
        <button type="button" class="sc-btn" data-act="${canPick ? 'publish' : 'download'}">${canPick ? 'Publish' : 'Download files'}</button>
      </footer>
    </form>`;
    document.body.appendChild(dlg);
    const $ = (sel) => dlg.querySelector(sel);
    const f = (name) => dlg.querySelector(`[name="${name}"]`);
    const msgEl = $('[data-msg]');
    const say = (text, kind) => { msgEl.textContent = text || ''; msgEl.className = `sc-msg${kind ? ` is-${kind}` : ''}`; };
    const state = { dir: null, busy: false };
    f('by').value = prefs.by || '';
    f('latestUrl').value = prefs.latestUrl || cfg.latestUrl || '';
    if (prefs.last) $('[data-last]').textContent = `Last published ${fmt.dateTime(prefs.last.at)}${prefs.last.folder ? ` to "${prefs.last.folder}"` : ' (downloaded)'}${prefs.last.by ? ` by ${prefs.last.by}` : ''}.`;

    function describeSource() {
      const d = getDataset();
      if (mode() === 'cloud' && (!d || d.source === 'cloud')) {
        return `<b>Live Synchro data</b>${d ? `: ${escapeHtml(d.label)}, ${fmt.int(d.rows.length)} items` : ''}. Publishing reads the latest statuses from Synchro first, so the copies show the data as at that moment.`;
      }
      const cache = d ? null : cachedWorkbook();
      const x = d || (cache ? datasetFromCache(cache) : null);
      if (x && x.rows && x.rows.length) {
        const kind = x.source === 'cloud' ? 'Synchro data' : 'Excel export';
        return `<b>${kind}</b>: ${escapeHtml(x.filename || x.label)}, ${fmt.int(x.rows.length)} items. The copies will show the data as at ${escapeHtml(fmt.dateTime(x.asAt || x.fetchedAt))}${x.modifiedAt ? ', when the file was saved' : ''}.`;
      }
      return '<b>No data loaded yet.</b> Load an Excel export or connect to Synchro, then publish.';
    }
    $('[data-source]').innerHTML = describeSource();
    async function showFolder() {
      if (!canPick) return;
      const label = $('[data-folder]');
      if (!state.dir) { label.textContent = 'No folder chosen yet.'; return; }
      let permission = 'prompt';
      try { permission = await state.dir.queryPermission({ mode: 'readwrite' }); } catch (e) { /* ignore */ }
      label.innerHTML = `Saving to <b>${escapeHtml(state.dir.name)}</b>${permission === 'granted' ? '' : '. The browser will ask to allow saving there.'}`;
    }
    const setBusy = (on) => {
      state.busy = on;
      dlg.querySelectorAll('.sc-foot button, [data-act="pick"]').forEach((b) => { b.disabled = on; });
    };
    function renderResult(r) {
      const box = $('[data-result]');
      box.hidden = false;
      box.innerHTML = `<h3><span>&#10003;</span>Published at ${escapeHtml(fmt.time(r.at))}</h3>
        <ul class="sc-files">${r.files.map((x) => `<li><b>${escapeHtml(x.name)}</b><span>${fmt.num(x.bytes / 1024, 0)} KB</span></li>`).join('')}</ul>
        <p class="sc-help">${r.folder
          ? `Saved in "${escapeHtml(r.folder)}". If that folder is synced with SharePoint, OneDrive uploads both files in a moment and SharePoint keeps the previous versions.`
          : 'Both files are downloading. Upload them to the SharePoint folder, replacing the old copies. If the browser asks about downloading multiple files, allow it.'}
          The copies show the data as at ${escapeHtml(fmt.dateTime(r.asAt))}.</p>`;
    }
    async function run(toFolder) {
      if (state.busy) return;
      const by = clean(f('by').value);
      const rawUrl = clean(f('latestUrl').value);
      const latestUrl = httpsOnly(rawUrl);
      if (rawUrl && !latestUrl) { say('The SharePoint link should be a full address starting with https://', 'bad'); f('latestUrl').focus(); return; }
      if (toFolder) {
        if (!state.dir) { say('Choose the folder first.', 'bad'); return; }
        try {
          if (!(await ensureWritable(state.dir))) { say('Saving to that folder was not allowed.', 'bad'); return; }
        } catch (e) { say(e.message, 'bad'); return; }
      }
      store.set(KEYS.publish, Object.assign(store.get(KEYS.publish, {}) || {}, { by, latestUrl }));
      setBusy(true);
      $('[data-result]').hidden = true;
      $('[data-act="signin"]').hidden = true;
      try {
        const result = await publishNow({ dataset: getDataset, dirHandle: toFolder ? state.dir : null, download: !toFolder, by, latestUrl, onProgress: (t) => say(`${t}...`) });
        renderResult(result);
        $('[data-last]').textContent = '';
        say(`Published both copies, ${fmt.int(result.items)} items.`, 'ok');
        if (hooks.onPublished) { try { hooks.onPublished(result); } catch (e) { /* the page reports its own errors */ } }
      } catch (e) {
        say(e.message, 'bad');
        if (e.authRequired) $('[data-act="signin"]').hidden = false;
      } finally {
        setBusy(false);
        showFolder();
      }
    }
    dlg.addEventListener('click', (ev) => {
      const act = ev.target.closest('[data-act]');
      if (!act || act.disabled) return;
      const which = act.dataset.act;
      if (which === 'close') dlg.close('cancel');
      else if (which === 'publish') run(true);
      else if (which === 'download') run(false);
      else if (which === 'signin') {
        session.set(KEYS.reopenPublish, '1');
        signIn(location.href).catch((e) => say(e.message, 'bad'));
      } else if (which === 'pick') {
        (async () => {
          try {
            const handle = await global.showDirectoryPicker({ id: 'mep-publish', mode: 'readwrite' });
            if (await dirHasFile(handle, 'Start Dashboards.bat')) { say("That is the dashboards' own folder. Choose the SharePoint folder instead.", 'bad'); return; }
            state.dir = handle;
            try { await idb.set('publishDir', handle); } catch (e) { /* remembered for this visit only */ }
            say('Folder chosen. Publish when ready.', 'ok');
          } catch (e) {
            if (e.name !== 'AbortError') say(e.message, 'bad');
          } finally {
            showFolder();
          }
        })();
      }
    });
    $('form').addEventListener('submit', (ev) => ev.preventDefault());
    dlg.addEventListener('close', () => setTimeout(() => dlg.remove(), 0));
    dlg.showModal();
    if (canPick) idb.get('publishDir').then((h) => { state.dir = h || null; }).catch(() => {}).then(showFolder);
    if (!f('by').value && mode() === 'cloud' && isSignedIn()) {
      whoAmI().then((me) => { if (!f('by').value && me && me.displayName) f('by').value = me.displayName; }).catch(() => {});
    }
    return dlg;
  }

  /* ------------------------------------------------------------------ sign-in gate */
  // Live pages open behind this screen: people are signed in first and see data, never an
  // empty dashboard. It takes each page's own colours from the --sc-* variables.
  const GATE_CSS = `
  html.sc-gated body{visibility:hidden}
  html.sc-gated,html.sc-gated body{overflow:hidden}
  html.sc-gated .sc-gate,html.sc-gated dialog.sc-dialog{visibility:visible}
  .sc-gate{position:fixed;inset:0;z-index:60;display:grid;place-items:center;padding:24px;overflow:auto;background:var(--sc-paper,#f6f7f5);color:var(--sc-ink,#18211f);font-family:var(--sc-font,"Segoe UI",system-ui,sans-serif)}
  .sc-gate-card{width:min(470px,100%);padding:30px 32px 28px;background:#fff;border:1px solid var(--sc-line,#d9dfda);border-top:4px solid var(--sc-accent,#1e3775);box-shadow:0 18px 50px rgba(10,20,40,.10);animation:scGateIn .25s ease .2s both}
  .sc-gate-kicker{margin:0 0 8px;font:600 11px var(--sc-mono,Consolas,monospace);letter-spacing:.08em;text-transform:uppercase;color:var(--sc-accent,#1e3775)}
  .sc-gate h2{margin:0 0 10px;font-size:21px;font-weight:700;letter-spacing:-.01em;line-height:1.25}
  .sc-gate p{margin:0 0 12px;font-size:14px;line-height:1.55;color:#4b5552}
  .sc-gate p b{color:var(--sc-ink,#18211f)}
  .sc-gate .sc-gate-error{padding:10px 12px;background:#fdecee;border:1px solid #f3b6bc;color:#8c1823;font-size:13px}
  .sc-gate-actions{display:flex;flex-wrap:wrap;align-items:center;gap:12px 18px;margin-top:18px}
  .sc-gate-btn{min-height:40px;padding:0 18px;border:1px solid var(--sc-accent,#1e3775);border-radius:3px;background:var(--sc-accent,#1e3775);color:#fff;font:700 13px var(--sc-font,"Segoe UI",sans-serif);cursor:pointer}
  .sc-gate-btn:hover{filter:brightness(1.12)}
  .sc-gate-link{border:0;background:none;padding:0;color:var(--sc-accent,#1e3775);font:600 13px var(--sc-font,"Segoe UI",sans-serif);text-decoration:underline;cursor:pointer}
  .sc-gate-spin{width:26px;height:26px;margin:0 0 16px;border:3px solid var(--sc-line,#d9dfda);border-top-color:var(--sc-accent,#1e3775);border-radius:50%;animation:scSpin .9s linear infinite}
  .sc-gate-progress{margin:0!important;min-height:16px;font:12px var(--sc-mono,Consolas,monospace)!important;color:#6b7672!important}
  @keyframes scSpin{to{transform:rotate(360deg)}}
  @keyframes scGateIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}
  @media (prefers-reduced-motion:reduce){.sc-gate-card{animation:none}.sc-gate-spin{animation-duration:2.4s}}`;
  function injectGateStyles() {
    if (document.getElementById('sc-gate-css')) return;
    const style = document.createElement('style');
    style.id = 'sc-gate-css';
    style.textContent = GATE_CSS;
    (document.head || document.documentElement).appendChild(style);
  }
  const gate = {
    el: null,
    show(opts) {
      injectGateStyles();
      const draw = () => {
        document.documentElement.classList.add('sc-gated');
        if (!gate.el) {
          gate.el = document.createElement('div');
          gate.el.className = 'sc-gate';
          gate.el.setAttribute('role', 'dialog');
          gate.el.setAttribute('aria-modal', 'true');
          gate.el.setAttribute('aria-labelledby', 'sc-gate-title');
          document.body.appendChild(gate.el);
        }
        const label = clean(config().projectLabel) || 'MEP dashboards';
        const actions = opts.actions || [];
        gate.el.dataset.state = opts.state || '';
        gate.el.innerHTML = `<div class="sc-gate-card">
          ${opts.busy ? '<div class="sc-gate-spin" aria-hidden="true"></div>' : ''}
          <p class="sc-gate-kicker">${escapeHtml(label)} · MEP cable containment</p>
          <h2 id="sc-gate-title">${escapeHtml(opts.title)}</h2>
          ${opts.message ? `<p>${opts.message}</p>` : ''}
          <p data-who hidden></p>
          ${opts.error ? `<p class="sc-gate-error">${escapeHtml(opts.error)}</p>` : ''}
          ${opts.busy ? `<p class="sc-gate-progress" data-progress role="status" aria-live="polite">${escapeHtml(opts.progress || '')}</p>` : ''}
          ${actions.length ? `<div class="sc-gate-actions">${actions.map((a, i) => `<button type="button" class="${a.primary ? 'sc-gate-btn' : 'sc-gate-link'}" data-gate-action="${i}">${escapeHtml(a.label)}</button>`).join('')}</div>` : ''}
        </div>`;
        gate.el.querySelectorAll('[data-gate-action]').forEach((button) => {
          button.addEventListener('click', () => {
            const action = actions[Number(button.dataset.gateAction)];
            if (action && action.onClick) action.onClick();
          });
        });
        const primary = gate.el.querySelector('.sc-gate-btn');
        if (primary) primary.focus();
      };
      if (document.body) draw(); else document.addEventListener('DOMContentLoaded', draw, { once: true });
    },
    busy(title, progress, state = 'busy') { gate.show({ busy: true, title, progress, state }); },
    progress(text) {
      const line = gate.el && gate.el.querySelector('[data-progress]');
      if (line) line.textContent = text;
    },
    hide() {
      if (gate.el) { gate.el.remove(); gate.el = null; }
      document.documentElement.classList.remove('sc-gated');
    },
    loading() { gate.busy('Loading the latest data', 'Reading Synchro...', 'loading'); },
    signIn(opts = {}) {
      gate.show({
        state: 'signin',
        title: 'Sign in to see the dashboard',
        message: 'This dashboard shows live installation data from Synchro. Sign in with your Bentley account to continue.',
        error: opts.error || '',
        actions: [{ label: 'Sign in with Bentley', primary: true, onClick: () => {
          session.set(KEYS.autoSignIn, String(Date.now()));
          gate.busy('Signing you in', 'Opening Bentley sign-in...', 'redirecting');
          signIn(location.href).catch((e) => gate.signIn({ error: e.message }));
        } }]
      });
    },
    noAccess(error) {
      const label = clean(config().projectLabel) || 'this project';
      gate.show({
        state: 'noaccess',
        title: 'No access to this project yet',
        message: `Your Bentley sign-in worked, but your account cannot read the Synchro schedule for ${escapeHtml(label)}. Ask the project's Bentley administrator to add you to the project, then try again.`,
        error: error && error.message,
        actions: [
          { label: 'Try again', primary: true, onClick: () => location.reload() },
          { label: 'Use a different Bentley account', onClick: () => {
            signOut();
            session.set(KEYS.autoSignIn, String(Date.now()));
            signIn(location.href, 'login').catch((e) => gate.problem(e));
          } }
        ]
      });
      whoAmI().then((me) => {
        const line = gate.el && gate.el.querySelector('[data-who]');
        if (line && me && (me.displayName || me.email)) {
          line.innerHTML = `Signed in as <b>${escapeHtml(me.displayName || me.email)}</b>${me.email && me.displayName && me.email !== me.displayName ? ` (${escapeHtml(me.email)})` : ''}.`;
          line.hidden = false;
        }
      }).catch(() => {});
    },
    problem(error, retry, extra = {}) {
      const actions = [retry
        ? { label: 'Try again', primary: true, onClick: () => { gate.loading(); retry(); } }
        : { label: 'Reload', primary: true, onClick: () => location.reload() }];
      if (extra.onExcel) actions.push({ label: 'Load an Excel export instead', onClick: extra.onExcel });
      actions.push({ label: 'Synchro settings', onClick: () => openSettings() });
      gate.show({ state: 'problem', title: "Couldn't load the Synchro data", message: 'The dashboard could not read the schedule from Bentley just now.', error: error && error.message, actions });
    },
    failed(error, retry, extra) {
      if (error && error.authRequired) gate.signIn({ error: error.message });
      else if (error && error.noAccess) gate.noAccess(error);
      else gate.problem(error, retry, extra);
    },
    setup(opts = {}) {
      const actions = [{ label: 'Set up the Synchro connection', primary: true, onClick: () => openSettings() }];
      if (opts.onExcel) actions.push({ label: 'Load an Excel export instead', onClick: opts.onExcel });
      gate.show({ state: 'setup', title: 'Connect to Synchro', message: "These dashboards haven't been connected to the Synchro schedule on this computer yet. It takes a minute, once, with your Bentley account.", actions });
    },
    welcome(opts = {}) {
      gate.show({
        state: 'welcome',
        title: 'Load the latest Synchro export',
        message: "Opened from a file, this dashboard reads an Excel export from Synchro. For live data, open the team's web link, where you sign in with your Bentley account.",
        actions: [{ label: 'Load Excel export', primary: true, onClick: opts.onExcel }]
      });
    }
  };

  // Live pages: sign the person in before anything else shows. Resolves once a token is
  // ready. Otherwise it goes straight to Bentley sign-in, or, when that was just tried or came
  // back with an error, shows a sign-in card instead of going round in a loop.
  async function requireSignIn(opts = {}) {
    if (IN_FRAME || mode() !== 'cloud') return true;
    if (readToken()) { session.remove(KEYS.autoSignIn); return true; }
    const lastTry = Number(session.get(KEYS.autoSignIn) || 0);
    let error = opts.authError ? `Bentley sign-in did not complete: ${opts.authError}` : '';
    if (!error && Date.now() - lastTry > 90e3) {
      session.set(KEYS.autoSignIn, String(Date.now()));
      gate.busy('Signing you in', 'Opening Bentley sign-in...', 'redirecting');
      try {
        await signIn(location.href);
        return new Promise(() => {});      // the browser is on its way to Bentley
      } catch (e) {
        error = e.message;
      }
    }
    gate.signIn({ error });
    return new Promise(() => {});          // the sign-in button takes it from here
  }

  /* ------------------------------------------------------------------ init */
  // A live page hides itself straight away until the person is signed in and data is ready,
  // so the empty dashboard never flashes up.
  if (IS_HTTP && !IN_FRAME && mode() === 'cloud') {
    injectGateStyles();
    document.documentElement.classList.add('sc-gated');
    setTimeout(() => {
      if (document.documentElement.classList.contains('sc-gated') && !gate.el) gate.problem(new Error('The page took too long to start.'));
    }, 20000);
    global.addEventListener('pageshow', (ev) => {
      // Back from the Bentley sign-in page: offer the button rather than a stuck spinner.
      if (ev.persisted && gate.el && gate.el.dataset.state === 'redirecting') gate.signIn({});
    });
  }
  const ready = (async () => {
    try { await handleRedirect(); } catch (e) { session.set(KEYS.authError, e.message); }
    return true;
  })();
  function init(options = {}) {
    settingsSavedHandler = options.onSettingsSaved || null;
    publishHooks = options.publish || null;
    return ready.then(() => {
      if (!IN_FRAME && session.take(KEYS.reopen) === '1') setTimeout(() => openSettings(), 50);
      else if (!IN_FRAME && session.take(KEYS.reopenPublish) === '1') setTimeout(() => openPublish(), 600);
      return { authError: session.take(KEYS.authError) };
    });
  }

  global.SynchroCloud = {
    version: VERSION, isHttp: IS_HTTP, inAuthFrame: IN_FRAME, rootUrl: ROOT.href,
    config, saveUserConfig, isCloudConfigured, mode, redirectUri, init, ready,
    signIn, signOut, isSignedIn, whoAmI, getToken, requireSignIn, gate,
    loadDataset, cachedDataset, openSettings, autoRefresh,
    parseSheet, datasetFromSheet, datasetFromCache, workbookToCache, cachedWorkbook,
    snapshot: { read: readSnapshot, describe: describeSnapshot }, openPublish, publishNow, lastPublish, publishFiles: PUBLISH_FILES,
    statusTone, looksComplete, sortStatuses, completion, snapshots,
    splitName, prettyFamily, classifyByName, fmt, escapeHtml,
    _internals: { api, getAll, store, KEYS, packSnapshot, unpackSnapshot, buildPublishedHtml, writeFiles, idb }
  };
})(window);
