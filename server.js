'use strict';
/*
 * server.js — NEA Coach's local web server.
 *
 * Listens on the network so a phone can use it, but only this Mac gets in
 * without the access code (see "Access control"). It serves the user
 * interface, scans the portfolio on demand, and watches the deck folder so it
 * can tell the student when the file has changed.
 *
 * SAFETY: the .pptx is opened read-only and is never written to. The only files
 * this app creates are its own state, cache and config, all inside the app
 * folder. See README.md for the full guarantee.
 */

const http = require('http');
const fs = require('fs');
const fsp = require('fs/promises');
const path = require('path');
const crypto = require('crypto');

const { parsePresentation } = require('./lib/pptx');
const { runLocalChecks } = require('./lib/checks');
const pdfdeck = require('./lib/pdfdeck');
const { assessPrediction: calibrationOf } = require('./lib/calibration');
const ai = require('./lib/ai');

const APP_DIR = __dirname;
const PUBLIC_DIR = path.join(APP_DIR, 'public');
const STATE_DIR = path.join(APP_DIR, 'state');
const SPEC_DIR = path.join(APP_DIR, 'spec');
const CONFIG_PATH = path.join(STATE_DIR, 'config.json');
const CACHE_PATH = path.join(STATE_DIR, 'cache.json');
const COST_PATH = path.join(STATE_DIR, 'cost.json');
const LAST_REPORT_PATH = path.join(STATE_DIR, 'last-report.json');
/*
 * The report from the check before last.
 *
 * Kept so the student can see what changed since they last ran it — which fixes
 * landed, what is new, and whether the predicted mark moved. Only one previous
 * report is kept on purpose: a full history would grow without bound and be a
 * worse answer than "since last time" for the question people actually ask.
 */
const PREVIOUS_REPORT_PATH = path.join(STATE_DIR, 'previous-report.json');
/*
 * Bounded snapshots of past scans, so the student can look back.
 *
 * Summaries rather than whole reports: a full report for a 200-page deck is a
 * quarter of a megabyte, so keeping 25 of them would be 6 MB of state and would
 * slow every page load. Each snapshot holds what the history view actually shows
 * — the mark, the counts, and the top findings — and nothing else.
 */
const HISTORY_PATH = path.join(STATE_DIR, 'history.json');
const HISTORY_LIMIT = 25;
/*
 * How many API REQUESTS have been made today, per model.
 *
 * Google's free tier caps requests per model per day (20, measured), not tokens,
 * so for a free key the thing worth showing is "how much of today's allowance have
 * I used", not a money figure that is always zero. Tracked per model because the
 * cap is per model — which is what makes the cheaper model a real escape hatch.
 */
const QUOTA_PATH = path.join(STATE_DIR, 'quota.json');

/*
 * Google resets free-tier quotas at midnight PACIFIC, not local time and not UTC.
 * Getting this wrong would tell a student in the UK their allowance had reset when
 * it had not, which is worse than saying nothing.
 */
const QUOTA_RESET_ZONE = 'America/Los_Angeles';

/** The calendar date in the reset zone, as YYYY-MM-DD. */
function quotaDate(at = new Date()) {
  try {
    return new Intl.DateTimeFormat('en-CA', {
      timeZone: QUOTA_RESET_ZONE, year: 'numeric', month: '2-digit', day: '2-digit',
    }).format(at);
  } catch (err) {
    return at.toISOString().slice(0, 10);
  }
}

/**
 * When the daily allowance next resets, and how long that is from now.
 *
 * Google resets free-tier quotas at midnight PACIFIC. Finding that instant is
 * fiddly and an 8-hour error is invisible in the output — an earlier attempt here
 * returned 15:00 UTC for "midnight Pacific", i.e. 08:00 Pacific, which would have
 * told a student their allowance had reset eight hours early.
 *
 * So this SEARCHES for the answer instead of deriving it: step through candidate
 * instants and ask the timezone database what the clock in that zone reads there,
 * taking the first that reads exactly midnight. One comparison per hour over three
 * days, so it is exact and cheap, and it cannot get a sign wrong.
 */
function quotaReset(at = new Date()) {
  const pad = (n) => String(n).padStart(2, '0');
  const localStamp = (d) => {
    try {
      const parts = new Intl.DateTimeFormat('en-CA', {
        timeZone: QUOTA_RESET_ZONE, hour12: false,
        year: 'numeric', month: '2-digit', day: '2-digit',
        hour: '2-digit', minute: '2-digit',
      }).formatToParts(d).reduce((acc, x) => { acc[x.type] = x.value; return acc; }, {});
      return `${parts.year}-${pad(parts.month)}-${pad(parts.day)} ${pad(Number(parts.hour) % 24)}:${pad(parts.minute)}`;
    } catch (err) {
      return null;
    }
  };

  const startMs = at.getTime();
  let found = null;
  // Up to 48 hours ahead, in one-minute steps near a candidate hour. Coarse-first:
  // step hourly to find the day, then by the minute to land exactly on midnight.
  for (let minutes = 0; minutes <= 60 * 48 && !found; minutes += 1) {
    const d = new Date(startMs + minutes * 60000);
    const stamp = localStamp(d);
    if (stamp && stamp.endsWith(' 00:00')) found = d;
  }
  if (!found) {
    // Fall back to a fixed 8 hours ahead rather than claiming a wrong time.
    found = new Date(startMs + 24 * 3600 * 1000);
  }
  return {
    at: found.toISOString(),
    inMinutes: Math.max(0, Math.round((found.getTime() - startMs) / 60000)),
    local: localStamp(found),
  };
}

/** Today's request counts, reset automatically when the Pacific date changes. */
function readQuota() {
  const today = quotaDate();
  const store = readJsonSafe(QUOTA_PATH, null);
  if (!store || store.date !== today) return { date: today, models: {}, requestsThisScan: 0 };
  return { date: store.date, models: store.models || {}, requestsThisScan: store.requestsThisScan || 0 };
}

function writeQuota(quota) {
  writeJsonSafe(QUOTA_PATH, quota);
}

/*
 * The daily request limit for a model.
 *
 * Google publishes the MECHANISM (requests per day, reset at midnight Pacific,
 * per project not per key) but deliberately does NOT publish the per-model
 * numbers — the rate-limits page points at AI Studio for the live values. So this
 * does not guess. It starts from a conservative assumption, and REPLACES it with
 * the true figure when a 429 tells us: the error body carries
 * `quotaValue` and the model it applies to.
 *
 * A wrong limit here would show "12% of your daily allowance used" to someone who
 * was actually at 90%, so learning it beats assuming it.
 */
const ASSUMED_DAILY_REQUESTS = 20;

function learnQuotaLimit(model, limit) {
  const n = Number(limit);
  if (!model || !Number.isFinite(n) || n <= 0) return;
  const quota = readQuota();
  quota.limits = quota.limits || {};
  if (quota.limits[model] !== n) {
    quota.limits[model] = n;
    writeQuota(quota);
    log(`Learned the daily request limit for ${model}: ${n}.`);
  }
}

/** Read the quota failure out of a 429 body, if it is one. */
function quotaLimitFromError(json) {
  try {
    const details = (json && json.error && json.error.details) || [];
    for (const d of details) {
      if (!d || !Array.isArray(d.violations)) continue;
      for (const v of d.violations) {
        if (!/PerDay/i.test(String(v.quotaId || ''))) continue;
        const model = v.quotaDimensions && v.quotaDimensions.model;
        const value = Number(v.quotaValue);
        if (model && Number.isFinite(value)) return { model, limit: value };
      }
    }
  } catch (err) { /* not a quota error */ }
  return null;
}

/** Today's usage, ready for the UI. */
function quotaReport(configuredModel, configuredMarking) {
  const quota = readQuota();
  const reset = quotaReset();
  const limits = quota.limits || {};
  const models = Object.keys(quota.models).sort();
  const rows = models.map((m) => {
    const used = quota.models[m] || 0;
    const learned = limits[m];
    const limit = learned || ASSUMED_DAILY_REQUESTS;
    /*
     * `percent` is only meaningful when the limit is the REAL one.
     *
     * It used to be reported whenever there were any counts, computed from an
     * assumed limit of 20 that Google never confirmed — and the rate-limits page
     * deliberately does not publish per-model daily numbers. A confident "70% of
     * your allowance" derived from a guess is worse than saying nothing, because the
     * student cannot tell which part is measured. So a percentage is given ONLY once
     * a 429 has told us the true figure.
     *
     * The raw count IS always real: it is the number of requests this app had
     * served, which is what it can actually know.
     */
    return {
      model: m,
      used,
      limit,
      limitKnown: !!learned,
      limitAssumed: !learned,
      remaining: learned ? Math.max(0, learned - used) : null,
      // No cap: a percentage over 100 is information (the count exceeded the limit),
      // and clamping it to 100 hides that.
      percent: learned ? Math.round((used / learned) * 100) : null,
      isMarking: !!configuredMarking && m === configuredMarking,
      isPageReader: m === configuredModel,
    };
  });
  // The tightest allowance is the thing worth showing first.
  rows.sort((a, b) => b.percent - a.percent);
  return {
    date: quota.date,
    resetAt: reset.at,
    resetInMinutes: reset.inMinutes,
    assumedLimit: ASSUMED_DAILY_REQUESTS,
    models: rows,
    // The worst case across the models actually in use. Null when no limit is known,
    // because a made-up "worst percentage" would be the same guess in another form.
    worstPercent: rows.length ? rows[0].percent : null,
    anyAssumed: rows.some((r) => r.limitAssumed),
    totalUsed: rows.reduce((t, r) => t + r.used, 0),
  };
}

/** Note one API request against a model. */
function noteRequest(model) {
  if (!model) return;
  const quota = readQuota();
  quota.models[model] = (quota.models[model] || 0) + 1;
  writeQuota(quota);
}

/*
 * Shown in the footer and returned by /api/health. It said '1.0.0' through three
 * releases, so it was actively misleading when a user reported a problem — the
 * first thing worth knowing is which build they are on. Kept in step with the
 * release tag in tools/release.sh.
 */
const VERSION = '1.0.17';
const DEFAULT_PORT = 4173;

/*
 * How much the model may "think" when the student turns deep thinking on.
 *
 * Thinking tokens are billed as output tokens, so this is a real cost, not a
 * free quality switch: at Gemini output prices a few thousand thinking tokens
 * per batch can multiply the bill. 4096 is enough to help on reasoning-heavy
 * pages without letting one request run away.
 */
const THINKING_BUDGET_TOKENS = 4096;

/* ------------------------------------------------------------------ *
 * Small utilities
 * ------------------------------------------------------------------ */

function log(msg) {
  const line = `[${new Date().toISOString().slice(11, 19)}] ${msg}`;
  console.log(line);
  scanLog.push(line);
  if (scanLog.length > 300) scanLog.shift();
}

const scanLog = [];

function readJsonSafe(file, fallback) {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch (err) {
    return fallback;
  }
}

function writeJsonSafe(file, data) {
  try {
    fs.mkdirSync(path.dirname(file), { recursive: true });
    // Write atomically so a crash mid-write cannot corrupt the cache.
    const tmp = `${file}.tmp`;
    fs.writeFileSync(tmp, JSON.stringify(data, null, 2));
    fs.renameSync(tmp, file);
    return true;
  } catch (err) {
    log(`Could not write ${path.basename(file)}: ${err.message}`);
    return false;
  }
}

function sendJson(res, status, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': Buffer.byteLength(body),
    'Cache-Control': 'no-store',
  });
  res.end(body);
}

function readBody(req, limitBytes = 1024 * 512) {
  return new Promise((resolve) => {
    let size = 0;
    const chunks = [];
    req.on('data', (c) => {
      size += c.length;
      if (size > limitBytes) {
        req.destroy();
        resolve({});
        return;
      }
      chunks.push(c);
    });
    req.on('end', () => {
      const text = Buffer.concat(chunks).toString('utf8');
      if (!text) return resolve({});
      try {
        resolve(JSON.parse(text));
      } catch (err) {
        resolve({});
      }
    });
    req.on('error', () => resolve({}));
  });
}

/* ------------------------------------------------------------------ *
 * Config
 * ------------------------------------------------------------------ */

function defaultConfig() {
  const sampleDir = path.join(APP_DIR, 'sample');
  return {
    /*
     * Where to look for the portfolio. Kept as a RELATIVE path ("..") on purpose:
     * a saved absolute path would break the moment the app folder is copied to
     * another Mac, which is exactly how this app is delivered. ".." means the
     * folder containing NEA-Coach, so the .pptx sits next to the app rather than
     * inside it — that avoids duplicating a 500 MB file in Dropbox.
     */
    deckDir: '..',
    // Blank means "use the most recently edited .pptx found in deckDir".
    deckFile: null,
    /*
     * A cheap, current model. The app re-picks automatically from what the key
     * can actually reach (see pickModel), so this is only the starting point.
     * It used to default to the most expensive flash model, which cost the
     * student roughly twice as much per scan for the same advice.
     */
    model: 'gemini-3.1-flash-lite',
    apiKey: '',
    // Automatic re-scan when the deck file changes on disk.
    watch: true,
    port: DEFAULT_PORT,
    /*
     * Listen on all interfaces so the app is reachable from a phone on the same
     * Wi-Fi, and from anywhere over Tailscale. Anything other than this Mac
     * needs the access code (see "Access control"). Set to '127.0.0.1' instead
     * to make it Mac-only.
     */
    host: '0.0.0.0',
    /*
     * Send the whole deck again instead of only pages that changed. This does
     * NOT make the AI cleverer — it only re-reads pages it already read, so it
     * costs a full scan every time. Previously exposed as "Deep check", which
     * implied smarter analysis and left the student unsure what they were paying
     * for. Renamed to say what it does.
     */
    recheckAll: false,
    /*
     * Which mark scheme to judge against. One board at a time: the app ships the
     * Eduqas A-level criteria and the AQA GCSE criteria, and sending both would
     * cost the student twice and confuse the coverage grid.
     */
    specId: 'eduqas-alevel',
    /*
     * Whether to predict a mark at all.
     *
     * The mark is the most useful single number in the report, but it is also the
     * only part that has to be asked for more than once to be trustworthy — one
     * visual review of the same portfolio came back 88% and then 81%. Three runs
     * cost about three times as much, so a student who only wants the fix list can
     * turn the mark off and pay a third of the price.
     */
    markEnabled: true,
    /*
     * The model used ONLY for the marking runs.
     *
     * Defaults to a Flash model, NOT the cheap Flash-Lite that reads the pages,
     * because the two are not close at marking. Measured against two portfolios
     * with known marks (99/100 and 96/100):
     *
     *   gemini-3.1-flash-lite  99 -> median 81 (runs 70-93)   96 -> 70-80
     *   gemini-3.5-flash       99 -> median 97 (96, 97, 98)    96 -> 92
     *   gemini-3.8-flash       99 -> median 92 (90, 92, 93)    96 -> 86
     *
     * So Flash-Lite under-marks a strong portfolio by roughly 17-26 marks, while
     * 3.5-flash is within 2-4 marks of the real mark on BOTH. It is also far more
     * repeatable: a 2-mark spread against Flash-Lite's ~20. That is a difference no
     * prompt wording ever achieved — five attempts at tuning the prompt moved
     * nothing, while changing this one setting fixed the accuracy and the
     * stability at once.
     *
     * It costs more per check (roughly 15p for a three-run marking pass on a
     * 60-page deck, against 3p), which is why it is a VISIBLE setting rather than
     * something done silently, and why "Predict a mark" can be switched off. Set
     * this to an empty string to fall back to the model above.
     */
    visionModel: 'gemini-3.5-flash',
    /*
     * Let the model reason before answering. Off by default because thinking
     * tokens are billed as output tokens, and this runs on the student's key.
     */
    thinking: false,
    // Used only until a real portfolio is found; keeps the app demonstrable.
    _sampleDir: sampleDir,
  };
}

/**
 * Turn a configured folder into an absolute path, resolving relative entries
 * against the app folder. Returns null if it cannot be resolved.
 */
function resolveDeckDir(dir) {
  if (!dir || typeof dir !== 'string') return null;
  let d = dir.trim();
  if (!d) return null;
  if (d.startsWith('~')) d = path.join(process.env.HOME || '', d.slice(1));
  if (!path.isAbsolute(d)) d = path.resolve(APP_DIR, d);
  return d;
}

function loadConfig() {
  const cfg = Object.assign(defaultConfig(), readJsonSafe(CONFIG_PATH, {}));
  // Environment variable wins, so a key can be supplied without editing files.
  if (process.env.GEMINI_API_KEY && !cfg.apiKey) cfg.apiKey = process.env.GEMINI_API_KEY;
  return cfg;
}

function saveConfig(cfg) {
  /*
   * Log the portfolio fields whenever they change.
   *
   * Chasing a report of "it kept Rohan's portfolio on the screen" meant guessing
   * what had altered the config and when, because nothing recorded it. The config
   * is written from several places (an upload, a removal, saving settings), so
   * without this the only evidence is the file's modification time.
   */
  try {
    const before = readJsonSafe(CONFIG_PATH, null);
    if (before) {
      const keys = ['deckFile', 'pdfFile', 'specId', 'model', 'visionModel'];
      const changes = keys
        .filter((k) => String(before[k] || '') !== String(cfg[k] || ''))
        .map((k) => `${k}: ${before[k] ? path.basename(String(before[k])) : '(none)'} -> ${cfg[k] ? path.basename(String(cfg[k])) : '(none)'}`);
      if (changes.length) log(`Config changed - ${changes.join('; ')}`);
    }
  } catch (err) { /* logging must never stop a save */ }
  return writeJsonSafe(CONFIG_PATH, cfg);
}

/** The config as a client may see it: no API key, no access code. */
/**
 * The config the browser is allowed to see.
 *
 * `readingSample` is passed in by buildStateResponse, which has already worked it
 * out. When the caller omits it — the /api/config handler, for instance, which does
 * not care — it is derived here so the field is never simply absent.
 */
function publicConfig(cfg, readingSample) {
  return Object.assign({}, cfg, {
    apiKey: undefined,
    accessCode: undefined,
    apiKeySet: !!cfg.apiKey,
    /*
     * The PDF's display name, derived from the LIVE path.
     *
     * This was missing, and that is why the "Marking" card never changed. The saved
     * report's own `config` block carries a `pdfName` from whenever it was written;
     * buildStateResponse merges the live config OVER that; with no `pdfName` in the
     * live config the STALE one survived the merge. So the card showed the previous
     * PDF's name no matter what the student chose, while `pdfFile` — which IS in the
     * live config — said something different.
     *
     * Deriving it here means the two can never disagree again.
     */
    pdfName: cfg.pdfFile ? path.basename(cfg.pdfFile) : null,
    /*
     * The last four characters of the key, so a student can tell WHICH key is loaded
     * without the key itself being sent to the browser. Worth having because swapping
     * keys is now routine, and two keys look identical until one runs out.
     */
    keyTail: cfg.apiKey ? String(cfg.apiKey).trim().slice(-4) : null,
    /*
     * Which mark scheme is selected, and the full list of boards the app ships.
     * The Settings dropdown is built from this list, so the UI and the server can
     * never disagree about which qualifications exist. Old saved reports predate
     * these fields, so the UI falls back to a single known board.
     */
    specId: cfg.specId || 'eduqas-alevel',
    specPresets: ai.SPEC_PRESETS.map((s) => ({ id: s.id, label: s.label })),
    /*
     * Whether the app is currently reading the BUNDLED SAMPLE rather than the
     * student's own file. Computed live, from the live config, for the same reason
     * pdfName is: buildStateResponse merges the live config OVER the saved report,
     * so any value only present in the report survives stale. That is exactly how
     * the sample flag went wrong — a report created while the sample was being read
     * kept claiming the sample even after a real portfolio was uploaded.
     */
    readingSample: readingSample !== undefined
      ? !!readingSample
      : (function () { const f = findDeck(cfg); return !!(f && !f.error && f.usedSample); }()),
  });
}

/* ------------------------------------------------------------------ *
 * Access control
 *
 * The server listens on the network so a phone can reach it. It used to have
 * no password at all: anyone on the same Wi-Fi could read the report, replace
 * the API key, point the app at any folder on the Mac and run paid scans, and
 * any web page open on the Mac could do the same by sending requests to
 * 127.0.0.1.
 *
 *  - Requests from this Mac, addressed to 127.0.0.1 or localhost, need nothing.
 *  - Everything else needs the access code, which the launcher window prints
 *    as part of the phone address (.../?key=xxxx-xxxx-xxxx). A valid code sets
 *    a cookie, so the page's own requests carry it from then on.
 *  - The Host check stops a web page that points its own domain name at
 *    127.0.0.1 (DNS rebinding) from passing as local.
 *  - Every POST must carry an X-NEA-Coach header. A browser will not let
 *    another site add a custom header without asking this server first, and
 *    this server never agrees, so a web page cannot change settings or scan.
 * ------------------------------------------------------------------ */

const LOOPBACK_ADDRESSES = new Set(['127.0.0.1', '::1', '::ffff:127.0.0.1']);
const LOCAL_HOSTNAMES = new Set(['127.0.0.1', 'localhost', '[::1]']);
const ACCESS_COOKIE = 'neaCoachKey';
// No 0/o or 1/l/i, because the code may be typed on a phone.
const ACCESS_ALPHABET = 'abcdefghjkmnpqrstuvwxyz23456789';
const ACCESS_CODE_SHAPE = /^[a-z2-9]{4}-[a-z2-9]{4}-[a-z2-9]{4}$/;

let accessCode = null;

/** About 59 bits of randomness, written as xxxx-xxxx-xxxx. */
function newAccessCode() {
  const bytes = crypto.randomBytes(12);
  let code = '';
  for (let i = 0; i < 12; i++) {
    if (i && i % 4 === 0) code += '-';
    code += ACCESS_ALPHABET[bytes[i] % ACCESS_ALPHABET.length];
  }
  return code;
}

function codesMatch(given, expected) {
  if (typeof given !== 'string' || typeof expected !== 'string' || !expected) return false;
  const a = Buffer.from(given.trim().toLowerCase());
  const b = Buffer.from(expected);
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}

function readCookie(req, name) {
  for (const part of String(req.headers.cookie || '').split(';')) {
    const i = part.indexOf('=');
    if (i > 0 && part.slice(0, i).trim() === name) {
      try {
        return decodeURIComponent(part.slice(i + 1).trim());
      } catch (err) {
        return null;
      }
    }
  }
  return null;
}

function hostnameOf(hostHeader) {
  try {
    return new URL(`http://${hostHeader || ''}`).hostname.toLowerCase();
  } catch (err) {
    return '';
  }
}

/** May this request proceed? `setCookie` means it brought a valid code in the URL. */
function checkAccess(req, url, code) {
  const fromThisMac = LOOPBACK_ADDRESSES.has(req.socket && req.socket.remoteAddress);
  if (fromThisMac && LOCAL_HOSTNAMES.has(hostnameOf(req.headers.host))) return { ok: true };
  const fromUrl = url.searchParams.get('key');
  if (fromUrl && codesMatch(fromUrl, code)) return { ok: true, setCookie: true };
  if (codesMatch(readCookie(req, ACCESS_COOKIE), code)) return { ok: true };
  return { ok: false };
}

function refuseAccess(res, urlPath) {
  if (urlPath.startsWith('/api/')) {
    return sendJson(res, 401, { ok: false, error: 'This needs the access code. Open the phone address printed in the NEA Coach window on the Mac.' });
  }
  const body = '<!doctype html><html lang="en"><head><meta charset="utf-8">'
    + '<meta name="viewport" content="width=device-width, initial-scale=1"><title>NEA Coach</title></head>'
    + '<body style="font-family:-apple-system,system-ui,sans-serif;max-width:32rem;margin:3rem auto;padding:0 16px;line-height:1.5">'
    + '<h1 style="font-size:1.4rem">NEA Coach needs its access code</h1>'
    + '<p>Open the phone address printed in the NEA Coach window on the Mac. It ends in <code>?key=</code> followed by the code.</p>'
    + '</body></html>';
  res.writeHead(401, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' });
  res.end(body);
}

/* ------------------------------------------------------------------ *
 * Deck discovery
 * ------------------------------------------------------------------ */

/** Find the .pptx to review: an explicit choice, else the most recently edited. */
function findDeck(cfg) {
  // Resolve a relative deckDir against the app folder so the app still works
  // after being copied to a different Mac with a different user name.
  const dir = resolveDeckDir(cfg.deckDir);
  if (!dir || !fs.existsSync(dir)) {
    return { error: `The portfolio folder does not exist: ${cfg.deckDir || '(not set)'}` };
  }

  if (cfg.deckFile) {
    const full = path.isAbsolute(cfg.deckFile) ? cfg.deckFile : path.join(dir, cfg.deckFile);
    if (fs.existsSync(full)) return { filePath: full, kind: 'pptx' };
    return { error: `The chosen file was not found: ${full}` };
  }

  /*
   * No PowerPoint chosen. If a PDF is set, THAT is the portfolio — the student
   * dropped or uploaded a PDF, which deliberately cleared the .pptx because the
   * PDF is the thing they want reviewed.
   *
   * It used to return an error here, so a PDF-only student got no structural
   * checks at all and could not even start a scan. The PDF is now read as the
   * source of page text, which is what lets the free checks run on it.
   */
  if (cfg.pdfFile && fs.existsSync(cfg.pdfFile)) {
    return { filePath: cfg.pdfFile, kind: 'pdf' };
  }

  let entries;
  try {
    entries = fs.readdirSync(dir, { withFileTypes: true });
  } catch (err) {
    return { error: `Could not read the portfolio folder: ${err.message}` };
  }

  const candidates = [];
  /*
   * The app's own bundled sample folder.
   *
   * The search walks up to three levels deep, and the default folder is "..",
   * the parent of the app. That means the walk can wander INTO this app's own
   * `sample/` directory and pick up the demonstration deck as though the student
   * had put it there. The consequences are quiet and bad: the fallback below
   * never runs, so `usedSample` is never set, and a report about a made-up
   * example deck is presented as though it were the student's own work.
   *
   * Skipping that one directory leaves the explicit fallback to do its job, and
   * that path DOES mark the result. Only skipped when it is not the configured
   * folder, so selecting sample/ deliberately still works.
   */
  const bundledSampleDir = path.resolve(path.join(APP_DIR, 'sample'));
  /**
   * Is this directory another copy of the app's bundled sample?
   *
   * The plain path check above misses a nested copy — an older release unzipped
   * inside the search folder, for instance. Identified by its contents rather
   * than its name, so a student's own folder called "sample" is left alone.
   */
  const looksLikeBundledSample = (d) => {
    try {
      return fs.readdirSync(d).some((n) => /^AQA-GCSE-Portfolio-SAMPLE\.pptx$/i.test(n));
    } catch (err) {
      return false;
    }
  };

  const walk = (d, depth) => {
    if (depth > 3) return;
    let list;
    try {
      list = fs.readdirSync(d, { withFileTypes: true });
    } catch (err) {
      return;
    }
    for (const e of list) {
      if (e.name.startsWith('~$') || e.name.startsWith('.')) continue; // Office lock files
      const full = path.join(d, e.name);
      if (e.isDirectory()) {
        if (path.resolve(full) !== path.resolve(dir)) {
          if (path.resolve(full) === bundledSampleDir) continue;
          if (e.name.toLowerCase() === 'sample' && looksLikeBundledSample(full)) continue;
        }
        walk(full, depth + 1);
      } else if (/\.pptx$/i.test(e.name)) {
        try {
          const st = fs.statSync(full);
          candidates.push({ filePath: full, mtime: st.mtimeMs, size: st.size });
        } catch (err) {
          /* skip unreadable */
        }
      }
    }
  };
  walk(dir, 0);

  if (!candidates.length) {
    /*
     * Nothing found. Before giving up, fall back to the bundled sample deck so a
     * freshly copied app is never dead on arrival: a student who double-clicks
     * and sees a working example is much better off than one who sees an error.
     */
    const sampleDir = path.join(APP_DIR, 'sample');
    if (path.resolve(sampleDir) !== path.resolve(dir) && fs.existsSync(sampleDir)) {
      const sample = findDeck({ deckDir: sampleDir, deckFile: null });
      if (sample && sample.filePath) {
        return Object.assign({}, sample, { usedSample: true });
      }
    }
    return { error: `No .pptx file was found in ${dir}. Put your portfolio there, or change the folder in Settings.` };
  }

  candidates.sort((a, b) => b.mtime - a.mtime);
  return { filePath: candidates[0].filePath, alternatives: candidates.map((c) => c.filePath) };
}

/* ------------------------------------------------------------------ *
 * Scan state
 * ------------------------------------------------------------------ */

const scanState = {
  running: false,
  phase: 'idle',
  progress: 0,
  message: 'Ready',
  detail: null,
  eta: null,
  startedAt: null,
  finishedAt: null,
  durationMs: null,
  slidesScanned: 0,
  slidesSkipped: 0,
  slidesChanged: 0,
  error: null,
  scanId: 0,
};

let lastReport = null;
let scanLock = Promise.resolve();

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

/* ------------------------------------------------------------------ *
 * The scan
 * ------------------------------------------------------------------ */

async function performScan({ force = false, deep = false, thinking = null, reason = 'manual' } = {}) {
  const cfg = loadConfig();
  const startedAt = Date.now();

  /*
   * Self-heal a config that names a retired model.
   *
   * A student upgraded from an older version can be left holding a model that
   * Google has since withdrawn. Rather than failing every scan with an opaque
   * 404, swap in a working model locally and say so in the log and the UI.
   * Detection is by name pattern, so this costs no API call.
   */
  if (ai.isRetiredModel(cfg.model)) {
    // Ask the picker, so the fallback follows the same rules as everywhere else
    // instead of hard-coding a model name that will itself age.
    const replacement = ai.pickModel([], null) || 'gemini-flash-lite-latest';
    log(`The saved model "${cfg.model}" has been retired by Google. Using ${replacement} instead.`);
    cfg.model = replacement;
    try { saveConfig(cfg); } catch (err) { /* best effort; the scan still runs */ }
  }

  scanState.running = true;
  scanState.phase = 'starting';
  scanState.progress = 0;
  scanState.message = 'Starting';
  scanState.detail = reason === 'auto' ? 'Detected a change to the file' : null;
  scanState.error = null;
  scanState.startedAt = new Date(startedAt).toISOString();
  scanState.finishedAt = null;
  scanState.durationMs = null;
  scanState.scanId += 1;
  scanState.slidesScanned = 0;
  scanState.slidesSkipped = 0;
  scanState.slidesChanged = 0;

  const onProgress = (p) => {
    if (p.phase) scanState.phase = p.phase;
    if (typeof p.progress === 'number') scanState.progress = Math.max(0, Math.min(1, p.progress));
    if (p.message) scanState.message = p.message;
    if (p.detail !== undefined) scanState.detail = p.detail;
    if (p.eta !== undefined) scanState.eta = p.eta;
  };

  try {
    /*
     * Today's request counts BEFORE any API work, so the scan's own usage can be
     * measured as a difference afterwards.
     *
     * Declared here rather than inside the AI branch because the count is reported
     * whether or not the AI ran — a PDF-only or key-less scan makes no requests and
     * must report zero, not crash. It previously threw "requestsBefore is not
     * defined" on every scan that skipped the AI.
     */
    let requestsBefore = readQuota();

    // ---- 1. Locate the deck.
    onProgress({ phase: 'finding', progress: 0.01, message: 'Looking for your portfolio' });
    const found = findDeck(cfg);
    if (found.error) throw new Error(found.error);

    const filePath = found.filePath;
    const stat = await fsp.stat(filePath);
    /*
     * Record the state of the file we are about to read. If the watcher fires
     * later with the same size and mtime, the scan already covered that version
     * and must not run again — which is what stopped a page reload from costing
     * money on unchanged work.
     */
    watchedStamp = { size: stat.size, mtimeMs: stat.mtimeMs };
    log(`Scanning ${path.basename(filePath)} (${(stat.size / 1048576).toFixed(1)} MB)`);

    // ---- 2. Read the pages.
    /*
     * A PDF is read for its page TEXT so the free structural checks can run on it.
     * They used to require a .pptx, so a student who supplied only a PDF got no
     * structural checks at all — even though a PDF export is a perfectly normal way
     * to hand in work. A PDF cannot supply font sizes, shape geometry or speaker
     * notes, so the rules needing those are skipped and named rather than fed
     * invented numbers (see lib/checks.js).
     */
    const fromPdf = found.kind === 'pdf' || /\.pdf$/i.test(filePath);
    onProgress({
      phase: 'reading',
      progress: 0.05,
      message: fromPdf ? 'Reading the pages of your PDF' : 'Opening the PowerPoint file',
    });
    const deckStart = Date.now();
    const deck = fromPdf
      ? pdfdeck.deckFromPdf(fs.readFileSync(filePath), { fileName: path.basename(filePath) }).deck
      : parsePresentation(filePath);
    const cache = readJsonSafe(CACHE_PATH, {});
    const fileHash = crypto
      .createHash('sha1')
      .update(`${stat.size}:${stat.mtimeMs}`)
      .digest('hex')
      .slice(0, 12);

    log(`Read ${deck.slideCount} pages in ${Date.now() - deckStart}ms`);

    // Detect which pages actually changed since last time, so the UI can be
    // honest about what was re-read.
    const prevFingerprints = cache.fingerprints || {};
    let changedPages = 0;
    const newFingerprints = {};
    for (const s of deck.slides) {
      newFingerprints[String(s.index)] = s.fingerprint;
      if (prevFingerprints[String(s.index)] !== s.fingerprint) changedPages += 1;
    }
    scanState.slidesChanged = changedPages;
    scanState.slidesScanned = deck.slideCount;

    // ---- 3. Free local checks.
    // These always run on every page: they cost nothing, so there is no reason
    // to skip them, and they are the backbone of the report.
    //
    // The mark scheme is loaded first because the free checks depend on it: some
    // rules are board-specific. Eduqas, for instance, does not require annotation,
    // so the caption rule must not fire for it.
    const spec = ai.loadSpec(SPEC_DIR, cfg.specId);
    onProgress({ phase: 'local', progress: 0.3, message: 'Running the free checks' });
    const local = runLocalChecks(deck, {
      annotationRequired: spec.annotationRequired !== false,
      /*
       * Tells the checks which rules can run. A PDF cannot supply font sizes, shape
       * geometry or speaker notes, so those rules are skipped and named rather than
       * silently reporting "nothing found", which would read as a pass.
       */
      source: fromPdf ? 'pdf' : 'pptx',
      /*
       * The page timeline from previous scans, so the app can say when each page
       * last actually changed. PowerPoint does not record that per slide — it
       * rewrites the whole archive on save — so the app keeps its own record,
       * matched on page CONTENT so inserting a page does not make every later page
       * look modified.
       */
      pageHistory: cache.pageHistory,
      now: new Date().toISOString(),
    });
    log(`Local checks: ${local.findings.length} findings, score ${local.score}`);

    // ---- 4. AI pass.
    const cost = readJsonSafe(COST_PATH, { totalGbp: 0, totalTokensIn: 0, totalTokensOut: 0, lastScanGbp: 0, scans: 0, pricingKnown: true });

    let aiResult = null;
    if (cfg.apiKey) {
      onProgress({ phase: 'ai', progress: 0.35, message: 'Starting the AI review' });
      /*
       * If the student supplied a PDF export of the portfolio, send it so the AI
       * can actually LOOK at the pages. Without it the review is text-only and
       * under-scores visual work — measured at 75% for a portfolio that really
       * scored 99/100. A missing or unreadable PDF is not an error; the scan
       * simply falls back to the text-only review.
       */
      let pdf = null;
      if (cfg.pdfFile) {
        try {
          if (fs.existsSync(cfg.pdfFile)) {
            pdf = fs.readFileSync(cfg.pdfFile);
            log(`Visual review will use ${path.basename(cfg.pdfFile)} (${(pdf.length / 1048576).toFixed(1)} MB)`);
          } else {
            log(`The saved PDF is missing (${cfg.pdfFile}), so this scan is text-only.`);
          }
        } catch (err) {
          log(`Could not read the PDF (${err.message}), so this scan is text-only.`);
          pdf = null;
        }
      }
      /*
       * The mark is asked for three times and averaged, because one answer is not
       * stable enough to quote. With marks switched off there is no reason to pay
       * for three runs: one review still supplies the findings and suggestions.
       * A separate, stronger model can be used for the marking runs alone.
       */
      const markOn = cfg.markEnabled !== false;
      const wantVisionModel = String(cfg.visionModel || '').trim();
      try {
        /*
         * Count the real API requests this scan makes, per model, against the
         * free tier's daily cap. The counter is passed all the way down to the
         * HTTP call so retries and the model probe are counted too — they consume
         * the same allowance and an under-count would tell the student they have
         * budget they do not.
         */
        requestsBefore = readQuota();
        const modelAtStart = cfg.model;
        const markingModelAtStart = String(cfg.visionModel || '').trim() || cfg.model;
        const countRequest = (which) => noteRequest(which || modelAtStart);

        aiResult = await ai.runAiPass({
          deck,
          apiKey: cfg.apiKey,
          model: cfg.model,
          spec,
          cache,
          onRequest: () => countRequest(cfg.model),
          pdf,
          pdfName: cfg.pdfFile ? path.basename(cfg.pdfFile) : null,
          // 3 runs when a mark is wanted, 1 when it is not.
          visionRuns: markOn ? 3 : 1,
          visionModel: wantVisionModel || null,
          markEnabled: markOn,
          force: force || deep || cfg.recheckAll,
          thinkingBudget: (thinking == null ? !!cfg.thinking : !!thinking) ? THINKING_BUDGET_TOKENS : 0,
          onProgress,
          onLog: log,
        });
      } catch (err) {
        aiResult = {
          ok: false,
          error: `The AI review stopped unexpectedly: ${err.message}`,
          findings: [],
          suggestions: [],
          strengths: [],
          coverage: [],
          tokensIn: 0,
          tokensOut: 0,
          costGbp: 0,
          pricingKnown: true,
        };
        log(`AI pass threw: ${err.message}`);
      }

      // Persist the batch cache and running cost.
      if (aiResult && aiResult.cache) {
        cache.batches = aiResult.cache.batches;
        cache.model = aiResult.cache.model;
        cache.crossCheck = aiResult.cache.crossCheck || null;
      }
      cache.fingerprints = newFingerprints;
      cache.pageHistory = (local.pageHistory && local.pageHistory.history) || cache.pageHistory;
      cache.deckFile = filePath;
      cache.savedAt = new Date().toISOString();
      writeJsonSafe(CACHE_PATH, cache);

      if (aiResult && (aiResult.tokensIn || aiResult.tokensOut)) {
        const nextCost = {
          totalGbp: Number((cost.totalGbp || 0) + (aiResult.costGbp || 0)),
          totalTokensIn: (cost.totalTokensIn || 0) + aiResult.tokensIn,
          totalTokensOut: (cost.totalTokensOut || 0) + aiResult.tokensOut,
          lastScanGbp: aiResult.costGbp || 0,
          lastScanTokensIn: aiResult.tokensIn,
          lastScanTokensOut: aiResult.tokensOut,
          scans: (cost.scans || 0) + 1,
          pricingKnown: (cost.pricingKnown !== false) && (aiResult.pricingKnown !== false),
          model: aiResult.model,
          updatedAt: new Date().toISOString(),
        };
        writeJsonSafe(COST_PATH, nextCost);
      } else {
        cost.pricingKnown = (cost.pricingKnown !== false) && (!aiResult || aiResult.pricingKnown !== false);
        writeJsonSafe(COST_PATH, cost);
      }
    } else {
      cache.fingerprints = newFingerprints;
      cache.pageHistory = (local.pageHistory && local.pageHistory.history) || cache.pageHistory;
      cache.deckFile = filePath;
      writeJsonSafe(CACHE_PATH, cache);
    }

    /*
     * How many API requests this scan consumed, by model. Recorded in the report
     * so the UI can show "requests used today" without a second round trip, and so
     * a scan's cost in quota is visible next to its findings.
     */
    /*
     * If the API told us the real daily cap, believe it over our assumption.
     */
    if (aiResult && aiResult.quotaError) {
      const learned = quotaLimitFromError({ error: aiResult.quotaError });
      if (learned) learnQuotaLimit(learned.model || aiResult.quotaErrorModel, learned.limit);
    }

    const requestsAfter = readQuota();
    const requestsThisScan = {};
    let requestsTotal = 0;
    for (const m of new Set([...Object.keys(requestsBefore.models), ...Object.keys(requestsAfter.models)])) {
      const delta = (requestsAfter.models[m] || 0) - (requestsBefore.models[m] || 0);
      if (delta > 0) { requestsThisScan[m] = delta; requestsTotal += delta; }
    }

    // ---- 5. Build the report.
    const aiFindings = (aiResult && aiResult.findings) || [];
    const aiCoverage = (aiResult && aiResult.coverage) || [];
    /*
     * The predicted mark per criterion, and its total, are computed HERE rather
     * than by the model: the model is asked for a mark per criterion, but only
     * the app knows the mark scheme's weights, so only the app can guarantee the
     * total never exceeds 100%. A criterion the model skipped falls back to its
     * coverage percentage and is flagged, so the number shown is never invented.
     */
    const predicted = ai.summarisePredictedMarks(spec.criteria, aiCoverage);
    /*
     * How much to trust that number, and what the app has measured about its own error.
     *
     * The raw figure is kept as the AI produced it; the calibrated one is a SEPARATE
     * value with its own explanation, applied only inside the range of portfolios the
     * app has actually been checked against. It is never silent: the student is told
     * the raw number, the adjusted one, and why they differ. A prediction presented
     * without its error bar invites more trust than it has earned.
     */
    const markModel = String(cfg.visionModel || '').trim() || cfg.model;
    const calibration = calibrationOf(predicted.predicted, predicted.maximum, {
      board: cfg.specId,
      model: markModel,
    });
    const allFindings = [...local.findings, ...aiFindings]
      .sort((a, b) => {
        if (b.priority !== a.priority) return b.priority - a.priority;
        const as = a.slide == null ? 1e9 : a.slide;
        const bs = b.slide == null ? 1e9 : b.slide;
        return as - bs;
      })
      .map((f, i) => Object.assign({}, f, { rank: i + 1 }));

    const bySeverity = { critical: 0, high: 0, medium: 0, low: 0, info: 0 };
    for (const f of allFindings) bySeverity[f.severity] = (bySeverity[f.severity] || 0) + 1;

    // Recompute the health score including AI findings, weighted so that
    // mark-costing problems dominate.
    const pages = Math.max(1, deck.slideCount);
    const penalty =
      ((bySeverity.critical * 12 + bySeverity.high * 4.5 + bySeverity.medium * 1.6 + bySeverity.low * 0.5) / pages) * 10;
    const combinedScore = Math.max(0, Math.min(100, Math.round(100 - penalty)));

    const finalCost = readJsonSafe(COST_PATH, cost);

    lastReport = {
      ok: true,
      version: VERSION,
      generatedAt: new Date().toISOString(),
      deck: {
        fileName: deck.fileName,
        filePath: deck.filePath,
        slideCount: deck.slideCount,
        stats: buildStats(deck, aiResult),
        /* Guarded: a source that cannot report its size must not crash the report. */
        slideWidthIn: Number.isFinite(deck.slideWidthIn) ? Number(deck.slideWidthIn.toFixed(2)) : null,
        slideHeightIn: Number.isFinite(deck.slideHeightIn) ? Number(deck.slideHeightIn.toFixed(2)) : null,
        sizeLabel:
          !Number.isFinite(deck.aspect) ? 'size not recorded in this file'
          : Math.abs(deck.aspect - Math.SQRT2) < 0.06
            ? 'A3/A4 landscape (ISO proportions)'
            : Math.abs(deck.aspect - 16 / 9) < 0.06
            ? '16:9 widescreen'
            : Math.abs(deck.aspect - 4 / 3) < 0.06
            ? '4:3 classic'
            : 'custom size',
        modified: stat.mtime.toISOString(),
        fileSizeBytes: stat.size,
        meta: deck.meta,
        changedPages,
        fileHash,
        /*
         * Per-page detail for the Slides table. This was missing at first, which
         * left the table showing em dashes for every count even though the parser
         * had all the data. Only the fields the table needs are sent, to keep the
         * report payload small on a 200-page deck.
         */
        slides: deck.slides.map((s) => {
          const tracked = ((local.pageHistory && local.pageHistory.pages) || []).find((x) => x.number === s.index);
          return {
            number: s.index,
            title: s.title,
            wordCount: s.wordCount,
            imageCount: s.imageCount,
            notesWordCount: s.notesWordCount,
            hasNotes: s.notesWordCount > 0,
            // When this page's content last actually changed, from the app's own
            // record rather than a file timestamp PowerPoint does not keep.
            changeStatus: tracked ? tracked.status : null,
            firstSeen: tracked ? tracked.firstSeen : null,
            lastChanged: tracked ? tracked.lastChanged : null,
          };
        }),
      },
      scan: Object.assign({}, scanState, { running: false }),
      local: {
        // Named for what it is. NOT a mark: these rules judge presentation and
        // completeness only, so the UI shows a count of points to check instead.
        presentationTidiness: local.presentationTidiness,
        score: local.presentationTidiness, // deprecated alias for older saved reports
        bySeverity: local.bySeverity,
        /*
         * byTier lets the UI separate "needs your thinking" from "quick tidies"
         * without re-deriving it, and disabledRules lets it say which checks the
         * student has switched off. Both were computed but dropped here, so the
         * UI could not tell what it was looking at.
         */
        byTier: local.byTier,
        disabledRules: local.disabledRules,
        findings: local.findings,
        ruleCount: local.ruleCount,
        /*
         * Which source the pages came from, and which rules could not run on it.
         *
         * A PDF cannot supply font sizes, shape geometry or speaker notes, so some
         * checks are skipped. The UI says which, because a rule that never runs
         * reports nothing and "nothing" reads as a pass.
         */
        source: local.source,
        totalRules: local.totalRules,
        skippedRules: (local.skippedRules || []).map((x) => x.rule),
        errors: local.errors,
      },
      ai: {
        enabled: !!cfg.apiKey,
        keySet: !!cfg.apiKey,
        model: (aiResult && aiResult.model) || cfg.model,
        ran: !!aiResult,
        findings: aiFindings,
        /*
         * "Worth adding when you have time" — a separate list from the fix list.
         * Only AI findings the model marked as concerning work not yet present
         * appear here, so the student is never told to add something they have
         * already done. Ranked by impact against effort.
         */
        suggestions: (aiResult && aiResult.suggestions) || [],
        // Additions the whole-portfolio check found already present, with why.
        suggestionsWithdrawn: (aiResult && aiResult.suggestionsWithdrawn) || [],
        strengths: (aiResult && aiResult.strengths) || [],
        coverage: aiCoverage,
        /*
         * An ESTIMATE of what the work currently looks like against this mark
         * scheme, produced by the AI and totalled by the app. It is not a grade
         * and not a prediction of the final result — the student's own teacher
         * and the moderator decide that. Every surface that shows it must say so.
         */
        predicted: Object.assign({}, predicted, { calibration }),
        criteria: spec.criteria || [],
        specName: spec.name,
        specId: spec.specId,
        specLabel: spec.label,
        specFiles: spec.files || [],
        specSkipped: spec.skipped || [],
        specTruncated: !!spec.truncated,
        batchesRun: (aiResult && aiResult.batchesRun) || 0,
        batchesCached: (aiResult && aiResult.batchesCached) || 0,
        /*
         * Whether the AI actually looked at the pages, or only read the text.
         * Surfaced so the UI can say which it was: the two produce materially
         * different marks, and the student deserves to know which they got.
         */
        visionRan: !!(aiResult && aiResult.visionRan),
        visionTokensIn: (aiResult && aiResult.visionTokensIn) || 0,
        // How many times the pages were reviewed, and how far the mark moved
        // between runs. Shown as a range: a single number implies a precision
        // that three runs of the same portfolio (88% then 81%) do not have.
        visionRuns: (aiResult && aiResult.visionRuns) || 0,
        visionMarkRange: (aiResult && aiResult.visionMarkRange) || null,
        visionMarkSpread: (aiResult && aiResult.visionMarkSpread) || 0,
        /*
         * Whether the repeated reviews disagreed enough that the mark should not
         * be quoted as a single confident figure. Surfaced so the UI can say so:
         * a median cannot make an unstable answer stable, and a student acting on
         * an unreliable number is worse served than one told to be cautious.
         */
        visionUnstable: ((aiResult && aiResult.visionMarkSpread) || 0) >= ai.VISION_SPREAD_CAUTION,
        visionModel: (aiResult && aiResult.visionModel) || cfg.visionModel || null,
        markEnabled: cfg.markEnabled !== false,
        pdfName: cfg.pdfFile ? path.basename(cfg.pdfFile) : null,
        slidesSkipped: (aiResult && aiResult.slidesSkipped) || 0,
        note: (aiResult && aiResult.note) || null,
        error: (aiResult && aiResult.error) || null,
        availableModels: (aiResult && aiResult.availableModels) || [],
      },
      combined: { score: combinedScore, bySeverity, findings: allFindings },
      cost: {
        totalGbp: Number((finalCost.totalGbp || 0).toFixed(4)),
        totalTokensIn: finalCost.totalTokensIn || 0,
        totalTokensOut: finalCost.totalTokensOut || 0,
        lastScanGbp: Number((finalCost.lastScanGbp || 0).toFixed(4)),
        lastScanTokensIn: finalCost.lastScanTokensIn || 0,
        lastScanTokensOut: finalCost.lastScanTokensOut || 0,
        scans: finalCost.scans || 0,
        pricingKnown: finalCost.pricingKnown !== false,
        currency: 'GBP',
        // Requests, not money: this is what a free-tier key actually spends.
        requestsThisScan,
        requestsTotal,
      },
      /*
       * Today's allowance, included in the report so the UI needs no second
       * request to draw it.
       */
      quota: quotaReport(cfg.model, String(cfg.visionModel || '').trim() || null),
      thresholds: local.thresholds,
      config: {
        deckDir: cfg.deckDir,
        deckFile: cfg.deckFile,
        pdfFile: cfg.pdfFile || null,
        pdfName: cfg.pdfFile ? path.basename(cfg.pdfFile) : null,
        model: cfg.model,
        apiKeySet: !!cfg.apiKey,
        watch: !!cfg.watch,
        recheckAll: !!cfg.recheckAll,
        thinking: !!cfg.thinking,
        specLoaded: !!(spec.files && spec.files.length),
        specName: spec.name,
        specCriteriaCount: (spec.criteria || []).length,
        specSkipped: spec.skipped || [],
        // Frozen into the report so opening an old report cannot claim the sample
        // was the student's own work.
        readingSample: !!found.usedSample,
      },
      log: scanLog.slice(-60),
    };

    /*
     * Record completion BEFORE building the report.
     *
     * The finally block also clears scanState.running, but it only runs after
     * this function returns, by which point the report has already been
     * serialised. That ordering bug meant the first scan after a restart was
     * saved with finishedAt: null, so the UI said "Last checked: never" even
     * though a scan had just completed.
     */
    const finishedAt = Date.now();
    scanState.running = false;
    scanState.phase = 'idle';
    scanState.progress = 1;
    scanState.message = aiResult && aiResult.note
      ? 'Finished with a note'
      : changedPages === 0 && aiResult && aiResult.skipped
      ? 'Up to date — nothing had changed'
      : 'Finished';
    scanState.detail = null;
    scanState.eta = null;
    scanState.finishedAt = new Date(finishedAt).toISOString();
    scanState.durationMs = finishedAt - startedAt;

    lastReport.scan = Object.assign({}, scanState, { running: false });

    /*
     * Preserve the report being replaced so the next page load can show what
     * changed. Written BEFORE the new one, so a crash between the two leaves the
     * previous report intact rather than losing both.
     */
    const before = readJsonSafe(LAST_REPORT_PATH, null);
    if (before && before.ok) writeJsonSafe(PREVIOUS_REPORT_PATH, before);

    /*
     * Append a summary of this scan to the bounded history, so past checks can be
     * looked back at. Written before the report itself so a crash cannot leave the
     * history out of step with what the UI is about to show.
     */
    appendHistory(snapshotOf(lastReport));

    writeJsonSafe(LAST_REPORT_PATH, lastReport);

    log(`Scan complete: ${allFindings.length} findings, score ${combinedScore}`);
    return lastReport;
  } catch (err) {
    scanState.phase = 'error';
    scanState.message = 'Could not finish';
    scanState.error = err && err.message ? err.message : String(err);
    log(`Scan failed: ${scanState.error}`);
    throw err;
  } finally {
    scanState.running = false;
    scanState.finishedAt = new Date().toISOString();
    scanState.durationMs = Date.now() - startedAt;
  }
}

/** Serialise scans so two can never run at once. */
function requestScan(opts) {
  const run = scanLock.then(
    () => performScan(opts),
    () => performScan(opts)
  );
  scanLock = run.catch(() => {});
  return run;
}

/* ------------------------------------------------------------------ *
 * File watcher
 * ------------------------------------------------------------------ */

let watcher = null;
let watchTimer = null;
let watchedPath = null;
/*
 * The size and mtime of the deck as of the last time we looked at it.
 *
 * Used to decide whether a filesystem event is a REAL change. A watcher on a
 * synced folder (Dropbox, iCloud) or one started while the file is open can fire
 * without the file having been edited at all. Any such event used to start a
 * 4-second-debounced scan, which meant reloading a page could cost the student
 * money with no change to the work.
 */
let watchedStamp = null;
/*
 * When the watcher was last (re)started. Events in the first couple of seconds
 * are the watcher's own start-up burst, not a save, so they are ignored. Set by
 * startWatching; startWatching also runs after every settings save.
 */
let watchStartedAt = 0;
const WATCH_SETTLE_MS = 2500;

/** Size and mtime of a file, or null if it cannot be read. */
function stampOf(file) {
  try {
    const st = fs.statSync(file);
    return { size: st.size, mtimeMs: st.mtimeMs };
  } catch (err) {
    return null;
  }
}

/** Has the deck on disk really changed since we last looked? */
function deckChangedSinceStamp(file) {
  const now = stampOf(file);
  if (!now) return false;
  if (!watchedStamp) {
    watchedStamp = now;
    return false;
  }
  const changed = now.size !== watchedStamp.size || Math.abs(now.mtimeMs - watchedStamp.mtimeMs) > 1000;
  if (!changed) watchedStamp = now;
  return changed;
}

function startWatching() {
  const cfg = loadConfig();
  if (watcher) {
    try {
      watcher.close();
    } catch (err) {
      /* ignore */
    }
    watcher = null;
  }
  if (!cfg.watch) {
    log('Automatic re-scan is off.');
    return;
  }

  const found = findDeck(cfg);
  if (found.error) return;
  const target = found.filePath;
  const dir = path.dirname(target);

  try {
    // Remember what the file looked like when we started watching, so the next
    // event can be compared against it.
    watchedStamp = stampOf(target);
    watchStartedAt = Date.now();

    watcher = fs.watch(dir, { persistent: true }, (eventType, filename) => {
      if (!filename) return;
      if (filename.startsWith('~$')) return; // Office lock file
      if (!/\.pptx$/i.test(filename)) return;
      const full = path.join(dir, filename.toString());
      if (path.resolve(full) !== path.resolve(target)) return;

      /*
       * Ignore the watcher's own start-up events. Without this, starting the app
       * (or saving any setting, which restarts the watcher) fired a scan on a
       * file nobody had edited.
       */
      if (Date.now() - watchStartedAt < WATCH_SETTLE_MS) return;

      if (watchTimer) clearTimeout(watchTimer);
      // Debounce: PowerPoint touches files several times while saving.
      watchTimer = setTimeout(() => {
        if (scanState.running) return;
        /*
         * Final check that the file actually differs from the last one we read.
         * This is what stops a page reload, or a sync client touching the file,
         * from spending the student's money on a scan of unchanged work.
         */
        if (!deckChangedSinceStamp(target)) {
          log('File event ignored — the portfolio has not actually changed.');
          return;
        }
        log('Change detected — starting an automatic scan.');
        requestScan({ force: false, reason: 'auto' }).catch(() => {});
      }, 4000);
    });
    watchedPath = target;
    log(`Watching for changes: ${path.basename(target)}`);
  } catch (err) {
    log(`Could not watch the portfolio file: ${err.message}`);
  }
}

/* ------------------------------------------------------------------ *
 * Static files
 * ------------------------------------------------------------------ */

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
};

function serveStatic(req, res, urlPath) {
  let rel = urlPath === '/' ? 'index.html' : urlPath.replace(/^\/+/, '');
  // Prevent path traversal out of the public folder.
  const full = path.normalize(path.join(PUBLIC_DIR, rel));
  if (!full.startsWith(PUBLIC_DIR)) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }
  fs.readFile(full, (err, data) => {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Not found');
      return;
    }
    res.writeHead(200, {
      'Content-Type': MIME[path.extname(full).toLowerCase()] || 'application/octet-stream',
      'Cache-Control': 'no-store',
    });
    res.end(data);
  });
}

/* ------------------------------------------------------------------ *
 * Routing
 * ------------------------------------------------------------------ */

/**
 * A stable identity for a finding, across scans.
 *
 * NOT the `id` field: that carries a sequence number which increments on every
 * scan (`thin-visual-evidence@deck#104`, then `#127`), and the AI's id also
 * embeds the title. Comparing raw ids therefore reports every finding as
 * simultaneously fixed and new — which is worse than showing nothing.
 *
 * Local checks are identified by RULE plus page, which is exact.
 *
 * AI findings get NO single key. The model rewords the same finding every scan —
 * a real example, from one unchanged portfolio: "Distance sensor failure" became
 * "Failure of core electronic functionality", "Functional Failure of Distance
 * Sensor" and "Core functionality failure" across three runs. Any key built from
 * the title therefore splits one finding into four, which is what produced a
 * report of "21 problems fixed, 22 new findings" on a file nobody had touched.
 * They are matched in matchAiFindings() instead, on pages AND wording.
 */
function findingKey(f) {
  if (!f) return null;
  if (String(f.source || '') === 'ai') return null;
  const rule = String(f.rule || '').toLowerCase();
  const slide = f.slide == null ? 'deck' : String(f.slide);
  return rule ? `${rule}@${slide}` : null;
}

/**
 * Pair up the AI findings of two reports by the PAGES they are about.
 *
 * Deliberately not by title. The model rewords and re-groups its own findings
 * between scans of an unchanged file: on one real portfolio "Distance sensor
 * failure" became "Failure of core electronic functionality", "Functional Failure
 * of Distance Sensor" and "Core functionality failure", and one problem spanning
 * pages 53,55 became three entries with slightly different page lists. Matching on
 * words produced "21 problems fixed, 22 new findings" on a file nobody had touched.
 *
 * Pages are far more stable, because the finding is about the same part of the
 * portfolio. A pair matches when their page sets overlap, and a larger overlap
 * scores higher. This still cannot be exact — the model may cite 53,54,55 where it
 * previously cited 53,55 — so matches are reported as "still open" and the UI says
 * plainly that AI findings are grouped by page, not identified.
 */
function matchAiFindings(previousList, currentList) {
  const prev = (previousList || []).filter((f) => String(f.source || '') === 'ai');
  const curr = (currentList || []).filter((f) => String(f.source || '') === 'ai');

  const pagesOf = (f) => new Set((f.slides && f.slides.length ? f.slides : [f.slide]).filter((n) => n != null));

  const candidates = [];
  for (let i = 0; i < prev.length; i++) {
    const a = pagesOf(prev[i]);
    if (!a.size) continue;
    for (let j = 0; j < curr.length; j++) {
      const b = pagesOf(curr[j]);
      if (!b.size) continue;
      let shared = 0;
      for (const p of a) if (b.has(p)) shared += 1;
      if (!shared) continue;
      // Jaccard overlap: a large shared portion beats a glancing one.
      const union = new Set([...a, ...b]).size;
      candidates.push({ i, j, score: shared / union });
    }
  }
  candidates.sort((a, b) => b.score - a.score);

  const usedPrev = new Set();
  const usedCurr = new Set();
  const matched = [];
  for (const c of candidates) {
    if (usedPrev.has(c.i) || usedCurr.has(c.j)) continue;
    usedPrev.add(c.i);
    usedCurr.add(c.j);
    matched.push({ previous: prev[c.i], current: curr[c.j], score: c.score });
  }

  return {
    matched,
    /*
     * Unpaired on either side. NOT reported to the student as "fixed" and "new":
     * the model's own grouping shifts between runs, so an unpaired finding usually
     * means the wording or the page list moved, not that the work changed.
     */
    unpairedPrevious: prev.filter((_, i) => !usedPrev.has(i)),
    unpairedCurrent: curr.filter((_, j) => !usedCurr.has(j)),
    previousCount: prev.length,
    currentCount: curr.length,
  };
}

/**
 * What changed between two checks.
 *
 * Kept pure and exported so it can be tested without running a scan. A finding
 * present last time and gone now has been FIXED, one that is new has appeared,
 * and the rest are unchanged. That is the answer to "did my work get better",
 * which is the question a student actually has after acting on the list.
 */
function compareReports(previous, current) {
  if (!previous || !previous.ok || !current || !current.ok) return null;

  const prevAll = (previous.combined && previous.combined.findings) || [];
  const currAll = (current.combined && current.combined.findings) || [];

  /*
   * Local checks are exact: rule plus page. If the same rule fires on the same
   * page it is the same problem, and if it stops firing it is genuinely fixed.
   */
  const indexLocal = (list) => {
    const map = new Map();
    for (const f of list) {
      const key = findingKey(f);
      // First one wins, so a repeated key does not flap between scans.
      if (key && !map.has(key)) map.set(key, f);
    }
    return map;
  };
  const prevLocalMap = indexLocal(prevAll);
  const currLocalMap = indexLocal(currAll);

  const fixed = [];
  for (const [key, f] of prevLocalMap) if (!currLocalMap.has(key)) fixed.push(f);
  const added = [];
  for (const [key, f] of currLocalMap) if (!prevLocalMap.has(key)) added.push(f);

  /*
   * AI findings are paired on the pages they concern, and are deliberately NOT
   * reported as "fixed" or "new". The model rewords and re-groups them between
   * scans of an untouched file, so an unpaired one almost always means the wording
   * moved rather than the work. Claiming they were fixed is what produced
   * "21 problems fixed, 22 new findings" on a portfolio nobody had edited.
   */
  const aiMatch = matchAiFindings(prevAll, currAll);

  /*
   * How many problems are still open.
   *
   * Defined as (what there was) minus (what got fixed) plus (what is new), rather
   * than as a sum of differently-counted things. The previous version added a
   * deduplicated local count to raw AI counts, so 103 findings was reported as
   * "105 still to do".
   */
  const stillOpenCount = (prevLocalMap.size + aiMatch.previousCount)
    - fixed.length + added.length;

  /*
   * Split the count by how much each finding deserves the student's attention.
   *
   * Without this the header said "103 still to do", which is true and useless: 93
   * of those 129 findings are mechanical items the app itself folds away in the
   * Fix list (a font size, a stray box). The handful that actually cost marks are
   * invisible inside a number that size, so substance and mechanics are counted
   * separately.
   */
  const tally = (list) => {
    const out = { substance: 0, presentation: 0, obvious: 0 };
    for (const f of list) {
      const t = (f && f.tier) || 'presentation';
      out[t] = (out[t] || 0) + 1;
    }
    return out;
  };
  const currentTiers = tally(currAll);

  const prevPred = (previous.ai && previous.ai.predicted) || null;
  const currPred = (current.ai && current.ai.predicted) || null;
  const markChange = (prevPred && currPred
    && typeof prevPred.percentage === 'number' && typeof currPred.percentage === 'number')
    ? {
        from: prevPred.percentage,
        to: currPred.percentage,
        delta: currPred.percentage - prevPred.percentage,
        fromMarks: prevPred.predicted,
        toMarks: currPred.predicted,
      }
    : null;

  /* Per-criterion movement, so the student can see WHICH part improved. */
  const byCriterion = [];
  if (prevPred && currPred && Array.isArray(prevPred.rows) && Array.isArray(currPred.rows)) {
    for (const row of currPred.rows) {
      const before = prevPred.rows.find((r) => String(r.id) === String(row.id));
      if (!before || typeof before.mark !== 'number' || typeof row.mark !== 'number') continue;
      byCriterion.push({
        id: row.id,
        name: row.name,
        from: before.mark,
        to: row.mark,
        max: row.max,
        delta: row.mark - before.mark,
      });
    }
  }

  // Cap the lists: the UI shows a summary, and a 100-item payload on every page
  // load would slow the phone down for no benefit.
  const brief = (f) => ({ id: f.id, title: f.title, slides: f.slides || [], severity: f.severity, source: f.source });

  return {
    previousAt: previous.generatedAt || null,
    currentAt: current.generatedAt || null,
    /* Only exact matches: same rule, same page. */
    fixedCount: fixed.length,
    addedCount: added.length,
    stillOpenCount,
    fixed: fixed.slice(0, 40).map(brief),
    added: added.slice(0, 40).map(brief),
    fixedTotal: fixed.length,
    addedTotal: added.length,
    /* How the "still to do" number is made up, so 103 cannot look catastrophic. */
    tiers: currentTiers,
    substantiveOpen: currentTiers.substance || 0,
    mechanicalOpen: (currentTiers.presentation || 0) + (currentTiers.obvious || 0),
    /*
     * AI findings, reported as grouped rather than identified. `paired` is how
     * many could be tied to a page that was already flagged, so the student can
     * see the AI's view has not changed much even though its wording has.
     */
    ai: {
      previousCount: aiMatch.previousCount,
      currentCount: aiMatch.currentCount,
      paired: aiMatch.matched.length,
      unpairedPrevious: aiMatch.unpairedPrevious.length,
      unpairedCurrent: aiMatch.unpairedCurrent.length,
      /* So the UI can be honest about why there is no per-finding list. */
      note: 'The AI rewords the same finding between checks, so its findings are grouped by page rather than listed as fixed or new.',
    },
    markChange,
    byCriterion,
    changedPages: (current.deck && current.deck.changedPages) || 0,
  };
}

/*
 * Words per minute used for the "time spent typing" figure.
 *
 * 40 wpm is the widely quoted average for sustained composition — slower than
 * copying speed, faster than careful editing — and it deliberately UNDER-states the
 * hours, because inventing a flattering number would be worse than a modest one. The
 * UI names the rate so the student can judge it.
 */
const TYPING_WPM = 40;
// A marker reading for assessment reads far more slowly than for pleasure.
const MARKER_WPM = 150;

/**
 * Interesting totals about the portfolio.
 *
 * Computed HERE, on the server, because the full text of each page never reaches the
 * browser — `/api/state` sends only a summary per slide, and sending every page's
 * text on every poll would be megabytes. The server has the text during a scan, so
 * the figures are worked out once and travel as a small object.
 *
 * Every figure is measured from the actual document. Nothing here is a mark, and
 * nothing is a guess dressed as a measurement.
 */
function buildStats(deck, aiResult) {
  if (!deck || !Array.isArray(deck.slides)) return null;
  const slides = deck.slides;

  let words = 0;
  let chars = 0;
  let images = 0;
  let notes = 0;
  let longestSentence = 0;
  let longestSentencePage = null;
  let emptyPages = 0;
  let titledPages = 0;
  const freq = new Map();

  // Words too common to say anything about how the portfolio reads.
  const STOP = new Set(('a an the and or but if then than that this these those of to in on at by for with from as is are was were be been being it its it\'s i you he she they we my your our their there here not no so such can could would should may might will shall do does did done have has had also more most very much many any all both each few other some only own same too s t just now over under again further once during before after above below up down out off about into through between while because when where which who whom what how').split(/\s+/));

  for (const s of slides) {
    const text = String(s.text || '');
    words += s.wordCount || 0;
    chars += s.charCount || text.length;
    images += s.imageCount || 0;
    notes += s.notesWordCount || 0;
    if ((s.wordCount || 0) === 0) emptyPages += 1;
    if (s.hasTitle || s.title) titledPages += 1;

    /*
     * Sentence length from the text itself. A crude split on . ! ? is right for a
     * readability signal — it is not grammar, and does not need to be.
     */
    const sentences = text.split(/[.!?]+\s+|[.!?]+$/).map((x) => x.trim()).filter(Boolean);
    for (const sent of sentences) {
      const n = sent.split(/\s+/).filter(Boolean).length;
      if (n > longestSentence) { longestSentence = n; longestSentencePage = s.index; }
    }

    for (const rawWord of text.toLowerCase().split(/[^a-z0-9'’-]+/)) {
      const w = rawWord.replace(/^['’-]+|['’-]+$/g, '');
      // Only words long enough to be worth remarking on, and not the glue words.
      if (w.length < 5 || STOP.has(w)) continue;
      freq.set(w, (freq.get(w) || 0) + 1);
    }
  }

  const pages = slides.length || 0;
  const topWords = Array.from(freq.entries())
    .filter(([, n]) => n >= 4)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8)
    .map(([word, count]) => ({ word, count }));

  /*
   * Evidence balance per criterion, from the marks the AI awarded and the pages it
   * cited. Per PAGE words and images only — the app does not attribute individual
   * words to a criterion, and claiming to would be a fabrication.
   */
  const criteria = (aiResult && aiResult.coverage) || [];
  const balance = criteria.map((c) => {
    const pagesCited = Array.from(new Set((c.slides || []).filter((n) => Number.isFinite(n))));
    let cw = 0;
    let ci = 0;
    for (const p of pagesCited) {
      const slide = slides.find((x) => x.index === p);
      if (!slide) continue;
      cw += slide.wordCount || 0;
      ci += slide.imageCount || 0;
    }
    return {
      id: String(c.id),
      name: String(c.name || ''),
      weight: c.weight,
      mark: c.suggestedMark != null ? c.suggestedMark : null,
      pagesCited: pagesCited.length,
      words: cw,
      images: ci,
      // Words per image on the pages that evidence this criterion. Null when the
      // criterion cites no images at all, which is itself the interesting fact.
      wordsPerImage: ci > 0 ? Math.round(cw / ci) : null,
    };
  });

  return {
    pages,
    words,
    chars,
    images,
    notesWords: notes,
    emptyPages,
    titledPages,
    averageWordsPerPage: pages ? Math.round(words / pages) : 0,
    imagesPerPage: pages ? Number((images / pages).toFixed(1)) : 0,
    wordsPerImage: images > 0 ? Math.round(words / images) : null,
    // The stated rates, so the reader can judge the figures.
    typingWpm: TYPING_WPM,
    markerWpm: MARKER_WPM,
    typingHours: Number((words / TYPING_WPM / 60).toFixed(1)),
    readingMinutes: Math.round(words / MARKER_WPM),
    // How long a marker spends on the whole thing, at a normal reading pace.
    markerMinutes: Math.round(words / MARKER_WPM),
    longestSentence,
    longestSentencePage,
    topWords,
    balance,
  };
}

/**
 * A single self-contained HTML document holding everything worth acting on.
 *
 * WHY HTML AND NOT A .docx
 *
 * It has to be SELF-CONTAINED: no images, no external stylesheet, no fonts. One file
 * that opens on any machine, prints to PDF or Word, and can be emailed or handed to a
 * teacher. A .docx would need the Office tooling and would break in more ways for no
 * gain, since the point is to read and act on it rather than to edit it.
 *
 * Every string from the report is escaped: it contains student writing and AI output,
 * and an unescaped angle bracket would silently break the document.
 */
function escapeHtml(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function exportHtml(report, opts = {}) {
  const r = report || {};
  const deck = r.deck || {};
  const ai = r.ai || {};
  const stats = r.stats || {};
  const pred = ai.predicted || null;
  const findings = (r.combined && r.combined.findings) || [];
  const local = (r.local && r.local.findings) || [];
  const criteria = ai.criteria || [];
  const when = r.generatedAt ? new Date(r.generatedAt) : new Date();

  const esc = escapeHtml;
  const plural = (n, w) => `${n} ${w}${n === 1 ? '' : 's'}`;

  /*
   * Ordered for someone working through it: substance first, then presentation, then
   * the mechanical tidy-ups, which is the same order the app uses.
   */
  const TIER_ORDER = { substance: 0, presentation: 1, obvious: 2 };
  const tierName = {
    substance: 'Worth marks — do these first',
    presentation: 'Presentation — quick wins',
    obvious: 'Mechanical tidy-ups',
  };
  const ordered = findings.slice().sort((a, b) => {
    const t = (TIER_ORDER[a.tier] || 1) - (TIER_ORDER[b.tier] || 1);
    if (t) return t;
    const sev = { critical: 0, high: 1, medium: 2, low: 3, info: 4 };
    return (sev[a.severity] || 9) - (sev[b.severity] || 9);
  });

  const byTier = { substance: [], presentation: [], obvious: [] };
  for (const f of ordered) {
    const t = byTier[f.tier] ? f.tier : 'presentation';
    byTier[t].push(f);
  }

  const findingHtml = (f, i) => {
    const pages = (f.slides && f.slides.length ? f.slides : (f.slide ? [f.slide] : []));
    return `<div class="f">
      <div class="f-head"><span class="f-num">${i + 1}</span>
        <span class="f-title">${esc(f.title)}</span>
        <span class="sev sev-${esc(f.severity)}">${esc(f.severity)}</span>
        ${pages.length ? `<span class="pages">page ${esc(pages.join(', '))}</span>` : ''}
      </div>
      ${f.evidence ? `<p class="ev">${esc(f.evidence)}</p>` : ''}
      ${f.why ? `<p><strong>Why it matters.</strong> ${esc(f.why)}</p>` : ''}
      ${f.fix ? `<p class="fix"><strong>What to do.</strong> ${esc(f.fix)}</p>` : ''}
      ${f.effort ? `<p class="effort">Effort: ${esc(f.effort)}</p>` : ''}
    </div>`;
  };

  const criteriaHtml = criteria.length ? criteria.map((c) => {
    const row = (ai.coverage || []).find((x) => String(x.id) === String(c.id)) || {};
    const mark = row.suggestedMark != null ? row.suggestedMark : (row.mark != null ? row.mark : null);
    return `<tr>
      <td>${esc(c.id)} — ${esc(c.name)}</td>
      <td class="num">${mark == null ? '—' : esc(mark)} / ${esc(c.weight)}</td>
      <td>${esc(row.status || '')}</td>
      <td>${esc(row.comment || '')}</td>
    </tr>`;
  }).join('') : '';

  const strengthsHtml = (ai.strengths || []).length
    ? `<ul>${ai.strengths.map((x) => `<li>${esc(x.title || x)}${x.slides && x.slides.length ? ` <span class="pages">(page ${esc(x.slides.join(', '))})</span>` : ''}</li>`).join('')}</ul>`
    : '<p class="muted">None recorded.</p>';

  const suggestionsHtml = (ai.suggestions || []).length
    ? `<ul>${ai.suggestions.map((x) => `<li><strong>${esc(x.title || '')}</strong>${x.why ? ` — ${esc(x.why)}` : ''}</li>`).join('')}</ul>`
    : '<p class="muted">None recorded.</p>';

  /*
   * A tick-list of every fix, because the document's purpose is to be worked through
   * and crossing things off on paper is how that happens.
   */
  const checklist = ordered.filter((f) => f.fix || f.title).map((f) => {
    const pages = (f.slides && f.slides.length ? f.slides : (f.slide ? [f.slide] : []));
    return `<li>☐ ${esc(f.title)}${pages.length ? ` <span class="pages">(page ${esc(pages.join(', '))})</span>` : ''}</li>`;
  }).join('');

  const balanceRows = (stats.balance || []).map((b) => `<tr>
    <td>${esc(b.id)} — ${esc(b.name)}</td>
    <td class="num">${esc(b.pagesCited)}</td>
    <td class="num">${esc(b.words)}</td>
    <td class="num">${esc(b.images)}</td>
    <td class="num">${b.wordsPerImage == null ? '—' : esc(b.wordsPerImage)}</td>
  </tr>`).join('');

  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Portfolio feedback — ${esc(deck.fileName || 'portfolio')}</title>
<style>
  :root { color-scheme: light; }
  body { margin: 0 auto; max-width: 900px; padding: 32px 22px 60px;
    font: 15px/1.6 -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    color: #14181f; background: #fff; }
  h1 { font-size: 26px; margin: 0 0 4px; }
  h2 { font-size: 19px; margin: 34px 0 10px; padding-bottom: 6px; border-bottom: 2px solid #e6e9ef; }
  h3 { font-size: 15px; margin: 22px 0 8px; text-transform: uppercase; letter-spacing: .05em; color: #5a6472; }
  p { margin: 0 0 10px; }
  .muted { color: #6b7280; }
  .head { border-bottom: 3px solid #14181f; padding-bottom: 14px; margin-bottom: 8px; }
  .meta { font-size: 13.5px; color: #5a6472; }
  .big { display: flex; flex-wrap: wrap; gap: 12px; margin: 16px 0 6px; }
  .big div { flex: 1 1 150px; border: 1px solid #e6e9ef; border-radius: 8px; padding: 10px 12px; }
  .big .v { font-size: 24px; font-weight: 700; line-height: 1.15; }
  .big .l { font-size: 12.5px; color: #6b7280; }
  table { width: 100%; border-collapse: collapse; font-size: 14px; margin: 6px 0 12px; }
  th, td { text-align: left; padding: 7px 8px; border-bottom: 1px solid #e6e9ef; vertical-align: top; }
  th { background: #f6f8fa; font-size: 12.5px; text-transform: uppercase; letter-spacing: .04em; color: #5a6472; }
  td.num, th.num { text-align: right; white-space: nowrap; }
  .f { border: 1px solid #e6e9ef; border-left-width: 4px; border-radius: 8px; padding: 12px 14px; margin: 0 0 12px; }
  .f-head { display: flex; flex-wrap: wrap; gap: 8px; align-items: baseline; margin-bottom: 6px; }
  .f-num { font-weight: 700; color: #6b7280; }
  .f-title { font-weight: 700; font-size: 16px; }
  .pages { font-size: 12.5px; color: #5a6472; }
  .sev { font-size: 11px; text-transform: uppercase; letter-spacing: .04em; padding: 2px 7px; border-radius: 999px; background: #eef1f5; }
  .sev-critical, .sev-high { border-left-color: #b91c1c; }
  .sev-medium { border-left-color: #b45309; }
  .sev-low, .sev-info { border-left-color: #1d4ed8; }
  .ev { font-style: italic; color: #3d4653; }
  .fix { background: #f0f7f0; border-radius: 6px; padding: 8px 10px; }
  .effort { font-size: 13px; color: #6b7280; margin-bottom: 0; }
  ul { margin: 6px 0 12px; padding-left: 22px; }
  li { margin-bottom: 5px; }
  @media print { body { padding: 0; max-width: none; } h2 { page-break-after: avoid; } .f { page-break-inside: avoid; } }
</style></head><body>

<div class="head">
  <h1>Portfolio feedback</h1>
  <p class="meta">
    ${esc(deck.fileName || 'portfolio')} · ${plural(deck.slideCount || 0, 'page')} ·
    ${esc(ai.specLabel || ai.specName || 'mark scheme not recorded')} ·
    prepared ${esc(when.toISOString().slice(0, 10))}
  </p>
  <p class="meta">${plural(findings.length, 'point')} to work through.</p>
</div>

${pred && pred.predicted != null ? `
<h2>Where the work stands</h2>
<div class="big">
  <div><div class="v">${esc(pred.calibration && pred.calibration.applied ? pred.calibration.calibrated : pred.predicted)} / ${esc(pred.maximum || 100)}</div><div class="l">predicted mark${pred.calibration && pred.calibration.applied ? ' (adjusted)' : (pred.percentage != null ? ` (${esc(pred.percentage)}%)` : '')}</div></div>
  <div><div class="v">${esc(byTier.substance.length)}</div><div class="l">worth marks</div></div>
  <div><div class="v">${esc(byTier.presentation.length + byTier.obvious.length)}</div><div class="l">presentation and tidy-ups</div></div>
</div>
<p class="muted">This is a second opinion, not the examiner's. The fix list below is the part to act on.
${ai.visionUnstable ? ' The app found this mark unstable between readings, so treat the number cautiously.' : ''}</p>
${pred.calibration && pred.calibration.note ? `<p class="muted"><strong>About that number.</strong> ${esc(pred.calibration.note)}</p>` : ''}
${criteriaHtml ? `<table><thead><tr><th>Criterion</th><th class="num">Mark</th><th>Status</th><th>Comment</th></tr></thead><tbody>${criteriaHtml}</tbody></table>` : ''}
` : ''}

<h2>What to do, in order</h2>
${byTier.substance.length ? `<h3>${tierName.substance} (${byTier.substance.length})</h3>${byTier.substance.map(findingHtml).join('')}` : ''}
${byTier.presentation.length ? `<h3>${tierName.presentation} (${byTier.presentation.length})</h3>${byTier.presentation.map((f, i) => findingHtml(f, byTier.substance.length + i)).join('')}` : ''}
${byTier.obvious.length ? `<h3>${tierName.obvious} (${byTier.obvious.length})</h3>${byTier.obvious.map((f, i) => findingHtml(f, byTier.substance.length + byTier.presentation.length + i)).join('')}` : ''}
${!findings.length ? '<p>No problems were found. That is unusual — check the mark scheme and page count above are right.</p>' : ''}

<h2>Tick-list</h2>
<p class="muted">Print this page and cross items off.</p>
<ul class="checklist">${checklist || '<li>Nothing outstanding.</li>'}</ul>

<h2>What is already working</h2>
${strengthsHtml}

<h2>Things worth adding</h2>
<p class="muted">Only if there is time — these are additions, not faults.</p>
${suggestionsHtml}

${stats.words ? `
<h2>The document itself</h2>
<div class="big">
  <div><div class="v">${esc(stats.words)}</div><div class="l">words</div></div>
  <div><div class="v">${esc(stats.pages)}</div><div class="l">pages</div></div>
  <div><div class="v">${esc(stats.images)}</div><div class="l">images</div></div>
  <div><div class="v">${esc(stats.typingHours)} h</div><div class="l">typing at ${esc(stats.typingWpm)} wpm</div></div>
  <div><div class="v">${esc(stats.readingMinutes)} min</div><div class="l">for a marker to read</div></div>
</div>
${stats.topWords && stats.topWords.length ? `<p><strong>Words you lean on:</strong> ${stats.topWords.map((w) => `${esc(w.word)} (${esc(w.count)})`).join(', ')}.</p>` : ''}
${balanceRows ? `<table><thead><tr><th>Criterion</th><th class="num">Pages cited</th><th class="num">Words</th><th class="num">Images</th><th class="num">Words/image</th></tr></thead><tbody>${balanceRows}</tbody></table>
<p class="muted">Counted from the pages the app said evidence each criterion.</p>` : ''}
` : ''}

<h2>How this was produced</h2>
<p class="muted">
  ${esc((local.length ? plural(local.length, 'structural check') + ' and ' : '') + (ai.findings ? plural(ai.findings.length, 'AI finding') : 'no AI findings'))} were combined into the list above.
  Structural checks are free and run on this computer; the AI review uses a Google Gemini key.
  ${ai.model ? `Marking model: ${esc(ai.model)}.` : ''}
  The app never modifies the original PowerPoint.
</p>

</body></html>`;
}

/**
 * The small, durable summary of one scan.
 *
 * Deliberately does NOT include every finding. A history of 25 full finding lists
 * would be megabytes and would make the history view itself slow; what a student
 * wants from history is the mark over time and what changed, which needs counts
 * and only the most important items.
 */
function snapshotOf(report) {
  if (!report || !report.ok) return null;
  const findings = (report.combined && report.combined.findings) || [];
  const pred = (report.ai && report.ai.predicted) || null;
  // Keep only the substance, capped: enough to recognise a scan, not a full list.
  const top = findings
    .filter((f) => f && f.tier === 'substance')
    .slice(0, 12)
    .map((f) => ({ title: f.title, slides: f.slides || [], severity: f.severity, source: f.source }));
  return {
    at: report.generatedAt,
    deck: (report.deck && report.deck.fileName) || null,
    pages: (report.deck && report.deck.slideCount) || 0,
    changedPages: (report.deck && report.deck.changedPages) || 0,
    predictedPercent: pred && typeof pred.percentage === 'number' ? pred.percentage : null,
    predictedMarks: pred ? pred.predicted : null,
    maximum: pred ? pred.maximum : null,
    markRange: pred ? pred.markRange : null,
    visionRan: !!(report.ai && report.ai.visionRan),
    visionModel: report.ai ? report.ai.visionModel : null,
    findingsTotal: findings.length,
    bySeverity: (report.combined && report.combined.bySeverity) || null,
    byTier: report.local && report.local.byTier ? report.local.byTier : null,
    substantive: top,
    costGbp: (report.cost && report.cost.lastScanGbp) || 0,
    /*
     * What the check spent in REQUESTS. On a free key that is the real cost, so the
     * history shows this rather than a money figure that is always zero.
     */
    requestsTotal: (report.cost && report.cost.requestsTotal) || 0,
    durationMs: (report.scan && report.scan.durationMs) || null,
  };
}

/** Append a snapshot, keeping the newest HISTORY_LIMIT. */
function appendHistory(snapshot) {
  if (!snapshot) return;
  const store = readJsonSafe(HISTORY_PATH, { scans: [] });
  const scans = Array.isArray(store.scans) ? store.scans : [];
  scans.push(snapshot);
  // Newest first for the UI, trimmed to the cap.
  const trimmed = scans.slice(-HISTORY_LIMIT);
  writeJsonSafe(HISTORY_PATH, { scans: trimmed });
}

function buildStateResponse() {
  const cfg = loadConfig();
  const report = lastReport || readJsonSafe(LAST_REPORT_PATH, null);

  /*
   * Is the app currently reading the STUDENT's portfolio, or the bundled sample?
   *
   * findDeck falls back to sample/ so that a freshly copied app is never dead on
   * arrival. That is the right behaviour for a demo, but it was silent: a user
   * with no portfolio pressed Check now, got a full report and a cheerful
   * "opened" banner, and had no way to know the findings were about a stranger's
   * coursework. Reported by the user, and it is a bad first impression.
   *
   * `usedSample` was already being set by findDeck and never read by anything.
   * This reads it and tells the UI, which shows a standing banner until a real
   * file is added.
   */
  const current = findDeck(cfg);
  const readingSample = !!(current && !current.error && current.usedSample);

  if (!report) {
    /*
     * No report yet — a brand-new install, or one that has never run a check.
     *
     * This branch used to hand-roll its own config object, and that object was
     * missing `specId` and `specPresets`. The consequences were not cosmetic:
     *
     *   - the exam-board dropdown fell back to a SINGLE hard-coded Eduqas option,
     *     so there was no way to choose AQA at all;
     *   - the first-run wizard's board list came back EMPTY for the same reason;
     *   - and because the client's fallback id matched the server's default, the
     *     wizard's "nothing changed, skip the save" shortcut fired, so the board
     *     was never written. A student could not pick their board, and their
     *     config silently stayed on the default.
     *
     * It now goes through publicConfig() exactly like the normal path, so the two
     * can never drift again. That was the whole point of publicConfig.
     */
    return {
      ok: true,
      version: VERSION,
      stateStale: true,
      deck: null,
      scan: Object.assign({}, scanState),
      local: { score: null, bySeverity: { critical: 0, high: 0, medium: 0, low: 0, info: 0 }, findings: [] },
      ai: {
        enabled: !!cfg.apiKey,
        keySet: !!cfg.apiKey,
        model: cfg.model,
        findings: [],
        strengths: [],
        coverage: [],
        criteria: [],
        error: null,
      },
      cost: readJsonSafe(COST_PATH, { totalGbp: 0, totalTokensIn: 0, totalTokensOut: 0, lastScanGbp: 0, scans: 0, pricingKnown: true, currency: 'GBP' }),
      config: Object.assign(publicConfig(cfg, readingSample), {
        specLoaded: false,
        specName: null,
        specCriteriaCount: 0,
      }),
      firstRun: true,
    };
  }

  /*
   * Report live scan state rather than the state frozen into the report, BUT fall
   * back to the report whenever a live value is missing.
   *
   * scanState is cleared when the process starts, so after a restart its
   * finishedAt and durationMs are null while the saved report still knows when
   * the last scan actually finished. Overwriting with those nulls is precisely
   * what produced "Last checked: never" on a portfolio that had just been
   * checked — the page was reading a null that the server had introduced, not a
   * missing timestamp.
   */
  const live = (v, fallback) => (v === null || v === undefined ? fallback : v);
  /*
   * The exam-board list is merged in here rather than trusted from the file.
   *
   * The UI reads THIS endpoint, not /api/config, so a saved report written before
   * the board selector existed carried no specPresets and the dropdown silently
   * fell back to a single hard-coded board — which is exactly why AQA was missing
   * as an option. Deriving it at serve time means the list is always current.
   */
  const publicCfg = publicConfig(cfg, readingSample);
  const previousReport = readJsonSafe(PREVIOUS_REPORT_PATH, null);
  return Object.assign({}, report, {
    /*
     * Today's allowance, derived live rather than trusted from the saved report.
     * A report written before this existed would otherwise leave the meter blank,
     * and one written yesterday would show yesterday's usage.
     */
    quota: quotaReport(cfg.model, String(cfg.visionModel || '').trim() || null),
    /*
     * Portfolio statistics, for the Stats card. A small summary — the page text they
     * are derived from never leaves the server, because sending every page's text on
     * every poll would be megabytes.
     */
    stats: report.stats || null,
    /*
     * What changed since the check before this one. Null when there is no earlier
     * report to compare against, which the UI treats as "nothing to compare yet"
     * rather than as an error.
     */
    comparison: compareReports(previousReport, report),
    config: Object.assign({}, report.config, publicCfg),
    /*
     * Keep the saved report's AI section, but derive the predicted marks from it
     * NOW rather than trusting a number frozen into an old file. A report saved
     * before this feature existed has coverage but no prediction, and recomputing
     * means the student sees the new panel immediately instead of after a scan.
     * It also means changing the exam board cannot leave marks from the old board
     * on screen under the new board's name.
     */
    ai: Object.assign({}, report.ai, {
      specId: (report.ai && report.ai.specId) || cfg.specId,
      specLabel: (report.ai && report.ai.specLabel) || ai.specPreset((report.ai && report.ai.specId) || cfg.specId).label,
      predicted: Object.assign(
        {},
        /*
         * Recomputed from the stored coverage so the totals always agree with the
         * criteria shown. The CALIBRATION is read from the report rather than
         * recomputed, because it depends on which model marked the work, and only the
         * report knows that — recomputing here from the live config would attribute a
         * mark produced by one model to whatever is selected now.
         */
        ai.summarisePredictedMarks(
          (report.ai && report.ai.criteria) || [],
          (report.ai && report.ai.coverage) || []
        ),
        { calibration: (report.ai && report.ai.predicted && report.ai.predicted.calibration) || null }
      ),
    }),
    scan: Object.assign({}, report.scan, {
      running: scanState.running,
      phase: live(scanState.phase, report.scan.phase),
      progress: live(scanState.progress, report.scan.progress),
      message: live(scanState.message, report.scan.message),
      detail: live(scanState.detail, report.scan.detail),
      error: scanState.error || report.scan.error || null,
      startedAt: live(scanState.startedAt, report.scan.startedAt),
      finishedAt: live(scanState.finishedAt, report.scan.finishedAt),
      durationMs: live(scanState.durationMs, report.scan.durationMs),
      scanId: live(scanState.scanId, report.scan.scanId),
    }),
  });
}

/**
 * Accept a portfolio the student picked with the browser's file chooser.
 *
 * Added because the alternative was asking a student to find their own file's
 * path, in Finder, and paste it into a box — which is exactly the step people
 * get wrong. This is still fully local: the bytes go from the browser to this
 * server on the same machine and are written into state/, and nothing is
 * uploaded anywhere.
 *
 * The body is the RAW file, not multipart/form-data, because parsing multipart
 * would mean either a dependency (forbidden here) or a hand-rolled parser, and
 * the browser can send the file's bytes directly with fetch(). The file name
 * travels in the query string and is sanitised below.
 *
 * SAFETY: writes only inside state/. The original file the student chose is read
 * by the browser and is never opened, moved or modified here.
 */
function handleUpload(req, res) {
  const url = new URL(req.url, 'http://localhost');
  const rawName = String(url.searchParams.get('name') || 'portfolio.pptx');
  // Take the final path segment and strip anything risky: this name comes from
  // the client, so it must never be able to escape state/.
  const safeName = path.basename(rawName).replace(/[^A-Za-z0-9 ._-]/g, '_').slice(-120);

  if (!/\.pptx$/i.test(safeName)) {
    return sendJson(res, 400, { ok: false, error: 'That is not a .pptx file. Export your portfolio from PowerPoint as .pptx, then choose it again.' });
  }

  // A deck this size is already far beyond anything a school portfolio needs,
  // and refusing early is better than filling the disk.
  const MAX_BYTES = 2 * 1024 * 1024 * 1024;
  const declared = Number(req.headers['content-length'] || 0);
  if (declared > MAX_BYTES) {
    return sendJson(res, 400, { ok: false, error: 'That file is larger than 2 GB, which is bigger than any portfolio should be.' });
  }

  const dir = path.join(STATE_DIR, 'portfolio');
  const target = path.join(dir, safeName);
  const tmp = `${target}.part`;

  /*
   * The folder is created here rather than at start-up so an install that never
   * uploads anything leaves no empty folder behind. Without this the first upload
   * failed with ENOENT, which the student would have seen as "Could not save".
   */
  try {
    fs.mkdirSync(dir, { recursive: true });
  } catch (err) {
    return sendJson(res, 500, { ok: false, error: `Could not create the folder to save into: ${err.message}` });
  }

  let size = 0;
  let aborted = false;
  const out = fs.createWriteStream(tmp);

  const fail = (message) => {
    if (aborted) return;
    aborted = true;
    try { out.destroy(); } catch (err) { /* ignore */ }
    try { fs.unlinkSync(tmp); } catch (err) { /* ignore */ }
    if (!res.headersSent) sendJson(res, 400, { ok: false, error: message });
    log(`Upload rejected: ${message}`);
  };

  req.on('data', (chunk) => {
    size += chunk.length;
    if (size > MAX_BYTES) {
      fail('That file is larger than 2 GB, which is bigger than any portfolio should be.');
      req.destroy();
    }
  });
  req.on('error', () => fail('The upload was interrupted. Try again.'));

  req.pipe(out);

  out.on('error', (err) => fail(`Could not save the file: ${err.message}`));

  out.on('close', () => {
    if (aborted) return;
    try {
      fs.renameSync(tmp, target);
    } catch (err) {
      return fail(`Could not save the file: ${err.message}`);
    }

    /*
     * Point the app at the uploaded copy. deckFile is absolute here, and the
     * config is saved, so the student does not have to touch a folder path at
     * all. Their original file is left exactly where it was.
     */
    const cfg = loadConfig();
    cfg.deckDir = dir;
    cfg.deckFile = target;
    saveConfig(cfg);
    startWatching();

    log(`Portfolio received: ${safeName} (${(size / 1048576).toFixed(1)} MB) saved to ${target}`);
    return sendJson(res, 200, {
      ok: true,
      file: safeName,
      savedTo: target,
      sizeBytes: size,
      // Enough for the UI to say something specific without another round trip.
      sizeLabel: `${(size / 1048576).toFixed(1)} MB`,
    });
  });
}

/**
 * Accept a PDF export of the portfolio.
 *
 * This is what turns on the visual review. The AI otherwise only ever sees text
 * extracted from the .pptx, and in Design and Technology the marks live largely
 * in sketches, photographs and CAD, so a text-only judgement under-scores
 * visual work. A real portfolio that scored 99/100 was predicted at 75% by the
 * text-only path and 88% once the pages could be seen.
 *
 * Stored beside the deck in state/portfolio. Nothing is uploaded anywhere: the
 * PDF sits on this machine and is sent to Google only as part of an AI scan.
 */
function handlePdfUpload(req, res) {
  const url = new URL(req.url, 'http://localhost');
  const rawName = String(url.searchParams.get('name') || 'portfolio.pdf');
  const safeName = path.basename(rawName).replace(/[^A-Za-z0-9 ._-]/g, '_').slice(-120);

  if (!/\.pdf$/i.test(safeName)) {
    return sendJson(res, 400, { ok: false, error: 'That is not a PDF. In PowerPoint use File → Export → Create PDF/XPS, then choose that file.' });
  }

  /*
   * The inline-data limit in the Gemini API is about 20 MB, and the app sends
   * the PDF in place. Refuse anything near it with an explanation, rather than
   * letting a scan fail deep inside the AI pass with an opaque error.
   */
  const MAX_BYTES = 18 * 1024 * 1024;
  const declared = Number(req.headers['content-length'] || 0);
  if (declared > MAX_BYTES) {
    return sendJson(res, 400, {
      ok: false,
      error: `That PDF is ${(declared / 1048576).toFixed(1)} MB. The app can send up to about 18 MB. In PowerPoint, export again with "Minimum size" compression, or export a page range.`,
    });
  }

  const dir = path.join(STATE_DIR, 'portfolio');
  const target = path.join(dir, safeName);
  const tmp = `${target}.part`;

  try {
    fs.mkdirSync(dir, { recursive: true });
  } catch (err) {
    return sendJson(res, 500, { ok: false, error: `Could not create the folder to save into: ${err.message}` });
  }

  let size = 0;
  let aborted = false;
  const out = fs.createWriteStream(tmp);

  const fail = (message) => {
    if (aborted) return;
    aborted = true;
    try { out.destroy(); } catch (err) { /* ignore */ }
    try { fs.unlinkSync(tmp); } catch (err) { /* ignore */ }
    if (!res.headersSent) sendJson(res, 400, { ok: false, error: message });
    log(`PDF upload rejected: ${message}`);
  };

  req.on('data', (chunk) => {
    size += chunk.length;
    if (size > MAX_BYTES) {
      fail(`That PDF is larger than the ${Math.round(MAX_BYTES / 1048576)} MB the app can send. Export it with "Minimum size" compression, or export a page range.`);
      req.destroy();
    }
  });
  req.on('error', () => fail('The upload was interrupted. Try again.'));
  req.pipe(out);
  out.on('error', (err) => fail(`Could not save the file: ${err.message}`));

  out.on('close', () => {
    if (aborted) return;
    try {
      fs.renameSync(tmp, target);
    } catch (err) {
      return fail(`Could not save the file: ${err.message}`);
    }

    /*
     * Store the PDF. The deck is NOT touched.
     *
     * This used to compare the PDF's filename against the deck's and warn when
     * they differed. That was wrong twice over: it blocked the legitimate and
     * common case of reviewing someone ELSE's portfolio (a friend's PDF with the
     * reader's own .pptx still configured), and the warning it gave was false —
     * it claimed the marks would be wrong when nothing had happened yet.
     *
     * Dropping a PDF means "review this document". So the deck in use is CLEARED,
     * which is both what the student meant and what removes the mismatch by
     * construction. The reply states that in plain terms instead of warning.
     */
    const cfg = loadConfig();
    const hadDeck = cfg.deckFile ? path.basename(cfg.deckFile) : null;
    cfg.pdfFile = target;
    // The PDF becomes the portfolio under review; a stale .pptx path would only
    // describe a different student's work.
    cfg.deckFile = null;

    saveConfig(cfg);
    startWatching();

    log(`PDF received: ${safeName} (${(size / 1048576).toFixed(1)} MB) saved to ${target}`
      + (hadDeck ? ` — now reviewing this instead of ${hadDeck}` : ''));

    return sendJson(res, 200, {
      ok: true,
      file: safeName,
      savedTo: target,
      sizeBytes: size,
      sizeLabel: `${(size / 1048576).toFixed(1)} MB`,
      replacedDeck: hadDeck,
      message: hadDeck
        ? `Now reviewing "${safeName}". "${hadDeck}" is no longer the portfolio in use — add it back with Choose your portfolio if you want it instead.`
        : `Now reviewing "${safeName}".`,
    });
  });
}

async function handle(req, res) {
  const url = new URL(req.url, 'http://localhost');
  const p = url.pathname;

  const access = checkAccess(req, url, accessCode);
  if (!access.ok) return refuseAccess(res, p);
  if (access.setCookie) {
    res.setHeader('Set-Cookie', `${ACCESS_COOKIE}=${encodeURIComponent(accessCode)}; Path=/; HttpOnly; SameSite=Strict; Max-Age=31536000`);
  }
  if (req.method === 'POST' && req.headers['x-nea-coach'] !== '1') {
    return sendJson(res, 403, { ok: false, error: 'Requests that change anything must come from the NEA Coach page itself.' });
  }

  // ---- API
  if (p === '/api/state' && req.method === 'GET') {
    return sendJson(res, 200, buildStateResponse());
  }

  if (p === '/api/progress' && req.method === 'GET') {
    const pct = scanState.progress;
    let eta = null;
    if (scanState.running && scanState.startedAt && pct > 0.08) {
      const elapsed = Date.now() - new Date(scanState.startedAt).getTime();
      const remaining = (elapsed / pct) * (1 - pct);
      if (remaining > 2000 && remaining < 3600000) {
        eta = remaining < 60000
          ? `about ${Math.ceil(remaining / 1000)} seconds`
          : `about ${Math.round(remaining / 60000)} minutes`;
      }
    }
    return sendJson(res, 200, {
      running: scanState.running,
      phase: scanState.phase,
      progress: pct,
      message: scanState.message,
      detail: scanState.detail,
      eta,
      error: scanState.error,
      startedAt: scanState.startedAt,
      slidesChanged: scanState.slidesChanged,
      slidesScanned: scanState.slidesScanned,
    });
  }

  if (p === '/api/scan' && req.method === 'POST') {
    const body = await readBody(req);
    if (scanState.running) {
      return sendJson(res, 409, { ok: false, error: 'A scan is already running.', scanId: scanState.scanId });
    }
    // Kick off in the background and return immediately so the UI stays live.
    requestScan({ force: !!body.force, deep: !!body.deep, thinking: typeof body.thinking === 'boolean' ? body.thinking : null, reason: 'manual' }).catch(() => {});
    await sleep(60);
    return sendJson(res, 200, { ok: true, scanId: scanState.scanId, running: scanState.running });
  }

  if (p === '/api/config' && req.method === 'GET') {
    const cfg = loadConfig();
    // Never transmit the key or the access code.
    return sendJson(res, 200, { ok: true, config: publicConfig(cfg) });
  }

  if (p === '/api/config' && req.method === 'POST') {
    const body = await readBody(req);
    const cfg = loadConfig();

    if (typeof body.deckDir === 'string' && body.deckDir.trim()) {
      const dir = body.deckDir.trim().replace(/^~/, process.env.HOME || '');
      if (!fs.existsSync(dir)) {
        return sendJson(res, 400, { ok: false, error: `That folder does not exist: ${dir}` });
      }
      if (cfg.deckDir !== dir) {
        cfg.deckDir = dir;
        cfg.deckFile = null; // re-discover in the new folder
      }
    }
    if (typeof body.deckFile === 'string') cfg.deckFile = body.deckFile.trim() || null;
    if (typeof body.model === 'string' && body.model.trim()) {
      const wanted = body.model.trim();
      /*
       * Refuse a model Google has retired. Saving it produced a config that
       * silently 404s on every scan, which is indistinguishable from the app
       * being broken.
       */
      if (ai.isRetiredModel(wanted)) {
        return sendJson(res, 400, {
          ok: false,
          error: `Google has retired ${wanted}, so it cannot be used. Pick another model from the list.`,
        });
      }
      cfg.model = wanted;
    }
    if (typeof body.watch === 'boolean') cfg.watch = body.watch;
    // Removing the PDF is how the student turns the visual review back off.
    if (body.clearPdf === true) {
      cfg.pdfFile = null;
      log('PDF removed — the AI review will be text-only until another is supplied.');
    }
    /*
     * Removing the PowerPoint clears the portfolio entirely.
     *
     * There was previously NO way to do this: a .pptx could be chosen but never
     * removed, so starting clean meant choosing another file instead — and the old
     * one stayed silently "in use". Deliberately leaves the PDF alone, since the
     * two are separate files and a PDF on its own is a valid thing to mark.
     */
    if (body.clearDeck === true) {
      cfg.deckFile = null;
      log('PowerPoint removed — nothing is being marked until a portfolio or PDF is chosen.');
    }
    if (typeof body.deep === 'boolean') cfg.recheckAll = body.deep;
    if (typeof body.recheckAll === 'boolean') cfg.recheckAll = body.recheckAll;
    if (typeof body.thinking === 'boolean') cfg.thinking = body.thinking;
    if (typeof body.markEnabled === 'boolean') cfg.markEnabled = body.markEnabled;
    if (typeof body.visionModel === 'string') {
      const wanted = body.visionModel.trim();
      // Refuse a retired model here too, or every marking run 404s.
      if (wanted && ai.isRetiredModel(wanted)) {
        return sendJson(res, 400, { ok: false, error: `Google has retired ${wanted}, so it cannot be used for marking. Pick another model.` });
      }
      cfg.visionModel = wanted;
    }
    if (typeof body.apiKey === 'string') cfg.apiKey = body.apiKey.trim();
    if (typeof body.specId === 'string' && body.specId.trim()) {
      const wanted = body.specId.trim();
      if (!ai.SPEC_PRESET_IDS.includes(wanted)) {
        return sendJson(res, 400, { ok: false, error: `Unknown exam board "${wanted}".` });
      }
      /*
       * Changing the mark scheme invalidates every cached AI judgement: the old
       * coverage and marks were produced against the other board's criteria.
       * Clearing the AI cache forces a real re-read rather than showing Eduqas
       * marks under an AQA heading. Local checks are unaffected and still cached.
       */
      if (cfg.specId !== wanted) {
        cfg.specId = wanted;
        const cache = readJsonSafe(CACHE_PATH, {});
        if (cache && cache.batches) {
          cache.batches = {};
          cache.crossCheck = null;
          writeJsonSafe(CACHE_PATH, cache);
        }
      }
    }

    saveConfig(cfg);
    startWatching();
    return sendJson(res, 200, { ok: true, config: publicConfig(cfg) });
  }

  if (p === '/api/export' && req.method === 'GET') {
    /*
     * Everything worth acting on, as one self-contained HTML file.
     *
     * Built here rather than in the browser because the full finding text does not
     * all reach the client, and because the document must be complete: one file that
     * opens anywhere, prints cleanly, and carries no external references.
     */
    const report = readJsonSafe(LAST_REPORT_PATH, null);
    if (!report || !report.ok) {
      return sendJson(res, 404, { ok: false, error: 'There is no finished check to export yet. Run a check first.' });
    }
    let html;
    try {
      html = exportHtml(report);
    } catch (err) {
      return sendJson(res, 500, { ok: false, error: `Could not build the export: ${err.message}` });
    }
    const base = String((report.deck && report.deck.fileName) || 'portfolio')
      .replace(/\.[a-z0-9]+$/i, '')
      .replace(/[^A-Za-z0-9 _-]/g, '_')
      .slice(0, 60)
      .trim() || 'portfolio';
    const stamp = new Date().toISOString().slice(0, 10);
    res.writeHead(200, {
      'Content-Type': 'text/html; charset=utf-8',
      'Content-Disposition': `attachment; filename="${base} - feedback ${stamp}.html"`,
      'Cache-Control': 'no-store',
      'X-Content-Type-Options': 'nosniff',
    });
    return res.end(html);
  }

  if (p === '/api/test-key' && req.method === 'POST') {
    /*
     * Check a key WITHOUT spending generation quota.
     *
     * This lists the models the key can see, which is a plain GET against the API.
     * It proves the key is valid and tells the student which models they may use,
     * and it costs nothing — which matters, because the whole reason to swap a key
     * is that the quota has run out. Testing a key by making it do work would be
     * absurd in exactly the situation this feature exists for.
     */
    const body = await readBody(req);
    const key = String((body && body.apiKey) || '').trim() || String(loadConfig().apiKey || '').trim();
    if (!key) {
      return sendJson(res, 400, { ok: false, error: 'No key to test. Paste one first.' });
    }
    let models = [];
    try {
      // No counter passed: listModels is a GET and does not consume the daily
      // generateContent quota, so testing a key must not move the meter.
      models = await ai.listModels(key, ai.API_HOST);
    } catch (err) {
      return sendJson(res, 200, { ok: false, error: `Could not reach Google: ${err.message}` });
    }
    if (!models.length) {
      return sendJson(res, 200, {
        ok: false,
        error: 'Google refused that key, or it cannot use the Gemini API. Check it was copied whole, '
          + 'with no spaces, and that the Generative Language API is enabled for its project.',
      });
    }
    /*
     * Report the models in a useful order: the ones the app itself uses first, so
     * the student can see at a glance whether marking will work.
     */
    const names = models.map((m) => m.name);
    const wanted = ['gemini-3.1-flash-lite', 'gemini-3.5-flash', 'gemini-3.8-flash'];
    const present = wanted.filter((w) => names.some((n) => n === w || n.startsWith(w)));
    return sendJson(res, 200, {
      ok: true,
      modelCount: names.length,
      models: present.length ? present : names.slice(0, 8),
      // The tail is a safe way to tell two keys apart without exposing either.
      tail: key.slice(-4),
      note: present.length
        ? 'This key works and can use the models the app needs.'
        : 'This key works, but none of the models the app uses by default were listed. Check the model settings.',
    });
  }

  if (p === '/api/quota' && req.method === 'GET') {
    /*
     * How much of today's free allowance is gone, and when it comes back.
     *
     * A free-tier key spends REQUESTS, not money, so this is the figure that
     * matters: a money meter reads "£0.00" forever and tells the student nothing.
     */
    const cfg = loadConfig();
    return sendJson(res, 200, Object.assign({ ok: true },
      quotaReport(cfg.model, String(cfg.visionModel || '').trim() || null)));
  }

  if (p === '/api/history' && req.method === 'GET') {
    const store = readJsonSafe(HISTORY_PATH, { scans: [] });
    const scans = Array.isArray(store.scans) ? store.scans : [];
    const which = url.searchParams.get('scan');
    // Without ?scan= return the index only: the UI does not need every finding
    // list just to draw a list of past checks.
    if (which == null) {
      return sendJson(res, 200, {
        ok: true,
        limit: HISTORY_LIMIT,
        scans: scans.map((s, i) => ({
          index: i, at: s.at, deck: s.deck, pages: s.pages,
          predictedPercent: s.predictedPercent, findingsTotal: s.findingsTotal,
          changedPages: s.changedPages, costGbp: s.costGbp,
          /*
           * Which model decided the mark. The same portfolio can score differently
           * across models, so without this a change in the mark looks like a change in
           * the work.
           */
          visionModel: s.visionModel || null,
        })),
      });
    }
    const i = Number(which);
    if (!Number.isInteger(i) || i < 0 || i >= scans.length) {
      return sendJson(res, 404, { ok: false, error: 'No scan at that position in the history.' });
    }
    return sendJson(res, 200, { ok: true, scan: scans[i] });
  }

  if (p === '/api/log' && req.method === 'GET') {
    return sendJson(res, 200, { ok: true, log: scanLog.slice(-200) });
  }

  if (p === '/api/models' && req.method === 'GET') {
    const cfg = loadConfig();
    if (!cfg.apiKey) return sendJson(res, 200, { ok: false, error: 'No API key set yet.' });
    try {
      const models = await ai.listModels(cfg.apiKey);
      /*
       * Price each model for THIS deck before the student picks one. The cost of
       * a scan was previously invisible until the bill arrived, which is how a
       * model costing several times the cheapest came to be selected.
       */
      const lastReport = readJsonSafe(LAST_REPORT_PATH, {});
      const pages = Math.max(0, Number(lastReport && lastReport.deck && lastReport.deck.slideCount) || 0);
      /*
       * Quote the price for what the next scan will ACTUALLY do. With a PDF
       * stored, the scan sends page images, so the text-only figure would
       * under-quote the student's bill — the one thing this app must never do.
       */
      const usesVision = !!cfg.pdfFile && fs.existsSync(cfg.pdfFile);
      const visionRuns = ai.visionRunsFor(cfg.markEnabled);
      const priced = models.map((m) => {
        return Object.assign({}, m, {
          supportsThinking: ai.supportsThinkingBudget(m.name),
          /*
           * Google still lists retired models that refuse new API keys with a
           * 404. Flagging them stops the picker offering an option that fails
           * the moment it is chosen.
           */
          retired: ai.isRetiredModel(m.name),
          /*
           * NO PRICE IS SENT.
           *
           * The key is on the free tier, where money is always zero — a price per
           * model told the student nothing and contradicted the rest of the page,
           * which speaks in requests. The client sorts these by the app's own model
           * preference order instead. `estimateScanCostGbp` still exists for a paid
           * key, but nothing here consumes it.
           */
        });
      });

      /*
       * There is deliberately no "switch to a cheaper model" suggestion.
       *
       * It existed to stop a student paying for the priciest model unknowingly. On a
       * free key every model costs the same — nothing — so there is no saving to
       * recommend, and quoting one would be inventing a problem. The real budget is
       * the daily REQUEST allowance, which each model has separately, and the UI
       * explains that where it matters.
       */

      return sendJson(res, 200, {
        ok: true,
        models: priced,
        pages,
        // No `preferred` here on purpose: passing the configured model made this
        // return whatever was already set, so it could never recommend a change.
        suggested: ai.pickModel(models, null),
      });
    } catch (err) {
      return sendJson(res, 200, { ok: false, error: err.message });
    }
  }

  if (p === '/api/upload' && req.method === 'POST') {
    return handleUpload(req, res);
  }

  if (p === '/api/upload-pdf' && req.method === 'POST') {
    return handlePdfUpload(req, res);
  }

  if (p === '/api/finding' && req.method === 'POST') {
    // Marking a finding done is stored in the browser; this endpoint exists so
    // the UI can mirror it if it ever wants to. Nothing is written to the deck.
    return sendJson(res, 200, { ok: true });
  }

  if (p === '/api/connect' && req.method === 'GET') {
    /*
     * Where to reach this Mac from a phone, and the key to type with it.
     *
     * Deliberately REFUSED to anything not on this Mac, even with a valid access
     * code: this is the only endpoint that returns the code, and a device that
     * already has the code does not need it. That keeps the code from ever
     * travelling over the network of its own accord, while solving the problem
     * that actually bites — the address changes when the router reassigns it, and
     * the only record of the new one was the terminal window, which is easy to
     * lose and impossible to read from the phone.
     */
    const fromThisMac = LOOPBACK_ADDRESSES.has(req.socket && req.socket.remoteAddress)
      && LOCAL_HOSTNAMES.has(hostnameOf(req.headers.host));
    if (!fromThisMac) {
      return sendJson(res, 403, {
        ok: false,
        error: 'The phone address is only shown on the Mac itself, because it contains the access code.',
      });
    }
    const cfg = loadConfig();
    const port = Number(process.env.NEA_COACH_PORT || cfg.port || DEFAULT_PORT);
    const hosts = listReachableHosts(port).map((h) => Object.assign({}, h, {
      fullUrl: `${h.url}/?key=${accessCode}`,
    }));
    return sendJson(res, 200, {
      ok: true,
      port,
      hosts,
      // Enough to render the advice without the page needing to guess.
      lanOnly: !hosts.some((h) => /Tailscale/.test(h.label)),
    });
  }

  if (p === '/api/health' && req.method === 'GET') {
    const cfg = loadConfig();
    const found = findDeck(cfg);
    return sendJson(res, 200, {
      ok: true,
      version: VERSION,
      node: process.version,
      platform: `${process.platform}/${process.arch}`,
      deckDir: cfg.deckDir,
      deckFound: !found.error,
      deckPath: found.filePath || null,
      deckError: found.error || null,
      apiKeySet: !!cfg.apiKey,
      running: scanState.running,
    });
  }

  // ---- Static
  if (req.method === 'GET') return serveStatic(req, res, p);

  res.writeHead(405);
  res.end('Method not allowed');
}

/**
 * The addresses this Mac can be reached on, WITHOUT the access code.
 *
 * Used by the in-app "connect your phone" panel, which is served only to this
 * Mac (see /api/connect), so the code is appended there and never travels over
 * the network on its own.
 */
function listReachableHosts(port) {
  const out = [];
  const seen = new Set();
  let ifaces;
  try {
    ifaces = require('os').networkInterfaces();
  } catch (err) {
    return out;
  }
  for (const name of Object.keys(ifaces)) {
    for (const info of ifaces[name] || []) {
      const family = typeof info.family === 'string' ? info.family : `IPv${info.family}`;
      if (family !== 'IPv4' || info.internal || !info.address) continue;
      if (info.address.startsWith('127.')) continue;
      if (seen.has(info.address)) continue;
      seen.add(info.address);
      out.push({
        url: `http://${info.address}:${port}`,
        // 100.x is the Tailscale range: reachable from anywhere, not just at home.
        label: info.address.startsWith('100.') ? 'Tailscale — works anywhere' : `Same Wi-Fi (${name})`,
        interface: name,
      });
    }
  }
  return out;
}

/**
 * Every address this machine can be reached on, so the launcher can tell the
 * student exactly what to type into their phone. Loopback is excluded because
 * a phone cannot use it.
 */
function listReachableUrls(port, code) {
  return listReachableHosts(port).map((h) => `${h.url}/?key=${code}   ${h.label}`);
}

/* ------------------------------------------------------------------ *
 * Start
 * ------------------------------------------------------------------ */

function start() {
  fs.mkdirSync(STATE_DIR, { recursive: true });
  fs.mkdirSync(SPEC_DIR, { recursive: true });

  const cfg = loadConfig();
  // Ensure a config file exists so the user has something to look at.
  if (!fs.existsSync(CONFIG_PATH)) saveConfig(cfg);

  /*
   * Keep one access code across restarts, so a phone's saved Home Screen link
   * keeps working. Saved from the file's own contents rather than cfg, so an
   * API key supplied by environment variable is never written to disk.
   */
  if (!ACCESS_CODE_SHAPE.test(String(cfg.accessCode || ''))) {
    cfg.accessCode = newAccessCode();
    saveConfig(Object.assign(defaultConfig(), readJsonSafe(CONFIG_PATH, {}), { accessCode: cfg.accessCode }));
  }
  accessCode = cfg.accessCode;

  lastReport = readJsonSafe(LAST_REPORT_PATH, null);

  const port = Number(process.env.NEA_COACH_PORT || cfg.port || DEFAULT_PORT);
  const host = process.env.NEA_COACH_HOST || cfg.host || '127.0.0.1';

  const server = http.createServer((req, res) => {
    handle(req, res).catch((err) => {
      log(`Request error: ${err.message}`);
      if (!res.headersSent) sendJson(res, 500, { ok: false, error: err.message });
      else res.end();
    });
  });

  server.on('error', (err) => {
    if (err.code === 'EADDRINUSE') {
      console.error(`\n  Port ${port} is already in use.`);
      console.error(`  NEA Coach may already be running — try opening http://${host}:${port}\n`);
    } else {
      console.error(`\n  Could not start the server: ${err.message}\n`);
    }
    process.exit(1);
  });

  server.listen(port, host, () => {
    const url = `http://${host}:${port}`;
    console.log('');
    console.log('  ┌──────────────────────────────────────────────┐');
    console.log('  │  NEA Coach is running                        │');
    console.log('  └──────────────────────────────────────────────┘');
    console.log('');
    // Show a clickable loopback address for the Mac itself, even when the
    // server is bound to all interfaces so a phone can reach it.
    const macUrl = host === '0.0.0.0' || host === '::' ? `http://127.0.0.1:${port}` : url;
    console.log(`   On this Mac:  ${macUrl}`);

    // Print the addresses a phone can use. This is the part that makes the app
    // usable from an iPhone, so it should not be buried in a config file.
    const phoneUrls = host === '127.0.0.1' || host === 'localhost' ? [] : listReachableUrls(port, accessCode);
    if (phoneUrls.length) {
      console.log('');
      console.log('   On your phone (same Wi-Fi, or Tailscale from anywhere):');
      for (const u of phoneUrls) console.log(`     ${u}`);
      console.log('   The key in that address is the password. Do not share it.');
    }
    console.log('');
    const found = findDeck(cfg);
    if (found.error) console.log(`   ! ${found.error}`);
    else console.log(`   Portfolio: ${found.filePath}`);
    console.log(`   API key:   ${cfg.apiKey ? 'set' : 'not set yet (free local checks only)'}`);
    console.log(`   Read-only: your .pptx is never modified.`);
    console.log('');
    console.log('   Leave this window open. Press Control-C to quit.');
    console.log('');

    startWatching();
  });

  // Shut down cleanly when the launcher window closes or the user presses Ctrl-C.
  const shutdown = () => {
    log('Shutting down.');
    if (watcher) {
      try {
        watcher.close();
      } catch (err) {
        /* ignore */
      }
    }
    server.close(() => process.exit(0));
    setTimeout(() => process.exit(0), 1500);
  };
  process.on('SIGINT', shutdown);
  process.on('SIGTERM', shutdown);

  // An unhandled error must never leave the student staring at a dead page.
  process.on('uncaughtException', (err) => {
    log(`Uncaught exception: ${err.message}`);
    console.error(err);
  });
}

if (require.main === module) {
  start();
}

module.exports = { start, performScan, buildStateResponse, findDeck, loadConfig, publicConfig, checkAccess, newAccessCode, handle, compareReports, findingKey, matchAiFindings, snapshotOf, appendHistory, quotaDate, quotaReset, readQuota, noteRequest, quotaReport, learnQuotaLimit, quotaLimitFromError, buildStats, exportHtml, escapeHtml };
