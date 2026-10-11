/*
 * Runs the real page script against a DOM stub and drives the first-run wizard.
 * This is an ADAPTER, not a second implementation: the script under test is the
 * exact text of the <script> block in public/index.html. It exists because the
 * wizard's failure modes are runtime ones (a missing element, a bad property)
 * that no amount of reading catches, and I have no browser.
 */
const fs = require('fs');
const vm = require('vm');
const html = fs.readFileSync(process.argv[2], 'utf8');
let script = html.match(/<script>([\s\S]*)<\/script>/)[1];

const ids = [...new Set([...html.matchAll(/\bid="([^"]+)"/g)].map(m => m[1]))];

function makeEl(id) {
  const el = {
    id, tagName: 'DIV', className: '', checked: false, disabled: false,
    _hidden: false, _display: '',
    value: '', textContent: '', innerHTML: '', nodeValue: '', children: [],
    files: null, open: false, dataset: {}, nodeType: 1,
    classList: {
      _s: new Set(),
      add(c) { this._s.add(c); }, remove(c) { this._s.delete(c); },
      contains(c) { return this._s.has(c); },
      toggle(c) { if (this._s.has(c)) { this._s.delete(c); } else { this._s.add(c); } },
    },
    appendChild(c) { this.children.push(c); return c; },
    removeChild(c) { this.children = this.children.filter(x => x !== c); },
    insertBefore(c) { this.children.unshift(c); return c; },
    replaceChildren() { this.children = []; },
    setAttribute(k, v) { this[k] = v; }, getAttribute(k) { return this[k]; },
    removeAttribute(k) { delete this[k]; }, hasAttribute(k) { return k in this; },
    matches() { return false; },
    cloneNode() { return makeEl(this.id); },
    get hidden() { return this._hidden || this._display === 'none'; },
    set hidden(v) { this._hidden = !!v; },
    get style() { var self = this; if (!this._style) { this._style = new Proxy({}, {
      get: function (t, k) { return k === 'display' ? self._display : (t[k] || ''); },
      set: function (t, k, v) { if (k === 'display') { self._display = v; } else { t[k] = v; } return true; },
    }); } return this._style; },
    get firstChild() { return this.children[0] || null; },
    get lastChild() { return this.children[this.children.length - 1] || null; },
    get innerText() { return this.textContent; },
    set innerText(v) { this.textContent = v; },
    addEventListener() {}, removeEventListener() {}, focus() {}, blur() {},
    querySelector() { return null; }, querySelectorAll() { return []; },
    getElementsByTagName() { return []; }, getElementsByClassName() { return []; },
    contains() { return false; }, closest() { return null; },
    scrollIntoView() {}, click() {}, remove() {}, scrollTo() {},
    getBoundingClientRect() { return { top: 0, left: 0, width: 100, height: 20, bottom: 20, right: 100 }; },
  };
  return el;
}

const store = {};
const els = new Map();
ids.forEach(id => els.set(id, makeEl(id)));

const document = {
  readyState: 'loading',
  documentElement: makeEl('html'),
  body: makeEl('body'),
  head: makeEl('head'),
  activeElement: null,
  _handlers: {},
  getElementById(id) { return els.get(id) || null; },
  querySelector(sel) { return sel === '.intro-card' ? makeEl('card') : null; },
  querySelectorAll() { return []; },
  createElement(tag) { const e = makeEl(''); e.tagName = String(tag).toUpperCase(); return e; },
  createTextNode(t) { const e = makeEl(''); e.nodeValue = t; e.textContent = t; e.nodeType = 3; return e; },
  createComment() { return makeEl(''); },
  createDocumentFragment() { return makeEl('frag'); },
  addEventListener(type, fn) { this._handlers[type] = fn; },
  removeEventListener() {},
  elementFromPoint() { return null; },
  cookie: '',
};

function xhrStub() {
  this.open = function () {}; this.setRequestHeader = function () {};
  this.send = function () {}; this.upload = {}; this.status = 0; this.responseText = '{}';
}

const window = {
  document,
  localStorage: {
    _d: store,
    getItem(k) { return Object.prototype.hasOwnProperty.call(this._d, k) ? this._d[k] : null; },
    setItem(k, v) { this._d[k] = String(v); },
    removeItem(k) { delete this._d[k]; },
  },
  addEventListener() {}, removeEventListener() {},
  matchMedia() { return { matches: false, addEventListener() {} }; },
  location: { href: 'http://127.0.0.1:4174/', protocol: 'http:', hostname: '127.0.0.1', port: '4174', search: '' },
  navigator: { userAgent: 'node', language: 'en-GB', clipboard: { writeText() {} } },
  console, setTimeout, clearTimeout, setInterval, clearInterval,
  fetch() { return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ ok: true, config: {}, scan: {} }) }); },
  XMLHttpRequest: xhrStub,
  requestAnimationFrame(fn) { return setTimeout(fn, 0); },
  getComputedStyle() { return {}; },
  history: { replaceState() {} },
};
window.window = window;

const ctx = {
  window, document, console, setTimeout, clearTimeout, setInterval, clearInterval,
  localStorage: window.localStorage, fetch: window.fetch, XMLHttpRequest: xhrStub,
  requestAnimationFrame: window.requestAnimationFrame, navigator: window.navigator,
  location: window.location, getComputedStyle: window.getComputedStyle,
  history: window.history,
};
ctx.globalThis = ctx;
vm.createContext(ctx);

// Expose the internals for driving, by injecting inside the IIFE (so the names
// resolve) immediately before it closes.
const names = ['state', 'init', 'maybeShowIntro', 'introShow', 'introIsDone', 'introRenderDone', 'introBoardLabel', 'ui', 'introNext', 'introClose', 'introSaveBoard'];
const marker = '}());';
const at = script.lastIndexOf(marker);
if (at < 0) { throw new Error('could not find the IIFE close'); }
const expose = '\n  document.documentElement.__t = {' + names.map(n => n + ': ' + n).join(', ') + '};\n';
script = script.slice(0, at) + expose + script.slice(at);

const errors = [];
try {
  vm.runInContext(script, ctx, { filename: 'nea-page.js' });
} catch (e) {
  errors.push('top-level: ' + e.message + '\n' + (e.stack || '').split('\n').slice(0, 6).join('\n'));
}

const T = ctx.document.documentElement.__t;
if (!T) { console.log('INTERNALS NOT EXPOSED — aborting'); process.exit(1); }

function call(name, ...args) {
  try { return T[name](...args); }
  catch (e) { errors.push(name + ' -> ' + e.message); return undefined; }
}

var CFG = {
  specId: 'aqa-gcse',
  specPresets: [
    { id: 'eduqas-alevel', label: 'WJEC Eduqas A-level Design and Technology (NEA)' },
    { id: 'aqa-gcse', label: 'AQA GCSE Design and Technology (8552)' },
  ],
  pdfName: 'DT - Google Slides.pdf',
  apiKeySet: true,
};

// The page script defers init to DOMContentLoaded, exactly as a browser does.
call('init');
if (document._handlers.DOMContentLoaded) { document._handlers.DOMContentLoaded(); }
console.log('init ran. errors:', errors.length ? errors : 'none');

(async function () {
  var wait = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
  var stepNow = function () {
    for (var k = 0; k < 5; k++) { if (!els.get('introStep' + k).hidden) { return k; } }
    return -1;
  };
  var state = T.state;
  state.data = { config: CFG };
  call('maybeShowIntro');
  console.log('overlay open:', !els.get('introOverlay').hidden, '| step', stepNow(), '| next label:', els.get('introNext').textContent);
  call('introShow', 1);
  var cards = els.get('introBoards').children;
  console.log('board cards when step 1 is shown:', cards.length);
  cards.forEach(function (c) {
    console.log('   -', c.children.map(function (k) { return k.textContent; }).join(' | '));
  });
  call('introShow', 0);
  for (var i = 0; i < 7; i++) {
    if (els.get('introOverlay').hidden) { console.log('CLOSED after ' + i + ' presses'); break; }
    var label = els.get('introNext').textContent;
    call('introNext');
    await wait(40);
    console.log('press ' + (i + 1) + ' ["' + label + '"] -> step ' + stepNow()
      + ' overlay ' + (els.get('introOverlay').hidden ? 'CLOSED' : 'open')
      + ' next="' + els.get('introNext').textContent + '"');
  }
  console.log('summary rows on the last step:', els.get('introDoneList').children.length);
  console.log('\nERRORS:', errors.length ? errors : 'NONE');
  process.exit(errors.length ? 1 : 0);
}());
