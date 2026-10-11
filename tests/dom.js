/*
 * A tiny DOM stand-in, so the page's real script can be run in Node.
 *
 * WHY THIS EXISTS
 *
 * `public/index.html` is one big IIFE of render functions, and they are only ever
 * exercised in a browser. A single undefined property in `renderFixList` means a
 * blank page for every user, and nothing in this test suite would have caught it.
 * This lets the tests feed a real report into the REAL script and check that the
 * render functions run and produce output.
 *
 * It is not a browser. It models only what the page actually uses, and the
 * behaviours that bit us are modelled deliberately:
 *
 *   - `hidden` is a reflected attribute whose effect is `display:none`, and an
 *     inline `style.display` overrides it. Getting this wrong produced a FALSE
 *     FAILURE once, where a working button looked broken.
 *   - `click()` fires the handlers registered through `addEventListener`, so
 *     wiring is tested rather than just the function bodies.
 *   - `select.options` is a real array, because the page builds dropdowns by
 *     appending `<option>` elements and then reads `.options`.
 *
 * Anything not modelled returns a harmless stub, so an unmodelled call is a
 * missing assertion rather than a crash that hides the test result.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

/** Build one fake element. */
function makeEl(id) {
  const handlers = {};
  const el = {
    id,
    tagName: 'DIV',
    className: '',
    checked: false,
    disabled: false,
    value: '',
    textContent: '',
    innerHTML: '',
    nodeValue: '',
    children: [],
    options: [],
    files: null,
    open: false,
    dataset: {},
    nodeType: 1,
    _hidden: false,
    _display: '',
    classList: {
      _s: new Set(),
      add(c) { this._s.add(c); },
      remove(c) { this._s.delete(c); },
      contains(c) { return this._s.has(c); },
      toggle(c) { if (this._s.has(c)) { this._s.delete(c); } else { this._s.add(c); } },
    },
    appendChild(c) { this.children.push(c); return c; },
    removeChild(c) { this.children = this.children.filter((x) => x !== c); },
    insertBefore(c) { this.children.unshift(c); return c; },
    replaceChildren() { this.children = []; },
    setAttribute(k, v) {
      /*
       * `style` is a getter here (an inline-style object), so a naive assignment
       * throws. The page's own `el()` helper sets inline styles this way, so it
       * has to work or every render that styles anything reports a false failure.
       */
      if (k === 'style') {
        String(v).split(';').forEach((decl) => {
          const i = decl.indexOf(':');
          if (i < 0) { return; }
          const prop = decl.slice(0, i).trim();
          const val = decl.slice(i + 1).trim();
          if (!prop) { return; }
          if (prop === 'display') { this._display = val; }
          else { this.style[prop] = val; }
        });
        return;
      }
      this[k] = v;
    },
    getAttribute(k) { return this[k]; },
    removeAttribute(k) { delete this[k]; },
    hasAttribute(k) { return k in this; },
    matches() { return false; },
    cloneNode() { return makeEl(this.id); },
    get firstChild() { return this.children[0] || null; },
    get lastChild() { return this.children[this.children.length - 1] || null; },
    get innerText() { return this.textContent; },
    set innerText(v) { this.textContent = v; },
    get hidden() { return this._hidden || this._display === 'none'; },
    set hidden(v) {
      this._hidden = !!v;
      if (v) { this._display = 'none'; } else if (this._display === 'none') { this._display = ''; }
    },
    get style() {
      const self = this;
      if (!this._style) {
        /*
         * CSS custom properties go through setProperty()/getPropertyValue(), not
         * through property assignment, so the proxy has to provide them or every
         * spark in the feel layer reports a false failure.
         */
        this._style = new Proxy({}, {
          get: (t, k) => {
            if (k === 'setProperty') {
              return (name, value) => { if (name === 'display') { self._display = value; } else { t[name] = value; } };
            }
            if (k === 'getPropertyValue') { return (name) => (t[name] === undefined ? '' : String(t[name])); }
            if (k === 'removeProperty') { return (name) => { delete t[name]; }; }
            if (k === 'cssText') { return Object.keys(t).map((n) => n + ': ' + t[n]).join('; '); }
            if (k === 'display') { return self._display; }
            return t[k] || '';
          },
          set: (t, k, v) => { if (k === 'display') { self._display = v; } else { t[k] = v; } return true; },
        });
      }
      return this._style;
    },
    addEventListener(type, fn) { (handlers[type] = handlers[type] || []).push(fn); },
    removeEventListener() {},
    fire(type, ev) {
      (handlers[type] || []).forEach((fn) => fn(ev || { preventDefault() {}, stopPropagation() {}, target: el }));
    },
    click() { this.fire('click'); },
    change() { this.fire('change'); },
    focus() {}, blur() {}, scrollIntoView() {}, remove() {}, scrollTo() {},
    querySelector() { return null; },
    querySelectorAll() { return []; },
    getElementsByTagName() { return []; },
    getElementsByClassName() { return []; },
    contains() { return false; },
    closest() { return null; },
    getBoundingClientRect() { return { top: 0, left: 0, width: 100, height: 20, bottom: 20, right: 100 }; },
  };
  return el;
}

/**
 * Load the page script with a DOM around it.
 *
 * The app root is derived from this file's own location, NOT from the working
 * directory, so the suite still runs when invoked as `node <app>/tests/test.js`
 * from somewhere else.
 *
 * @param {string} [appRoot] override; defaults to the folder containing this file
 * @returns {{els, store, errors, call, state, ui, document, window, html}}
 */
function loadPage(appRoot) {
  const root = appRoot || path.join(__dirname, '..');
  const html = fs.readFileSync(path.join(root, 'public', 'index.html'), 'utf8');
  let script = html.match(/<script>([\s\S]*)<\/script>/)[1];

  const ids = [...new Set([...html.matchAll(/\bid="([^"]+)"/g)].map((m) => m[1]))];
  const els = new Map();
  /*
   * Tag name AND attributes are read from the markup. Every element used to be a bare
   * DIV with no attributes at all, so no test could tell a real <button type="button">,
   * with a class and an accessible name, from a <div> wearing a click handler. "Can a
   * keyboard reach this, and does it say what it does?" was therefore untestable, and
   * that is the kind of fault this file exists to catch. Bare attributes such as
   * `hidden` are deliberately skipped: they are booleans, and assigning them a string
   * would make every one of them true.
   */
  const markupFor = (id) => {
    const m = new RegExp('<(\\w+)([^>]*\\bid="' + id + '"[^>]*)>', 'i').exec(html);
    if (!m) { return null; }
    const attrs = {};
    for (const a of m[2].matchAll(/([\w-]+)\s*=\s*"([^"]*)"/g)) { attrs[a[1]] = a[2]; }
    return { tag: m[1].toUpperCase(), attrs };
  };
  ids.forEach((id) => {
    const element = makeEl(id);
    const markup = markupFor(id);
    if (markup) {
      element.tagName = markup.tag;
      for (const [name, value] of Object.entries(markup.attrs)) {
        if (name === 'class') {
          element.className = value;
          String(value).split(/\s+/).forEach((c) => { if (c) { element.classList.add(c); } });
        } else if (name === 'style') {
          element.setAttribute('style', value);
        } else {
          element[name] = value;
        }
      }
    }
    els.set(id, element);
  });

  const store = {};
  const document = {
    // A browser runs this while the document is still loading, so init() must be
    // deferred to DOMContentLoaded. Modelling it as 'complete' hides real errors.
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

  function XhrStub() {
    this.open = function () {};
    this.setRequestHeader = function () {};
    this.send = function () {};
    this.upload = {};
    this.status = 0;
    this.responseText = '{}';
  }

  const fetchCalls = [];
  const window = {
    document,
    localStorage: {
      _d: store,
      getItem(k) { return Object.prototype.hasOwnProperty.call(this._d, k) ? this._d[k] : null; },
      setItem(k, v) { this._d[k] = String(v); },
      removeItem(k) { delete this._d[k]; },
    },
    /*
     * Viewport and scroll, which the page reads for the back-to-top threshold.
     * Recorded rather than performed: the page asks for scrollTo({top: 0, behavior: ...})
     * and a test needs to see both the target and whether it wanted smooth or instant.
     */
    innerWidth: 1024,
    innerHeight: 768,
    pageYOffset: 0,
    scrollCalls: [],
    scrollTo(...args) { window.scrollCalls.push(args); },
    addEventListener() {}, removeEventListener() {},
    matchMedia() { return { matches: false, addEventListener() {} }; },
    location: { href: 'http://127.0.0.1:4173/', protocol: 'http:', hostname: '127.0.0.1', port: '4173', search: '' },
    navigator: { userAgent: 'node', language: 'en-GB', clipboard: { writeText() {} } },
    console,
    setTimeout, clearTimeout, setInterval, clearInterval,
    fetch(url) {
      fetchCalls.push(String(url));
      return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ ok: true, config: {}, scan: {} }) });
    },
    XMLHttpRequest: XhrStub,
    requestAnimationFrame(fn) { return setTimeout(fn, 0); },
    getComputedStyle() { return {}; },
    history: { replaceState() {} },
  };
  window.window = window;

  const ctx = {
    window, document, console,
    setTimeout, clearTimeout, setInterval, clearInterval,
    localStorage: window.localStorage,
    fetch: window.fetch,
    XMLHttpRequest: XhrStub,
    requestAnimationFrame: window.requestAnimationFrame,
    navigator: window.navigator,
    location: window.location,
    getComputedStyle: window.getComputedStyle,
    history: window.history,
  };
  ctx.globalThis = ctx;
  vm.createContext(ctx);

  // Expose the internals by injecting an assignment INSIDE the IIFE, immediately
  // before it closes, so the names resolve.
  const names = [
    'state', 'ui', 'init', 'renderAll', 'normaliseFinding', 'verdictFor', 'allFindings',
    'maybeShowIntro', 'introShow', 'introIsDone', 'introClose', 'loadState',
    // Render functions, so a test can drive one in isolation and name the one that
    // throws instead of only knowing that renderAll failed.
    'renderBanners', 'renderHeader', 'renderStatus', 'renderProgress', 'renderActions',
    'renderCost', 'renderScanCostHint', 'renderKeySection', 'renderCheckCoverage',
    'renderTiles', 'renderControls', 'renderTabs', 'renderAiNotice', 'renderStrengths',
    'renderFixList', 'renderSuggestions', 'renderCoverage', 'renderComparison',
    'renderPredicted', 'renderSlides', 'renderSettings', 'renderStats', 'renderReviewed',
    'renderUploadSection', 'renderModelCostPreview', 'renderPortfolioNotice',
    // The feel layer, so a test can fire a ripple or a burst and inspect the DOM
    // it produces instead of only being able to see that the page loaded.
    'spawnRipple', 'celebrate', 'rollFigure', 'rolledFigure', 'sparkLayer',
    'renderToTop', 'wireToTop',
  ];
  const marker = '}());';
  const at = script.lastIndexOf(marker);
  if (at < 0) { throw new Error('could not find the page IIFE close'); }
  const expose = '\n  document.documentElement.__t = {'
    + names.map((n) => n + ': ' + n).join(', ')
    + ', getStep: function () { return introStep; }'
    + ', getCalm: function () { return calmMotion; }'
    + ', setCalm: function (v) { calmMotion = !!v; } };\n';
  script = script.slice(0, at) + expose + script.slice(at);

  const errors = [];
  try {
    vm.runInContext(script, ctx, { filename: 'nea-page.js' });
  } catch (e) {
    errors.push('top-level: ' + e.message);
  }

  const T = document.documentElement.__t || {};
  if (T.init && document._handlers.DOMContentLoaded) {
    try { document._handlers.DOMContentLoaded(); }
    catch (e) { errors.push('init: ' + e.message); }
  }

  /** Call an exposed page function, capturing any throw instead of stopping. */
  function call(name, ...args) {
    try {
      if (typeof T[name] !== 'function') { return { ok: false, error: name + ' is not exposed' }; }
      return { ok: true, value: T[name](...args) };
    } catch (e) {
      return { ok: false, error: name + ' threw: ' + e.message };
    }
  }

  return { els, store, errors, call, state: T.state, ui: T.ui, document, window, html, fetchCalls, T };
}

module.exports = { loadPage, makeEl };
