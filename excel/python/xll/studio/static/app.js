'use strict';
// ASF Python Studio: the page. The server (server.py) holds the files, the checks, the console and
// pip; Excel's commands go through it too.

const $ = id => document.getElementById(id);
const dark = matchMedia('(prefers-color-scheme: dark)');
const state = { tabs: [], active: null, folder: '', excel: null, jobsRevision: -1, packages: [], history: [], historyAt: 0 };

async function api(name, body) {
  const options = body === undefined ? {} :
    { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Studio': '1' }, body: JSON.stringify(body) };
  const r = await fetch('/api/' + name, options);
  const data = await r.json().catch(() => ({ error: 'The studio closed: open it again from Excel.' }));
  if (!r.ok) throw new Error(data.error || r.statusText);
  return data;
}
const get = (name, args) => api(name + (args ? '?' + new URLSearchParams(args) : ''));
const post = (name, body) => api(name, body || {});

function el(tag, attrs, ...children) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (k === 'class') e.className = v;
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else e.setAttribute(k, v);
  }
  for (const c of children) if (c !== null && c !== undefined) e.append(c);
  return e;
}
function foot(text) { $('foot-left').textContent = text; }

// --- dialogs: in the page, since an app window may not show the browser's own ------------------
// ask({ title, text, fields: [{ name, label, value, hint }], checks: [{ name, label, value }], ok, cancel, pre })
// resolves to the fields' and checks' values on OK, to null on Cancel or Esc
function ask(o) {
  return new Promise(resolve => {
    const d = el('dialog', { class: 'modal', 'aria-label': o.title || 'ASF Python Studio' });
    const inputs = (o.fields || []).map(f => el('input', { name: f.name, value: f.value ?? '', spellcheck: 'false', autocomplete: 'off' }));
    const checks = (o.checks || []).map(c => { const i = el('input', { type: 'checkbox', name: c.name }); i.checked = !!c.value; return i; });
    let answer = null;
    const form = el('form', { method: 'dialog' },
      o.title ? el('h2', {}, o.title) : null,
      o.text ? el(o.pre ? 'pre' : 'p', { class: o.pre ? 'modal-pre' : 'modal-text' }, o.text) : null,
      ...(o.fields || []).map((f, k) => el('label', { class: 'field' }, el('span', {}, f.label || f.name), inputs[k],
        f.hint ? el('small', { class: 'muted' }, f.hint) : null)),
      ...(o.checks || []).map((c, k) => el('label', { class: 'check' }, checks[k], ' ', c.label)),
      el('div', { class: 'modal-buttons' },
        o.cancel === null ? null : el('button', { type: 'button', onclick: () => d.close() }, o.cancel || 'Cancel'),
        el('button', { type: 'submit', class: o.danger ? 'danger' : 'primary' }, o.ok || 'OK')));
    form.addEventListener('submit', () => {
      answer = {};
      (o.fields || []).forEach((f, k) => { answer[f.name] = inputs[k].value; });
      (o.checks || []).forEach((c, k) => { answer[c.name] = checks[k].checked; });
    });
    d.append(form);
    d.addEventListener('close', () => { d.remove(); resolve(answer); });
    document.body.append(d);
    d.showModal();
    (inputs[0] || form.querySelector('button[type=submit]')).focus();
    if (inputs[0]) inputs[0].select();
  });
}
const message = (text, title, pre) => ask({ title, text, pre, ok: 'OK', cancel: null });
const sure = async (text, ok, title) => (await ask({ title, text, ok: ok || 'OK', danger: /remove|delete|close/i.test(ok || '') })) !== null;
async function input(text, value, title) {
  const r = await ask({ title, text, fields: [{ name: 'value', label: '', value }] });
  return r ? r.value : null;
}

// --- output panel ---------------------------------------------------------------------------
function show(panel) {
  document.querySelectorAll('.ptab').forEach(b => b.classList.toggle('active', b.dataset.panel === panel));
  document.querySelectorAll('.panel-body').forEach(b => b.classList.toggle('hidden', b.dataset.body !== panel));
  if (panel === 'console') $('console-in').focus();
}
function output(text, kind) {
  const out = $('output');
  out.append(el('span', { class: kind ? 'out-' + kind : '' }, text.endsWith('\n') ? text : text + '\n'));
  out.scrollTop = out.scrollHeight;
}

// --- editor ---------------------------------------------------------------------------------
const editor = CodeMirror($('editor'), {
  value: '', mode: 'python', lineNumbers: true, indentUnit: 4, tabSize: 4, matchBrackets: true,
  autoCloseBrackets: true, styleActiveLine: true, theme: dark.matches ? 'material-darker' : 'default',
  gutters: ['problem-gutter', 'CodeMirror-linenumbers'], readOnly: 'nocursor',
  extraKeys: {
    'Ctrl-S': () => save(), 'Cmd-S': () => save(), 'F5': () => run(), 'Ctrl-Space': cm => complete(cm),
    'Ctrl-/': cm => cm.toggleComment(), 'Tab': cm => cm.somethingSelected() ? cm.indentSelection('add') : cm.execCommand('insertSoftTab'),
    'Shift-Tab': cm => cm.indentSelection('subtract'),
  },
});
dark.addEventListener('change', e => editor.setOption('theme', e.matches ? 'material-darker' : 'default'));

function tabOf(path) { return state.tabs.find(t => t.path === path); }

async function open(path, line) {
  let tab = tabOf(path);
  if (!tab) {
    const r = await get('file', { path });
    tab = { path, doc: CodeMirror.Doc(r.text, 'python'), saved: r.text, marks: [], problems: [] };
    tab.doc.on('change', () => { renderTabs(); lintSoon(tab); });
    state.tabs.push(tab);
  }
  activate(tab);
  if (line) { editor.setCursor({ line: line - 1, ch: 0 }); editor.scrollIntoView(null, 120); }
  editor.focus();
}

function activate(tab) {
  state.active = tab;
  $('empty').classList.toggle('hidden', !!tab);
  if (tab) { editor.swapDoc(tab.doc); editor.setOption('readOnly', false); }
  else { editor.swapDoc(CodeMirror.Doc('', 'python')); editor.setOption('readOnly', 'nocursor'); }
  renderTabs(); renderFiles(); renderProblems();
  editor.refresh();
}

function renderTabs() {
  const tabs = $('tabs');
  tabs.replaceChildren(...state.tabs.map(t => el('div', {
    class: 'tab' + (t === state.active ? ' active' : '') + (t.doc.getValue() !== t.saved ? ' dirty' : ''),
    title: t.path, onclick: () => activate(t),
  }, el('span', { class: 'name' }, t.path.split('/').pop()),
     el('button', { class: 'x', title: 'Close', onclick: e => { e.stopPropagation(); close(t); } }, '✕'))));
}

async function close(tab) {
  if (tab.doc.getValue() !== tab.saved && !await sure(tab.path + ' has changes that are not saved.', 'Close without saving')) return;
  const i = state.tabs.indexOf(tab);
  state.tabs.splice(i, 1);
  if (state.active === tab) activate(state.tabs[Math.min(i, state.tabs.length - 1)] || null);
  else renderTabs();
}

async function save() {
  const tab = state.active;
  if (!tab) return;
  const text = tab.doc.getValue();
  try {
    const r = await post('save', { path: tab.path, text });
    tab.saved = text;
    setProblems(tab, r.problems);
    renderTabs();
    foot('Saved ' + tab.path + (r.problems.length ? ' — ' + r.problems.length + ' problem(s)' : ''));
    if ($('reload-on-save').checked && state.excel && !tab.path.includes('/')) await reload(true);
    refreshOutline();
  } catch (e) { output('Save failed: ' + e.message, 'err'); show('output'); }
}

async function run() {
  const tab = state.active;
  if (!tab) return;
  if (tab.doc.getValue() !== tab.saved) await save();
  show('output');
  output('▶ ' + tab.path + ' in Excel', 'head');
  foot('Running ' + tab.path + ' in Excel…');
  try {
    const r = await post('excel/run', { path: tab.path });
    output(r.output || '(nothing printed)');
    foot('Ran ' + tab.path);
  } catch (e) { output(e.message, 'err'); foot('Not run'); }
}

async function reload(quiet) {
  foot('Excel imports the Python folder again…');
  try {
    await post('excel/reload');
    if (!quiet) { show('output'); }
    output('Excel imported the Python folder again and calculated the workbooks.', 'ok');
    foot('Reloaded in Excel');
  } catch (e) { output('Reload: ' + e.message, 'err'); show('output'); foot('Not reloaded'); }
}

// --- completions and signatures ---------------------------------------------------------------
function complete(cm) {
  if (!state.active) return;
  cm.showHint({ hint: pythonHint, completeSingle: false });
}
function pythonHint(cm, callback) {
  const cur = cm.getCursor(), line = cm.getLine(cur.line);
  let start = cur.ch;
  while (start && /\w/.test(line.charAt(start - 1))) start--;
  post('complete', { text: cm.getValue(), line: cur.line + 1, column: cur.ch, path: state.active.path }).then(r => {
    callback({
      list: r.completions.map(c => ({ text: c.text, displayText: c.text + '  · ' + c.type })),
      from: CodeMirror.Pos(cur.line, start), to: cur,
    });
  }).catch(() => callback(null));
}
pythonHint.async = true;

let signatureTimer;
editor.on('inputRead', (cm, change) => {
  const typed = change.text.join('');
  if (typed === '.' || /^[A-Za-z_]$/.test(typed) && /\.\w$/.test(cm.getLine(change.from.line).slice(0, change.from.ch + 1))) complete(cm);
  if (typed === '(' || typed === ',') signature();
  if (typed === ')') $('signature').textContent = '';
});
function signature() {
  clearTimeout(signatureTimer);
  signatureTimer = setTimeout(async () => {
    if (!state.active) return;
    const cur = editor.getCursor();
    const r = await post('complete', { text: editor.getValue(), line: cur.line + 1, column: cur.ch, path: state.active.path }).catch(() => null);
    $('signature').textContent = r && r.signatures.length ? r.signatures[0].label : '';
  }, 150);
}

// --- problems ----------------------------------------------------------------------------------
const lintTimers = new WeakMap();
function lintSoon(tab) {
  clearTimeout(lintTimers.get(tab));
  lintTimers.set(tab, setTimeout(async () => {
    const r = await post('lint', { text: tab.doc.getValue() }).catch(() => null);
    if (r) setProblems(tab, r);
  }, 600));
}
function setProblems(tab, problems) {
  tab.marks.forEach(m => m.clear());
  tab.marks = [];
  tab.problems = problems;
  if (tab === state.active) editor.clearGutter('problem-gutter');
  for (const p of problems) {
    const line = Math.max(p.line - 1, 0), text = tab.doc.getLine(line) || '';
    let end = p.col;
    while (end < text.length && /\w/.test(text[end])) end++;
    if (end === p.col) end = Math.min(p.col + 1, text.length);
    tab.marks.push(tab.doc.markText({ line, ch: p.col }, { line, ch: end }, { className: 'cm-problem-' + p.severity, title: p.message }));
    if (tab === state.active) {
      editor.setGutterMarker(line, 'problem-gutter',
        el('span', { class: 'gutter-mark sev-' + p.severity, title: p.message }, p.severity === 'error' ? '●' : '▲'));
    }
  }
  renderProblems();
}
function renderProblems() {
  const tab = state.active, problems = tab ? tab.problems : [];
  $('problem-count').textContent = problems.length || '';
  $('problems').replaceChildren(...problems.map(p => el('li', {
    class: 'problem', onclick: () => { editor.setCursor({ line: p.line - 1, ch: p.col }); editor.focus(); },
  }, el('span', { class: 'sev sev-' + p.severity }, p.severity), el('span', {}, p.message),
     el('span', { class: 'muted' }, tab.path + ':' + p.line))));
  if (tab) {
    editor.clearGutter('problem-gutter');
    for (const p of problems) editor.setGutterMarker(p.line - 1, 'problem-gutter',
      el('span', { class: 'gutter-mark sev-' + p.severity, title: p.message }, p.severity === 'error' ? '●' : '▲'));
  }
}

// --- files and functions --------------------------------------------------------------------------
async function refreshFiles() {
  state.files = await get('files');
  renderFiles();
}
function renderFiles() {
  const files = state.files || [];
  $('file-list').replaceChildren(...(files.length ? files.map(f => el('li', {
    class: state.active && state.active.path === f ? 'active' : '', title: f, onclick: () => open(f),
  }, f)) : [el('li', { class: 'muted' }, 'No modules yet: + makes one.')]));
}
async function newFile() {
  const name = await input('A .py file of the Python folder; a name with / makes a subfolder, which Excel does not load.',
                           'my_functions.py', 'New module');
  if (!name || !name.trim()) return;
  try {
    const r = await post('new', { path: name.trim() });
    await refreshFiles();
    await open(r.path);
  } catch (e) { message(e.message, 'New module'); }
}

async function refreshOutline() {
  const items = await get('outline').catch(() => []);
  const groups = [['xl_func', 'Worksheet functions'], ['xl_macro', 'Macros'], ['export', 'For PY_CALL and ASF scripts']];
  const box = $('outline');
  box.replaceChildren();
  for (const [kind, title] of groups) {
    const these = items.filter(i => i.kind === kind);
    if (!these.length) continue;
    box.append(el('div', { class: 'group' }, title + ' (' + these.length + ')'));
    for (const f of these) {
      const sig = kind === 'xl_func' ? '=' + f.name + '(' + f.args.join(', ') + ')' :
                  kind === 'export' ? 'PY_CALL("' + f.name + '"' + f.args.map(a => ', ' + a).join('') + ')' : f.name;
      const row = el('div', { class: 'fn', title: f.file + ':' + f.line, onclick: () => open(f.file, f.line) },
        kind === 'xl_macro' ? el('button', { title: 'Run in Excel', onclick: e => { e.stopPropagation(); runMacro(f); } }, '▶') : null,
        el('div', { class: 'sig' }, sig), f.doc ? el('div', { class: 'doc' }, f.doc) : null,
        el('div', { class: 'where' }, f.file + ':' + f.line));
      box.append(row);
    }
  }
  if (!box.children.length) box.append(el('p', { class: 'muted', style: 'padding: 0 12px' }, 'No functions marked with @xl_func, @xl_macro or @export yet.'));
}
// a typed argument as Excel would take it from a cell: a number, TRUE or FALSE, blank, or a text
// ("quoted" keeps a text that looks like a number)
function cellValue(text) {
  const t = text.trim();
  if (!t) return null;
  if (/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(t)) return Number(t);
  if (/^(true|false)$/i.test(t)) return t.toLowerCase() === 'true';
  if (t.length > 1 && t.startsWith('"') && t.endsWith('"')) return t.slice(1, -1);
  return text;
}
function remembered(key, fallback) {
  try { return JSON.parse(localStorage.getItem(key)) ?? fallback; } catch (e) { return fallback; }
}
function remember(key, value) {
  try { localStorage.setItem(key, JSON.stringify(value)); } catch (e) { /* not kept: no harm */ }
}
async function runMacro(f) {
  const last = remembered('macro:' + f.name, {});
  const r = await ask({
    title: 'Run ' + f.name + ' in Excel',
    text: (f.doc ? f.doc + '\n' : '') + (f.args.length ? 'Its arguments, as in a cell: 2.5, TRUE, a text, "007" for a text of digits; blank for its default.' : 'It takes no arguments.'),
    fields: f.args.map(a => ({ name: a, label: a, value: last[a] ?? '' })),
    checks: [{ name: '__front', label: 'Bring Excel to the front: for a macro that shows a dialog (xl_app().InputBox, a message box) or works on the selection',
               value: remembered('macro:front', false) }],
    ok: '▶ Run',
  });
  if (!r) return;
  const values = {};
  f.args.forEach(a => { values[a] = r[a]; });
  remember('macro:' + f.name, values);
  remember('macro:front', r.__front);
  const args = f.args.map(a => cellValue(r[a]));
  show('output');
  output('▶ macro ' + f.name + '(' + args.map(v => v === null ? 'blank' : JSON.stringify(v)).join(', ') + ')', 'head');
  foot('Excel runs ' + f.name + '…');
  const waiting = setTimeout(() => foot('Excel still runs ' + f.name + ': if it shows a dialog, answer it in Excel' +
    (r.__front ? '' : ' (tick Bring Excel to the front next time)')), 1500);
  try {
    const res = await post('excel/macro', { name: f.name, args, front: r.__front });
    if (r.__front && res.front === false)
      output('Windows kept Excel behind the studio: its taskbar button flashed. A macro\'s dialog shows there.', 'err');
    output(res.result || '(done: no result)', res.error ? 'err' : 'ok');
    foot(res.error ? f.name + ' failed' : 'Ran ' + f.name);
    message(res.result || 'Done: the macro gave no result.', f.name + (res.error ? ' failed' : ''), true);
  } catch (e) {
    output(e.message, 'err');
    foot(f.name + ' did not run');
    message(e.message, f.name + ' did not run');
  } finally { clearTimeout(waiting); }
}

// --- console -----------------------------------------------------------------------------------
async function consoleRun(restart) {
  const input = $('console-in'), code = restart ? '' : input.value;
  const out = $('console-out');
  if (!restart) {
    if (!code.trim()) return;
    out.append(el('span', { class: 'out-head' }, code.split('\n').map((l, i) => (i ? '... ' : '>>> ') + l).join('\n') + '\n'));
    state.history.push(code);
    state.historyAt = state.history.length;
    input.value = '';
    input.rows = 1;
  }
  try {
    const r = await post('console', { code, restart: !!restart });
    if (r.output) out.append(el('span', { class: /Traceback|Error/.test(r.output) ? 'out-err' : '' }, r.output));
  } catch (e) { out.append(el('span', { class: 'out-err' }, e.message + '\n')); }
  out.scrollTop = out.scrollHeight;
}
$('console-in').addEventListener('keydown', e => {
  const input = e.target;
  if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); consoleRun(); }
  else if (e.key === 'Enter') { input.rows = Math.min(input.rows + 1, 8); }
  else if ((e.key === 'ArrowUp' || e.key === 'ArrowDown') && !input.value.includes('\n')) {
    const h = state.history;
    if (!h.length) return;
    state.historyAt = Math.max(0, Math.min(h.length, state.historyAt + (e.key === 'ArrowUp' ? -1 : 1)));
    input.value = h[state.historyAt] || '';
    e.preventDefault();
  }
});

// --- packages ----------------------------------------------------------------------------------
const SUGGESTED = [
  ['numpy', 'Arrays and maths: the numpy.ndarray type hint'], ['pandas', 'Tables: DataFrame and Series type hints (64-bit)'],
  ['scipy', 'Statistics, optimization, interpolation, signal processing'], ['statsmodels', 'Regression, time series, statistical tests'],
  ['scikit-learn', 'Machine learning: models, clustering, preprocessing'], ['matplotlib', 'Charts, saved as images'],
  ['openpyxl', 'Reading and writing .xlsx files'], ['xlsxwriter', 'Writing formatted .xlsx files'],
  ['requests', 'HTTP: web APIs and downloads'], ['python-dateutil', 'Dates: parsing and relative deltas'],
  ['pywin32', 'Windows and COM: xl_app() in macros'], ['sympy', 'Symbolic maths'],
];
function installed(name) {
  const key = n => n.toLowerCase().replace(/[-_.]+/g, '-');
  return state.packages.find(p => key(p.name) === key(name));
}
async function refreshPackages() {
  state.packages = await get('packages').catch(() => []);
  $('pk-count').textContent = '(' + state.packages.length + ')';
  $('pk-suggested').replaceChildren(...SUGGESTED.map(([name, about]) => {
    const have = installed(name);
    return el('div', { class: 'card' }, el('div', { class: 't' }, name), el('div', { class: 'd' }, about),
      el('div', { class: 'a' }, have ? el('span', { class: 'badge' }, '✓ ' + have.version) : el('span'),
        have ? el('button', { onclick: () => pip('install --upgrade ' + name) }, 'Upgrade')
             : el('button', { class: 'primary', onclick: () => pip('install ' + name) }, 'Install')));
  }));
  renderInstalled();
}
function renderInstalled() {
  const q = $('pk-filter').value.trim().toLowerCase();
  $('pk-installed').replaceChildren(...state.packages.filter(p => !q || p.name.toLowerCase().includes(q) || p.summary.toLowerCase().includes(q))
    .map(p => el('tr', {}, el('td', {}, p.name), el('td', {}, p.version), el('td', { class: 'muted' }, p.summary),
      el('td', {}, el('button', { onclick: () => pip('install --upgrade ' + p.name) }, 'Upgrade'), ' ',
        p.name.toLowerCase() === 'pip' ? null : el('button', { onclick: () => {
          sure('Remove ' + p.name + ' from the add-in\'s Python? Close Excel first if it has imported it.', 'Remove')
            .then(yes => { if (yes) pip('uninstall -y ' + p.name); });
        } }, 'Remove')))));
}
async function pip(command) {
  try { await post('pip', { command }); pollJobs(); } catch (e) { message(e.message, 'pip'); }
}
let jobTimer;
async function pollJobs() {
  clearTimeout(jobTimer);
  const r = await get('jobs').catch(() => null);
  if (!r) return;
  const busy = r.jobs.some(j => j.state === 'queued' || j.state === 'running');
  if (r.revision !== state.jobsRevision) {
    const finished = state.jobsRevision >= 0;
    state.jobsRevision = r.revision;
    const box = $('pk-jobs');
    const opened = new Set([...box.querySelectorAll('details[open]')].map(d => d.dataset.id));
    box.replaceChildren(...r.jobs.slice().reverse().map(j => el('details', {
      class: 'job', 'data-id': String(j.id), ...(opened.has(String(j.id)) || j.state === 'running' ? { open: '' } : {}),
    }, el('summary', {}, el('span', { class: 'state ' + j.state }, j.state), el('code', {}, (j.kind ? '' : 'pip ') + j.command),
        j.code !== null && j.code !== 0 ? el('span', { class: 'sev-error' }, 'exit ' + j.code) : null),
       el('pre', {}, j.output || '…'))));
    box.querySelectorAll('pre').forEach(p => { p.scrollTop = p.scrollHeight; });
    if (finished && !busy) { refreshPackages(); foot('pip finished: restart Excel to use what changed'); }
    const upd = r.jobs.filter(j => j.kind === 'update').pop();
    if (upd) updateProgress(upd);
  }
  if (busy) jobTimer = setTimeout(pollJobs, 800);
}
async function refreshPrevious() {
  const list = await get('previous').catch(() => []);
  const found = list.filter(p => p.packages.length || p.mb);
  $('pk-previous').classList.toggle('hidden', !found.length);
  $('pk-previous-list').replaceChildren(...found.map(p => el('div', { class: 'card' },
    el('div', { class: 't' }, p.folder), el('div', { class: 'd' },
      p.packages.length ? 'Packages this Python lacks: ' + p.packages.map(k => k.name + ' ' + k.version).join(', ')
                        : 'No package this Python lacks.', ' ', p.mb + ' MB on disk.'),
    el('div', { class: 'a' },
      p.packages.length ? el('button', { class: 'primary', onclick: () => post('previous/install', { folder: p.folder })
        .then(pollJobs).catch(e => message(e.message, 'Packages')) }, 'Install them here') : el('span'),
      el('button', { onclick: async () => {
        if (!await sure('Delete ' + p.folder + ' (' + p.mb + ' MB)? An Excel still running the older add-in uses it: close it first.', 'Remove')) return;
        post('previous/remove', { folder: p.folder }).then(refreshPrevious).catch(e => message(e.message, 'Remove'));
      } }, 'Remove')))));
}

// --- the add-in's updates --------------------------------------------------------------------------
function showAddin(s) {
  const pill = $('addin-pill');
  pill.classList.remove('hidden', 'new', 'ok');
  state.update = s;
  const when = s.checked ? 'checked ' + new Date(s.checked * 1000).toLocaleString() : 'not checked yet';
  if (s.restart) {
    pill.textContent = 'ASF Python ' + s.installed + ': restart Excel';
    pill.title = 'The add-in was updated to ' + s.installed + '; this Excel still runs ' + s.running;
  } else if (s.available) {
    pill.textContent = 'Update to ' + s.latest;
    pill.title = 'ASF Python ' + s.latest + ' is available; this add-in is ' + s.installed;
    pill.classList.add('new');
  } else {
    pill.textContent = 'ASF Python ' + s.installed;
    pill.title = 'The add-in: its version, updates and settings';
  }
  $('ad-installed').textContent = s.installed + (s.restart ? ' (installed; this Excel still runs ' + s.running + ': restart Excel to use it)' : '');
  $('ad-latest').textContent = s.latest ? s.latest + (s.available ? ' — newer than this add-in' : ' — this add-in is up to date') + ' (' + when + ')'
                                        : 'not known: ' + when;
  $('ad-install').classList.toggle('hidden', !s.available);
  $('ad-install').textContent = 'Install ' + s.latest + (s.size ? ' (' + Math.round(s.size / 1e6) + ' MB)' : '');
  $('ad-page').classList.toggle('hidden', !s.page);
  if (s.page) $('ad-page').href = s.page;
}
function updateProgress(job) {
  const pill = $('addin-pill'), box = $('ad-progress');
  box.classList.remove('hidden');
  const lines = (job.output || '…').split('\n'), pct = lines.filter(l => /^\d+%$/.test(l)).pop();
  box.textContent = lines.filter(l => !/^\d+%$/.test(l)).join('\n').trim() + (pct && job.state === 'running' ? '\n' + pct : '');
  box.scrollTop = box.scrollHeight;
  if (job.state === 'running' || job.state === 'queued') {
    const pct = (job.output.match(/(\d+)%\n$/) || [])[1];
    pill.textContent = 'Updating' + (pct ? ' ' + pct + '%' : '…');
    $('ad-install').disabled = $('ad-check').disabled = true;
  } else if (!state.updateShown || state.updateShown < job.id) {
    state.updateShown = job.id;
    $('ad-install').disabled = $('ad-check').disabled = false;
    get('update').then(showAddin).catch(() => {});
    foot(job.state === 'done' ? (job.result ? 'ASF Python ' + job.result.version + ' is installed: restart Excel to use it'
                                            : 'ASF Python is up to date') : 'The add-in was not updated: see Add-in');
  }
}
async function checkUpdates() {
  const button = $('ad-check');
  button.disabled = true;
  button.textContent = 'Checking…';
  try {
    showAddin(await get('update', { live: 1 }));
    foot(state.update.available ? 'ASF Python ' + state.update.latest + ' is available' : 'ASF Python ' + state.update.installed + ' is the latest version');
  } catch (e) { message('ASF Python could not check for updates: ' + e.message, 'Updates'); }
  button.disabled = false;
  button.textContent = 'Check for updates';
}
async function installUpdate() {
  const u = state.update;
  if (!u || !u.available) return;
  if (!await sure('Download ASF Python ' + u.latest + ' (' + Math.round((u.size || 0) / 1e6) + ' MB) and install it in the place of ' +
                  u.installed + '? Excel uses it the next time it starts.', 'Install', 'Update')) return;
  try { await post('update'); pollJobs(); } catch (e) { message(e.message, 'Update'); }
}
async function refreshAddin() {
  get('update').then(showAddin).catch(e => { $('ad-installed').textContent = e.message; });
  try {
    const a = await get('addin');
    document.querySelectorAll('input[name=ad-updates]').forEach(r => { r.checked = r.value === a.PythonUpdates; });
    $('ad-tips').checked = a.PythonTips !== 'off';
    $('ad-xll').textContent = a.xll || '(the studio was opened without the add-in)';
    $('ad-python').textContent = a.python + (a.venv ? ' (Excel runs ' + a.venv + ', the setting PythonVenv)' : '');
    $('ad-log').textContent = a.log;
    $('ad-update-log').textContent = a.updateLog;
    $('ad-source').textContent = 'Releases: ' + a.releases;
  } catch (e) { foot(e.message); }
  pollJobs();
}
async function saveSetting(name, value) {
  try { await post('setting', { name, value }); foot(name + ' = ' + value + ': Excel reads it the next time it starts'); }
  catch (e) { message(e.message, name); refreshAddin(); }
}
$('addin-pill').addEventListener('click', () => view('addin'));
$('ad-check').addEventListener('click', checkUpdates);
$('ad-install').addEventListener('click', installUpdate);
document.querySelectorAll('input[name=ad-updates]').forEach(r => r.addEventListener('change', () => saveSetting('PythonUpdates', r.value)));
$('ad-tips').addEventListener('change', e => saveSetting('PythonTips', e.target.checked ? 'on' : 'off'));
document.querySelectorAll('[data-open]').forEach(b => b.addEventListener('click', () =>
  post('open', { what: b.dataset.open }).catch(e => message(e.message, 'Open'))));
$('pk-form').addEventListener('submit', async e => {
  e.preventDefault();
  const name = $('pk-name').value.trim();
  if (!name) return;
  const box = $('pk-lookup');
  box.classList.remove('hidden');
  box.replaceChildren(el('span', { class: 'muted' }, 'Looking ' + name + ' up on PyPI…'));
  try {
    const p = await get('pypi', { name });
    const have = installed(p.name);
    box.replaceChildren(el('div', { class: 't' }, el('b', {}, p.name), ' ', p.version, have ? ' — installed: ' + have.version : ''),
      el('div', { class: 'muted' }, p.summary), el('div', { style: 'margin-top: 8px' },
        el('button', { class: 'primary', onclick: () => { pip('install ' + (have ? '--upgrade ' : '') + p.name); box.classList.add('hidden'); } },
          have ? 'Upgrade to ' + p.version : 'Install ' + p.version)));
  } catch (err) { box.replaceChildren(el('span', { class: 'sev-error' }, err.message)); }
});
$('pk-filter').addEventListener('input', renderInstalled);

// --- views, keys, start ---------------------------------------------------------------------------
function view(name) {
  const page = name === 'packages' || name === 'addin';      // a page of its own, without the editor
  document.querySelectorAll('.act').forEach(b => b.classList.toggle('active', b.dataset.view === name));
  $('packages-view').classList.toggle('hidden', name !== 'packages');
  $('addin-view').classList.toggle('hidden', name !== 'addin');
  $('editor-view').classList.toggle('hidden', page);
  $('side').classList.toggle('hidden', page);
  document.querySelectorAll('.pane').forEach(p => p.classList.toggle('hidden', p.dataset.pane !== name));
  if (name === 'packages') { refreshPackages(); refreshPrevious(); pollJobs(); }
  if (name === 'addin') refreshAddin();
  if (name === 'functions') refreshOutline();
  if (!page) editor.refresh();
}
document.querySelectorAll('.act').forEach(b => b.addEventListener('click', () => view(b.dataset.view)));
document.querySelectorAll('.ptab').forEach(b => b.addEventListener('click', () => show(b.dataset.panel)));
$('clear-panel').addEventListener('click', () => {
  const panel = document.querySelector('.ptab.active').dataset.panel;
  if (panel === 'output') $('output').replaceChildren();
  if (panel === 'console') $('console-out').replaceChildren();
});
$('save').addEventListener('click', save);
$('run').addEventListener('click', run);
$('reload').addEventListener('click', () => reload());
$('new-file').addEventListener('click', newFile);
$('reveal').addEventListener('click', () => post('reveal').then(() => foot('Opened the Python folder in Explorer'))
  .catch(e => { output(e.message, 'err'); show('output'); }));
$('refresh-outline').addEventListener('click', refreshOutline);
$('console-restart').addEventListener('click', () => consoleRun(true));
document.addEventListener('keydown', e => {
  if (document.querySelector('dialog[open]')) return;     // a dialog's keys are its own
  if ((e.ctrlKey || e.metaKey) && e.key === 's') { e.preventDefault(); save(); }
  else if (e.key === 'F5') { e.preventDefault(); run(); }
  else if (e.ctrlKey && e.shiftKey && (e.key === 'R' || e.key === 'r')) { e.preventDefault(); reload(); }
  else if (e.ctrlKey && ['1', '2', '3', '4'].includes(e.key)) { e.preventDefault(); view(['files', 'functions', 'packages', 'addin'][+e.key - 1]); }
});
window.addEventListener('beforeunload', e => {
  if (state.tabs.some(t => t.doc.getValue() !== t.saved)) { e.preventDefault(); e.returnValue = ''; }
});

async function start() {
  try {
    const info = await get('info');
    state.folder = info.folder;
    state.excel = info.excel;
    $('python-pill').textContent = 'Python ' + info.python + ' · ' + info.bits + '-bit';
    $('python-pill').title = info.prefix;
    const pill = $('excel-pill');
    pill.textContent = info.excel ? 'Excel ' + info.excel : 'Excel not reached';
    pill.className = 'pill ' + (info.excel ? 'ok' : 'off');
    $('folder-path').textContent = info.folder || 'No Python folder';
    $('foot-right').textContent = info.folder;
    get('update').then(showAddin).catch(() => {});
    await refreshFiles();
    if (state.files.length) await open(state.files[0]);
    const asked = location.hash.slice(1);                  // Excel's Python updates button opens #addin
    if (['functions', 'packages', 'addin'].includes(asked)) view(asked);
    foot('Ready');
  } catch (e) { foot(e.message); }
  setInterval(() => post('ping').catch(() => foot('The studio stopped: open it again from Excel (ASF tab, Python studio).')), 20000);
}
start();
