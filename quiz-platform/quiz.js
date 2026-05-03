/**
 * quiz.js  —  Complete Quiz System (single file)
 *
 * Admin page : quiz-admin.html  →  needs  <div id="admin-root">
 * Agent page : quiz-agent.html  →  needs  <div id="agent-root">
 *
 * Set window.QUIZ_AGENT_PAGE_URL before loading this file in quiz-admin.html:
 *   <script>window.QUIZ_AGENT_PAGE_URL = 'https://your-tenant.sharepoint.com/SitePages/quiz-agent.html';</script>
 */

// ─────────────────────────────────────────────────────────────
// CONSTANTS
// ─────────────────────────────────────────────────────────────

const _L = {
  templates : 'YOUR-QUIZ-TEMPLATES-LIST-GUID',
  quizzes   : 'YOUR-QUIZZES-LIST-GUID',
  questions : 'YOUR-QUESTIONS-LIST-GUID',
  attempts  : 'YOUR-ATTEMPTS-LIST-GUID',
  links     : 'YOUR-LINKS-LIST-GUID',
  agents    : 'YOUR-RECOGNITION-LIST-GUID'
};

const PASS_SCORE = 80;   // % required to pass

const _site = (() => {
  try { return (_spPageContextInfo && _spPageContextInfo.webAbsoluteUrl) || window.location.origin; }
  catch (_) { return window.location.origin; }
})();

let _digest = '';

// Persists selected template / open question editor across partial refreshes
const _setupState = { templateKey: '', editingTemplateKey: '', editingQuizTitle: '', dir: [] };


// ─────────────────────────────────────────────────────────────
// SHAREPOINT UTILITIES
// ─────────────────────────────────────────────────────────────

async function _getDigest() {
  if (_digest) return _digest;
  const r = await fetch(`${_site}/_api/contextinfo`, {
    method: 'POST',
    headers: { Accept: 'application/json;odata=nometadata' },
    credentials: 'same-origin'
  });
  if (!r.ok) throw new Error(`Digest request failed (${r.status})`);
  _digest = (await r.json()).FormDigestValue;
  setTimeout(() => { _digest = ''; }, 25 * 60 * 1000); // auto-expire before SP 30-min limit
  return _digest;
}

async function _spFetch(listGuid, { filter, select, expand, top = 5000 } = {}) {
  let url = `${_site}/_api/web/lists(guid'${listGuid}')/items?$top=${top}`;
  if (filter) url += `&$filter=${encodeURIComponent(filter)}`;
  if (select) url += `&$select=${encodeURIComponent(select)}`;
  if (expand) url += `&$expand=${encodeURIComponent(expand)}`;
  const r = await fetch(url, {
    headers: { Accept: 'application/json;odata=nometadata' },
    credentials: 'same-origin'
  });
  if (!r.ok) throw new Error(`Fetch failed (${r.status})`);
  return (await r.json()).value || [];
}

async function _spCreate(listGuid, payload) {
  const digest = await _getDigest();
  const r = await fetch(`${_site}/_api/web/lists(guid'${listGuid}')/items`, {
    method: 'POST',
    headers: {
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'X-RequestDigest': digest
    },
    credentials: 'same-origin',
    body: JSON.stringify(payload)
  });
  if (!r.ok) {
    const msg = await r.text().catch(() => '');
    throw new Error(`Create failed (${r.status}): ${msg.substring(0, 300)}`);
  }
  return r.json();
}

async function _spUpdate(listGuid, id, payload) {
  const digest = await _getDigest();
  const r = await fetch(`${_site}/_api/web/lists(guid'${listGuid}')/items(${id})`, {
    method: 'POST',
    headers: {
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'X-RequestDigest': digest,
      'X-HTTP-Method': 'MERGE',
      'IF-MATCH': '*'
    },
    credentials: 'same-origin',
    body: JSON.stringify(payload)
  });
  if (!r.ok) throw new Error(`Update failed (${r.status})`);
}

async function _spDelete(listGuid, id) {
  const digest = await _getDigest();
  const r = await fetch(`${_site}/_api/web/lists(guid'${listGuid}')/items(${id})`, {
    method: 'POST',
    headers: {
      'X-RequestDigest': digest,
      'X-HTTP-Method': 'DELETE',
      'IF-MATCH': '*'
    },
    credentials: 'same-origin'
  });
  if (!r.ok) throw new Error(`Delete failed (${r.status})`);
}


// ─────────────────────────────────────────────────────────────
// PARSERS
// ─────────────────────────────────────────────────────────────

function _yn(v)       { return String(v || '').trim() === 'Yes'; }
function _html(str)   { return String(str || '').replace(/<[^>]+>/g, '').trim(); }

function _parseOpts(raw) {
  if (!raw) return [];
  try { const p = JSON.parse(raw); return Array.isArray(p) ? p : []; }
  catch (_) { return String(raw).split(/\n|,/).map(s => s.trim()).filter(Boolean); }
}

function _parseCorrectAnswers(raw) {
  if (!raw) return [];
  try { const p = JSON.parse(raw); return Array.isArray(p) ? p.map(String) : [String(p)]; }
  catch (_) { return [String(raw)]; }
}

function _fmtCA(raw) {
  try { const p = JSON.parse(raw); return Array.isArray(p) ? p.join(', ') : (raw || ''); }
  catch (_) { return raw || ''; }
}

function _textAnswersMatch(userAnswer, correctAnswer) {
  const u = String(userAnswer   || '').trim();
  const c = String(correctAnswer || '').trim();
  if (u.toLowerCase() === c.toLowerCase()) return true;
  const uNum = parseFloat(u.replace(/[^0-9.]/g, ''));
  const cNum = parseFloat(c.replace(/[^0-9.]/g, ''));
  if (!isNaN(uNum) && !isNaN(cNum) && uNum === cNum) return true;
  return false;
}

function _parseAnswers(raw) {
  if (!raw) return { responses: [], resultStatus: '' };
  try { return JSON.parse(raw); }
  catch (_) { return { responses: [], resultStatus: '' }; }
}

function _parseTmpl(item) {
  return {
    id          : item.ID,
    title       : item.Title || '',
    templateKey : item.TemplateKey || '',
    description : _html(item.Description)
  };
}

function _parseQuiz(item) {
  return {
    id               : item.ID,
    title            : item.Title || '',
    templateKey      : item.TemplateKey || '',
    description      : _html(item.Description),
    isActive         : _yn(item.IsActive),
    isDeleted        : String(item.IsActive || '').trim() === 'Deleted',
    targetDepartment : item.TargetDepartment || '',
    targetLocation   : item.TargetLocation   || '',
    dueDate          : item.DueDate  || null,
    batch            : item.Batch    || '',
    supervisor       : item.Supervisor?.Title || '',
    createdByName    : item.CreatedByName || ''
  };
}

function _parseQuestion(item) {
  return {
    id           : item.ID,
    templateKey  : item.TemplateKey  || '',
    questionType : String(item.QuestionType || 'mcq').toLowerCase() === 'text' ? 'text' : 'mcq',
    questionText : item.QuestionText || '',
    options      : _parseOpts(item.Options),
    correctAnswer : item.CorrectAnswer || '',
    correctAnswers: _parseCorrectAnswers(item.CorrectAnswer || ''),
    reason        : item.Reason       || '',
    sortOrder    : Number(item.SortOrder || 0),
    isActive     : _yn(item.IsActive)
  };
}

function _parseAttempt(item) {
  return {
    id           : item.ID,
    agent        : item.Agent        || '',
    templateKey  : item.TemplateKey  || '',
    quizId       : Number(item.QuizId || 0),
    batchKey     : item.BatchKey     || '',
    score        : Number(item.Score || 0),
    completedDate: item.CompletedDate || null,
    tokenUsed    : item.TokenUsed    || '',
    answers      : _parseAnswers(item.Answers)
  };
}

function _parseLink(item) {
  return {
    id          : item.ID,
    token       : item.Title       || '',
    quizId      : Number(item.QuizId || 0),
    templateKey : item.TemplateKey || '',
    quizTitle   : item.QuizTitle   || '',
    isActive    : _yn(item.IsActive),
    isDeleted   : String(item.IsActive || '').trim() === 'Deleted',
    createdDate : item.CreatedDate || null,
    linkLabel   : item.LinkLabel   || ''
  };
}


// ─────────────────────────────────────────────────────────────
// DATA LOADERS
// ─────────────────────────────────────────────────────────────

async function _loadAll() {
  const [t, q, qs, a] = await Promise.all([
    _spFetch(_L.templates),
    _spFetch(_L.quizzes),
    _spFetch(_L.questions),
    _spFetch(_L.attempts)
  ]);
  return {
    templates : t.map(_parseTmpl),
    quizzes   : q.map(_parseQuiz),
    questions : qs.map(_parseQuestion).sort((a, b) => (a.sortOrder - b.sortOrder) || (a.id - b.id)),
    attempts  : a.map(_parseAttempt)
  };
}

async function _loadLinks() {
  const items = await _spFetch(_L.links);
  return items.map(_parseLink).sort((a, b) => new Date(b.createdDate || 0) - new Date(a.createdDate || 0));
}

async function _loadAgentDir() {
  try {
    const items = await _spFetch(_L.agents, { top: 5000 });
    const seen  = new Set();
    return items.map(item => ({
      name       : String(item.FullName || item.fullName || item.Name || item.EmployeeName || item.Title || '').trim(),
      department : String(item.Department || item.Dept || item.DepartmentName || '').trim(),
      location   : String(item.Location  || item.SiteLocation || item.Office || '').trim(),
      hireDate   : item.HireDate || item.Hiredate || item.StartDate || null
    })).filter(a => {
      if (!a.name) return false;
      const key = a.name.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    }).sort((a, b) => a.name.localeCompare(b.name));
  } catch (_) { return []; }
}


// ─────────────────────────────────────────────────────────────
// UTILITIES
// ─────────────────────────────────────────────────────────────

function _slugify(str) {
  return String(str || '').trim().toLowerCase()
    .replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
}

function _token() {
  return Date.now().toString(36) + Math.random().toString(36).substr(2, 9);
}

function _batch() {
  const n = new Date();
  return `batch-${n.getFullYear()}${String(n.getMonth() + 1).padStart(2, '0')}`;
}

function _qp(name) {
  return new URLSearchParams(window.location.search).get(name);
}

// HTML-escape for safe innerHTML insertion
function _e(str) {
  return String(str || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// Escape single quotes for SP REST $filter strings
function _spStr(str) { return String(str || '').replace(/'/g, "''"); }

function _fmtDate(str) {
  if (!str) return '—';
  return new Date(str).toLocaleString();
}

function _agentPageUrl() {
  if (window.QUIZ_AGENT_PAGE_URL) return window.QUIZ_AGENT_PAGE_URL;
  const attr = document.getElementById('admin-root')?.dataset?.agentUrl;
  if (attr) return attr;
  return window.location.href.split('?')[0].replace('quiz-admin.html', 'quiz-agent.html');
}

function _buildTokenUrl(token) {
  const base = _agentPageUrl();
  return `${base}${base.includes('?') ? '&' : '?'}token=${token}`;
}

function _currentUserId() {
  try { return (_spPageContextInfo && _spPageContextInfo.userId) || null; } catch (_) { return null; }
}

// Returns the SharePoint display name of the logged-in user
function _spDisplayName() {
  try { return (_spPageContextInfo && _spPageContextInfo.userDisplayName) || ''; } catch (_) { return ''; }
}

// Normalise a date value to YYYY-MM-DD string for comparison
function _normDateKey(val) {
  if (!val) return '';
  const d = new Date(val);
  if (isNaN(d.getTime())) return '';
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

// Short readable date label  e.g. "Jan 15, 2026"
function _fmtDateShort(val) {
  if (!val) return '';
  const d = new Date(val);
  if (isNaN(d.getTime())) return '';
  return d.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
}

// Deterministic color badge for a name — same name always gets the same color
function _creatorBadge(name) {
  if (!name) return '';
  let hash = 0;
  for (let i = 0; i < name.length; i++) hash = (hash * 31 + name.charCodeAt(i)) >>> 0;
  const hue = hash % 360;
  const bg  = `hsl(${hue},60%,88%)`;
  const fg  = `hsl(${hue},55%,28%)`;
  return `<span style="display:inline-block;padding:1px 8px;border-radius:10px;font-size:11px;font-weight:600;background:${bg};color:${fg};">${_e(name)}</span>`;
}

// Convert batch key "batch-20260312" → "12 March 2026"
function _fmtBatchLabel(batchKey) {
  const m = batchKey.match(/^batch-(\d{4})(\d{2})(\d{2})$/);
  if (!m) return batchKey;
  return new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]))
    .toLocaleDateString('en-GB', { year: 'numeric', month: 'long', day: 'numeric' });
}

// Convert a hire date to the batch key format used on quizzes  e.g. "batch-20260312"
function _hireDateToBatch(hireDate) {
  if (!hireDate) return '';
  const d = new Date(hireDate);
  if (isNaN(d.getTime())) return '';
  return `batch-${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
}

// Try to find the current SharePoint user in the agent directory
function _autoDetect(dir) {
  const name = _spDisplayName();
  if (!name) return null;
  return dir.find(a => a.name.toLowerCase() === name.toLowerCase()) || null;
}

function _csvDownload(filename, rows) {
  const csv  = rows.map(r => r.map(c => `"${String(c ?? '').replace(/"/g, '""')}"`).join(',')).join('\r\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const a    = document.createElement('a');
  a.href     = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}


// ═════════════════════════════════════════════════════════════
// ADMIN DASHBOARD
// ═════════════════════════════════════════════════════════════

async function _initAdmin() {
  const root = document.getElementById('admin-root');
  if (!root) return;

  root.innerHTML = `
    <div class="header">
      <div>
        <h1>Quiz Admin</h1>
        <p class="muted">Create templates, manage quizzes, generate links, and view analytics.</p>
      </div>
    </div>
    <div class="tab-nav">
      <button class="tab-btn active" data-tab="setup">⚙ Setup</button>
      <button class="tab-btn" data-tab="distribute">🔗 Distribute</button>
      <button class="tab-btn" data-tab="analytics">📊 Analytics</button>
    </div>
    <div id="tab-body"></div>
  `;

  const tabBody = document.getElementById('tab-body');

  async function switchTab(tab) {
    root.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
    tabBody.innerHTML = '<p class="muted loading-msg">Loading…</p>';
    try {
      if (tab === 'setup')      await _renderSetup(tabBody);
      else if (tab === 'distribute') await _renderDistribute(tabBody);
      else if (tab === 'analytics')  await _renderAnalytics(tabBody);
    } catch (e) {
      tabBody.innerHTML = `<p class="error">Failed to load: ${_e(e.message)}</p>`;
      console.error('[Quiz]', e);
    }
  }

  root.querySelector('.tab-nav').addEventListener('click', e => {
    const btn = e.target.closest('.tab-btn');
    if (btn) switchTab(btn.dataset.tab);
  });

  await switchTab('setup');
}


// ─────────────────────────────────────────────────────────────
// SETUP TAB
// ─────────────────────────────────────────────────────────────

async function _renderSetup(container) {
  const [data, dir] = await Promise.all([_loadAll(), _loadAgentDir()]);
  _setupState.dir = dir;
  const deptOpts = [...new Set(dir.map(a => a.department).filter(Boolean))].sort();
  const locOpts  = [...new Set(dir.map(a => a.location).filter(Boolean))].sort();

  container.innerHTML = `

    <!-- ── 1. TEMPLATES ── -->
    <div class="section">
      <h2>1. Templates</h2>
      <p class="muted small">A template holds the question bank. Multiple quiz instances can be created from one template.</p>
      <div class="form-row" style="flex-wrap:wrap;">
        <input id="tpl-name" type="text" placeholder="Template name  (e.g. Agent Onboarding)" style="flex:2; min-width:180px;">
        <input id="tpl-desc" type="text" placeholder="Short description  (optional)" style="flex:3; min-width:200px;">
        <button id="tpl-create-btn">Create Template</button>
      </div>
      <div id="tpl-list" class="item-list" style="margin-top:12px;"></div>
    </div>

    <!-- ── 2. QUIZZES ── -->
    <div class="section">
      <h2>2. Quizzes</h2>
      <p class="muted small">A quiz is one instance of a template assigned to a group. Flip it Active to make it visible to agents.</p>
      <div class="form-row">
        <select id="quiz-tpl-sel" style="flex:1; min-width:200px;">
          <option value="">Select a template to work with…</option>
          ${data.templates.map(t => `<option value="${_e(t.templateKey)}">${_e(t.title)}</option>`).join('')}
        </select>
      </div>
      <div id="quiz-create-form" style="display:none; margin-top:12px;">
        <div class="form-row" style="flex-wrap:wrap;">
          <input id="quiz-name" type="text" placeholder="Quiz name  (e.g. January 2026 Cohort)" style="flex:3; min-width:200px;">
          <select id="quiz-dept" style="flex:1; min-width:130px;">
            <option value="">All Departments</option>
            ${deptOpts.map(d => `<option value="${_e(d)}">${_e(d)}</option>`).join('')}
          </select>
          <select id="quiz-loc" style="flex:1; min-width:130px;">
            <option value="">All Locations</option>
            ${locOpts.map(l => `<option value="${_e(l)}">${_e(l)}</option>`).join('')}
          </select>
          <select id="quiz-batch" style="flex:1; min-width:150px;">
            <option value="">All Agents (no batch)</option>
          </select>
          <button id="quiz-create-btn">Create Quiz</button>
        </div>
      </div>
      <div id="quiz-list" class="item-list" style="margin-top:12px;"></div>
    </div>

    <!-- ── 3. QUESTIONS ── -->
    <div class="section">
      <h2>3. Questions</h2>
      <div id="q-scope" class="info-bar">Click <strong>Manage Questions</strong> on a quiz above to edit its question bank.</div>
      <div id="q-editor" style="display:none; margin-top:14px;">
        <div class="form-row" style="flex-wrap:wrap; gap:8px;">
          <select id="q-type" style="width:170px; flex-shrink:0;">
            <option value="mcq">Multiple Choice</option>
            <option value="text">Text Answer</option>
          </select>
          <input id="q-text" type="text" placeholder="Question text  (max 255 characters)" style="flex:1; min-width:220px;">
        </div>
        <div id="q-opts-wrap" style="margin-top:8px;">
          <textarea id="q-opts" rows="5"
            placeholder="One option per line:&#10;Option A&#10;Option B&#10;Option C&#10;Option D"
            style="width:100%; font-family:inherit; padding:8px; border:1px solid #e1dfdd; border-radius:6px; font-size:13px;"></textarea>
        </div>
        <div class="form-row" style="margin-top:8px; flex-wrap:wrap;">
          <input id="q-answer" type="text" placeholder="Correct answer(s) — for multiple separate with comma: Option A, Option B" style="flex:3; min-width:200px;">
          <input id="q-reason" type="text" placeholder="Explanation shown after answer  (optional, max 255)" style="flex:2; min-width:160px;">
          <button id="q-add-btn">Add Question</button>
        </div>
      </div>
      <div id="q-list" class="item-list" style="margin-top:12px;"></div>
    </div>
  `;

  // Question type toggle
  const qTypeEl   = document.getElementById('q-type');
  const qOptsWrap = document.getElementById('q-opts-wrap');
  qTypeEl.addEventListener('change', () => {
    qOptsWrap.style.display = qTypeEl.value === 'mcq' ? '' : 'none';
  });

  // ── Render existing templates ──
  _renderTplList(data.templates);

  // ── Create template ──
  document.getElementById('tpl-create-btn').onclick = async () => {
    const name = document.getElementById('tpl-name').value.trim();
    const desc = document.getElementById('tpl-desc').value.trim();
    if (!name) { alert('Enter a template name.'); return; }
    const key = 'tmpl-' + _slugify(name);
    try {
      await _spCreate(_L.templates, { Title: name, TemplateKey: key, Description: desc });
      document.getElementById('tpl-name').value = '';
      document.getElementById('tpl-desc').value = '';
      const fresh = (await _spFetch(_L.templates)).map(_parseTmpl);
      _renderTplList(fresh);
      _refreshTplDropdown(fresh);
    } catch (e) { alert('Could not create template: ' + e.message); }
  };

  // ── Template dropdown ──
  const quizTplSel     = document.getElementById('quiz-tpl-sel');
  const quizCreateForm = document.getElementById('quiz-create-form');
  const quizListEl     = document.getElementById('quiz-list');

  function _onTemplateSelect(key) {
    _setupState.templateKey = key;
    quizTplSel.value = key;
    const sel1 = document.getElementById('tpl-existing-sel');
    if (sel1) sel1.value = key;
    quizCreateForm.style.display = key ? '' : 'none';
    _renderQuizList(quizListEl, data.quizzes.filter(q => q.templateKey === key), data);
  }

  quizTplSel.onchange = () => _onTemplateSelect(quizTplSel.value.trim());

  // Section 1 dropdown — use delegation so it survives re-renders (e.g. after creating a template)
  document.getElementById('tpl-list').addEventListener('change', (ev) => {
    if (ev.target.id === 'tpl-existing-sel') _onTemplateSelect(ev.target.value.trim());
  });

  // ── Restore state (user came back from a save) ──
  if (_setupState.templateKey) {
    quizTplSel.value = _setupState.templateKey;
    const sel1 = document.getElementById('tpl-existing-sel');
    if (sel1) sel1.value = _setupState.templateKey;
    quizCreateForm.style.display = 'block';
    _renderQuizList(quizListEl, data.quizzes.filter(q => q.templateKey === _setupState.templateKey), data);
    if (_setupState.editingTemplateKey) {
      const qs = data.questions.filter(q => q.templateKey === _setupState.editingTemplateKey);
      _openQEditor(_setupState.editingTemplateKey, _setupState.editingQuizTitle, qs);
    }
  }

  // ── Batch dropdown cascade (filters by selected dept + loc) ──
  function _rebuildBatchSel() {
    const dept = document.getElementById('quiz-dept').value.toLowerCase();
    const loc  = document.getElementById('quiz-loc').value.toLowerCase();
    const batches = [...new Set(
      dir
        .filter(a => (!dept || a.department.toLowerCase() === dept) && (!loc || a.location.toLowerCase() === loc))
        .map(a => _hireDateToBatch(a.hireDate))
        .filter(Boolean)
    )].sort().reverse();
    const sel  = document.getElementById('quiz-batch');
    const prev = sel.value;
    sel.innerHTML = '<option value="">All Agents (no batch)</option>' +
      batches.map(b => `<option value="${_e(b)}">${_fmtBatchLabel(b)}</option>`).join('');
    if (batches.includes(prev)) sel.value = prev;
  }
  document.getElementById('quiz-dept').onchange = _rebuildBatchSel;
  document.getElementById('quiz-loc').onchange  = _rebuildBatchSel;
  _rebuildBatchSel();

  // ── Create quiz ──
  document.getElementById('quiz-create-btn').onclick = async () => {
    const key  = quizTplSel.value.trim();
    const name = document.getElementById('quiz-name').value.trim();
    const dept  = document.getElementById('quiz-dept').value.trim();
    const loc   = document.getElementById('quiz-loc').value.trim();
    const batch = document.getElementById('quiz-batch').value.trim();
    if (!key || !name) { alert('Select a template and enter a quiz name.'); return; }
    const createdByName = await _fetchCurrentUserName();
    const payload = { Title: name, TemplateKey: key, IsActive: 'No', TargetDepartment: dept, TargetLocation: loc, Batch: batch, CreatedByName: createdByName };
    const uid = _currentUserId();
    if (uid) payload.SupervisorId = uid;
    try {
      await _spCreate(_L.quizzes, payload);
      document.getElementById('quiz-name').value  = '';
      document.getElementById('quiz-dept').value  = '';
      document.getElementById('quiz-loc').value   = '';
      _rebuildBatchSel();
      const fresh = (await _spFetch(_L.quizzes)).map(_parseQuiz);
      _renderQuizList(quizListEl, fresh.filter(q => q.templateKey === key), { ...data, quizzes: fresh });
    } catch (e) { alert('Could not create quiz: ' + e.message); }
  };
}

function _refreshTplDropdown(templates) {
  const sel  = document.getElementById('quiz-tpl-sel');
  const sel1 = document.getElementById('tpl-existing-sel');
  if (!sel) return;
  const current = sel.value;
  sel.innerHTML = ['<option value="">Select a template to work with…</option>',
    ...templates.map(t => `<option value="${_e(t.templateKey)}">${_e(t.title)}</option>`)
  ].join('');
  if (sel1) {
    sel1.innerHTML = ['<option value="">— Select a template —</option>',
      ...templates.map(t => `<option value="${_e(t.templateKey)}">${_e(t.title)}${t.description ? '  ·  ' + _e(t.description) : ''}</option>`)
    ].join('');
    if (current) sel1.value = current;
  }
  if (current) sel.value = current;
}

function _renderTplList(templates) {
  const el = document.getElementById('tpl-list');
  if (!el) return;
  if (!templates.length) { el.innerHTML = '<p class="muted">No templates yet. Create one above.</p>'; return; }
  el.innerHTML = `
    <div class="form-row" style="margin-top:4px;">
      <select id="tpl-existing-sel" style="flex:1;">
        <option value="">— Select a template —</option>
        ${templates.map(t => `<option value="${_e(t.templateKey)}">${_e(t.title)}${t.description ? '  ·  ' + _e(t.description) : ''}</option>`).join('')}
      </select>
    </div>`;
}

function _renderQuizList(el, quizzes, data) {
  if (!el) return;
  const visible = quizzes.filter(q => !q.isDeleted);
  if (!visible.length) { el.innerHTML = '<p class="muted">No quizzes yet for this template. Create one above.</p>'; return; }
  el.innerHTML = visible.map(q => `
    <div class="list-item" id="quiz-row-${q.id}">
      <div style="flex:1;">
        <strong>${_e(q.title)}</strong>
        <span class="badge ${q.isActive ? 'badge-active' : 'badge-draft'}">${q.isActive ? 'Active' : 'Draft'}</span>
        ${q.targetDepartment ? `<span class="muted small"> · Dept: ${_e(q.targetDepartment)}</span>` : ''}
        ${q.targetLocation   ? `<span class="muted small"> · Loc: ${_e(q.targetLocation)}</span>`   : ''}
        ${q.batch          ? `<span class="muted small"> · ${_fmtBatchLabel(q.batch)}</span>`             : ''}
        ${q.createdByName  ? `<span class="muted small"> · Created by: </span>${_creatorBadge(q.createdByName)}`  : ''}
      </div>
      <div class="btn-group">
        <button class="btn-sm" data-action="edit" data-id="${q.id}" data-title="${_e(q.title)}" data-dept="${_e(q.targetDepartment)}" data-loc="${_e(q.targetLocation)}" data-batch="${_e(q.batch)}">Edit</button>
        <button class="btn-sm" data-action="manage"
          data-tpl="${_e(q.templateKey)}" data-qtitle="${_e(q.title)}">Manage Questions</button>
        <button class="btn-sm ${q.isActive ? 'btn-warn' : 'btn-ok'}"
          data-action="toggle" data-id="${q.id}" data-active="${q.isActive}">
          ${q.isActive ? 'Deactivate' : 'Activate'}
        </button>
        <button class="btn-sm" data-action="duplicate" data-id="${q.id}" data-title="${_e(q.title)}" data-tpl="${_e(q.templateKey)}" data-dept="${_e(q.targetDepartment)}" data-loc="${_e(q.targetLocation)}" data-batch="${_e(q.batch)}">Duplicate</button>
        <button class="btn-sm btn-danger" data-action="delete" data-id="${q.id}" data-title="${_e(q.title)}">Delete</button>
      </div>
    </div>
  `).join('');

  el.onclick = async (e) => {
    const btn = e.target.closest('[data-action]');
    if (!btn) return;

    if (btn.dataset.action === 'edit') {
      const id    = Number(btn.dataset.id);
      const title = btn.dataset.title;
      const dept  = btn.dataset.dept;
      const loc   = btn.dataset.loc;
      const batch = btn.dataset.batch;
      const dir   = _setupState.dir;
      const deptOpts = [...new Set(dir.map(a => a.department).filter(Boolean))].sort();
      const locOpts  = [...new Set(dir.map(a => a.location).filter(Boolean))].sort();
      function _editBatches(d, l) {
        return [...new Set(
          dir.filter(a => (!d || a.department.toLowerCase() === d.toLowerCase()) && (!l || a.location.toLowerCase() === l.toLowerCase()))
            .map(a => _hireDateToBatch(a.hireDate)).filter(Boolean)
        )].sort().reverse();
      }
      const rowEl = document.getElementById(`quiz-row-${id}`);
      if (!rowEl) return;
      btn.disabled = true; btn.textContent = 'Loading…';
      let locked = false;
      try {
        const existing = await _spFetch(_L.attempts, { filter: `QuizId eq ${id}`, top: 1, select: 'Id' });
        locked = existing.length > 0;
      } catch (_) { /* if check fails, allow editing */ }
      btn.disabled = false; btn.textContent = 'Edit';
      const dis = locked ? 'disabled style="opacity:.55;cursor:not-allowed;"' : '';
      rowEl.innerHTML = `
        <div style="width:100%;">
          ${locked ? `<div style="padding:6px 10px;margin-bottom:8px;background:#fff4ce;border:1px solid #f0d060;border-radius:6px;font-size:12px;color:#7a5c00;">
            Dept, Location and Batch are locked because this quiz has existing attempts. Only the name can be changed.
          </div>` : ''}
          <div class="form-row" style="flex-wrap:wrap; gap:8px; margin-bottom:8px;">
            <input id="edit-name-${id}" type="text" value="${_e(title)}" placeholder="Quiz name" style="flex:3; min-width:180px;">
            <select id="edit-dept-${id}" style="flex:1; min-width:130px;" ${dis}>
              <option value="">All Departments</option>
              ${deptOpts.map(d => `<option value="${_e(d)}" ${d === dept ? 'selected' : ''}>${_e(d)}</option>`).join('')}
            </select>
            <select id="edit-loc-${id}" style="flex:1; min-width:130px;" ${dis}>
              <option value="">All Locations</option>
              ${locOpts.map(l => `<option value="${_e(l)}" ${l === loc ? 'selected' : ''}>${_e(l)}</option>`).join('')}
            </select>
            <select id="edit-batch-${id}" style="flex:1; min-width:150px;" ${dis}>
              <option value="">All Agents (no batch)</option>
              ${_editBatches(dept, loc).map(b => `<option value="${_e(b)}" ${b === batch ? 'selected' : ''}>${_fmtBatchLabel(b)}</option>`).join('')}
            </select>
          </div>
          <div style="display:flex; gap:8px; justify-content:flex-end;">
            <button class="btn-sm btn-ok" data-action="edit-save" data-id="${id}" data-locked="${locked}">Save</button>
            <button class="btn-sm" data-action="edit-cancel">Cancel</button>
          </div>
        </div>`;
      if (!locked) {
        const deptSel  = document.getElementById(`edit-dept-${id}`);
        const locSel   = document.getElementById(`edit-loc-${id}`);
        const batchSel = document.getElementById(`edit-batch-${id}`);
        function _rebuildEditBatch() {
          const d = deptSel.value; const l = locSel.value;
          const batches = _editBatches(d, l);
          const prev = batchSel.value;
          batchSel.innerHTML = '<option value="">All Agents (no batch)</option>' +
            batches.map(b => `<option value="${_e(b)}">${_fmtBatchLabel(b)}</option>`).join('');
          if (batches.includes(prev)) batchSel.value = prev;
        }
        deptSel.onchange = _rebuildEditBatch;
        locSel.onchange  = _rebuildEditBatch;
      }
    }

    if (btn.dataset.action === 'edit-save') {
      const id     = Number(btn.dataset.id);
      const locked = btn.dataset.locked === 'true';
      const name   = (document.getElementById(`edit-name-${id}`)?.value || '').trim();
      if (!name) { alert('Quiz name is required.'); return; }
      const payload = { Title: name };
      if (!locked) {
        payload.TargetDepartment = document.getElementById(`edit-dept-${id}`)?.value || '';
        payload.TargetLocation   = document.getElementById(`edit-loc-${id}`)?.value  || '';
        payload.Batch            = document.getElementById(`edit-batch-${id}`)?.value || '';
      }
      try {
        btn.disabled = true; btn.textContent = 'Saving…';
        await _spUpdate(_L.quizzes, id, payload);
        const fresh = (await _spFetch(_L.quizzes)).map(_parseQuiz);
        _renderQuizList(el, fresh.filter(q => q.templateKey === _setupState.templateKey), { ...data, quizzes: fresh });
      } catch (e2) { alert('Could not save: ' + e2.message); btn.disabled = false; btn.textContent = 'Save'; }
    }

    if (btn.dataset.action === 'edit-cancel') {
      _renderQuizList(el, data.quizzes.filter(q => q.templateKey === _setupState.templateKey), data);
    }

    if (btn.dataset.action === 'toggle') {
      const id       = Number(btn.dataset.id);
      const isActive = btn.dataset.active === 'true';
      try {
        await _spUpdate(_L.quizzes, id, { IsActive: isActive ? 'No' : 'Yes' });
        const fresh = (await _spFetch(_L.quizzes)).map(_parseQuiz);
        _renderQuizList(el, fresh.filter(q => q.templateKey === _setupState.templateKey), { ...data, quizzes: fresh });
      } catch (e2) { alert('Could not update quiz: ' + e2.message); }
    }

    if (btn.dataset.action === 'duplicate') {
      const title   = btn.dataset.title;
      const tpl     = btn.dataset.tpl;
      const dept    = btn.dataset.dept;
      const loc     = btn.dataset.loc;
      const batch   = btn.dataset.batch;
      const srcTmpl = data.templates.find(t => t.templateKey === tpl);
      const defaultTplName = `Copy of ${srcTmpl?.title || title}`;
      const newTplName = prompt(`Name for the new template:`, defaultTplName);
      if (!newTplName) return;
      try {
        btn.disabled = true;
        btn.textContent = 'Duplicating…';
        const newTplKey     = `tpl-${Date.now()}`;
        const createdByName = await _fetchCurrentUserName();
        await _spCreate(_L.templates, { Title: newTplName.trim(), TemplateKey: newTplKey, Description: '' });
        await _spCreate(_L.quizzes, {
          Title: `Copy of ${title}`, TemplateKey: newTplKey, IsActive: 'No',
          TargetDepartment: dept, TargetLocation: loc, Batch: batch,
          CreatedByName: createdByName
        });
        const srcQuestions = (await _spFetch(_L.questions, { filter: `TemplateKey eq '${_spStr(tpl)}'` })).map(_parseQuestion)
          .filter(q => q.isActive)
          .sort((a, b) => (a.sortOrder - b.sortOrder) || (a.id - b.id));
        for (let i = 0; i < srcQuestions.length; i++) {
          const q = srcQuestions[i];
          await _spCreate(_L.questions, {
            Title        : `${newTplKey}-q${i + 1}`,
            TemplateKey  : newTplKey,
            QuestionType : q.questionType,
            QuestionText : q.questionText,
            Options      : JSON.stringify(q.options),
            CorrectAnswer: q.correctAnswer,
            Reason       : q.reason || '',
            SortOrder    : i + 1,
            IsActive     : 'Yes'
          });
        }
        const [freshQuizzes, freshTpls] = await Promise.all([
          _spFetch(_L.quizzes).then(r => r.map(_parseQuiz)),
          _spFetch(_L.templates).then(r => r.map(_parseTmpl))
        ]);
        _setupState.templateKey = newTplKey;
        _refreshTplDropdown(freshTpls);
        const quizTplSel = document.getElementById('quiz-tpl-sel');
        if (quizTplSel) { quizTplSel.value = newTplKey; document.getElementById('quiz-create-form').style.display = ''; }
        _renderQuizList(el, freshQuizzes.filter(q => q.templateKey === newTplKey), { ...data, quizzes: freshQuizzes, templates: freshTpls });
      } catch (e2) { alert('Could not duplicate quiz: ' + e2.message); btn.disabled = false; btn.textContent = 'Duplicate'; }
    }

    if (btn.dataset.action === 'delete') {
      const id    = Number(btn.dataset.id);
      const title = btn.dataset.title;
      if (!confirm(`Delete "${title}"?\n\nThe quiz will be removed from the admin view but historical attempt data is preserved in analytics.`)) return;
      try {
        await _spUpdate(_L.quizzes, id, { IsActive: 'Deleted' });
        const fresh = (await _spFetch(_L.quizzes)).map(_parseQuiz);
        _renderQuizList(el, fresh.filter(q => q.templateKey === _setupState.templateKey), { ...data, quizzes: fresh });
      } catch (e2) { alert('Could not delete quiz: ' + e2.message); }
    }

    if (btn.dataset.action === 'manage') {
      const tplKey   = btn.dataset.tpl;
      const qTitle   = btn.dataset.qtitle;
      _setupState.editingTemplateKey = tplKey;
      _setupState.editingQuizTitle   = qTitle;
      _openQEditor(tplKey, qTitle, data.questions.filter(q => q.templateKey === tplKey));
    }
  };
}

function _openQEditor(templateKey, quizTitle, questions) {
  const scopeEl  = document.getElementById('q-scope');
  const editorEl = document.getElementById('q-editor');
  const qListEl  = document.getElementById('q-list');
  if (!scopeEl || !editorEl || !qListEl) return;

  scopeEl.innerHTML = `Editing questions for: <strong>${_e(quizTitle)}</strong>`;
  editorEl.style.display = '';

  _renderQList(qListEl, questions, templateKey, quizTitle);

  // Replace onclick each time to avoid stacking handlers
  document.getElementById('q-add-btn').onclick = async () => {
    const type    = document.getElementById('q-type').value;
    const text    = document.getElementById('q-text').value.trim();
    const answer  = document.getElementById('q-answer').value.trim();
    const reason  = document.getElementById('q-reason').value.trim();
    const rawOpts = document.getElementById('q-opts').value;
    const opts    = type === 'mcq' ? rawOpts.split('\n').map(s => s.trim()).filter(Boolean) : [];

    if (!text)                     { alert('Enter the question text.'); return; }
    if (!answer)                   { alert('Enter the correct answer.'); return; }
    if (type === 'mcq' && opts.length < 2) { alert('Enter at least 2 options for multiple choice.'); return; }
    if (text.length   > 255)       { alert('Question text is over 255 characters.'); return; }
    if (reason.length > 255)       { alert('Explanation is over 255 characters.'); return; }

    const correctAnswers = type === 'mcq'
      ? answer.split(',').map(s => s.trim()).filter(Boolean)
      : [answer];
    if (type === 'mcq' && !correctAnswers.every(ca => opts.includes(ca))) {
      alert('All correct answers must exactly match one of the options (case-sensitive).\n\nFor multiple correct answers separate with comma: Option A, Option B');
      return;
    }
    const storedAnswer = (type === 'mcq' && correctAnswers.length > 1)
      ? JSON.stringify(correctAnswers)
      : answer;

    const maxSort  = questions.reduce((m, q) => Math.max(m, q.sortOrder), 0);
    const nextSort = maxSort + 1;

    try {
      await _spCreate(_L.questions, {
        Title        : `${templateKey}-q${nextSort}`,
        TemplateKey  : templateKey,
        QuestionType : type === 'text' ? 'text' : 'mcq',
        QuestionText : text,
        Options      : JSON.stringify(type === 'mcq' ? opts : []),
        CorrectAnswer: storedAnswer,
        Reason       : reason,
        SortOrder    : nextSort,
        IsActive     : 'Yes'
      });
      document.getElementById('q-text').value   = '';
      document.getElementById('q-opts').value   = '';
      document.getElementById('q-answer').value = '';
      document.getElementById('q-reason').value = '';
      const fresh = (await _spFetch(_L.questions, { filter: `TemplateKey eq '${_spStr(templateKey)}'` }))
        .map(_parseQuestion).sort((a, b) => (a.sortOrder - b.sortOrder) || (a.id - b.id));
      questions = fresh; // update local reference
      _renderQList(qListEl, fresh, templateKey, quizTitle);
    } catch (e) { alert('Could not add question: ' + e.message); }
  };
}

function _renderQList(el, questions, templateKey, quizTitle) {
  if (!el) return;
  if (!questions.length) { el.innerHTML = '<p class="muted">No questions yet. Add one using the form above.</p>'; return; }

  el.innerHTML = questions.map(q => `
    <div class="list-item ${q.isActive ? '' : 'item-dim'}">
      <div style="flex:1;">
        <span class="muted small">#${q.sortOrder}</span>
        <span class="badge" style="margin-left:6px;">${q.questionType === 'text' ? 'Text' : 'MCQ'}</span>
        ${!q.isActive ? '<span class="badge badge-draft" style="margin-left:4px;">Disabled</span>' : ''}
        <strong style="margin-left:6px;">${_e(q.questionText)}</strong>
        ${q.questionType === 'mcq' ? `
          <div style="margin-top:5px; padding-left:12px;">
            ${q.options.map(o => `<span class="opt-chip ${q.correctAnswers.includes(o) ? 'opt-correct' : ''}">${_e(o)}</span>`).join('')}
          </div>` : ''}
        <div class="muted small" style="margin-top:3px; padding-left:12px;">
          Answer: <strong>${_e(_fmtCA(q.correctAnswer))}</strong>
          ${q.reason ? ` &nbsp;·&nbsp; <em>${_e(q.reason)}</em>` : ''}
        </div>
      </div>
      <div class="btn-group">
        <button class="btn-sm" data-action="toggle-q" data-id="${q.id}" data-active="${q.isActive}">
          ${q.isActive ? 'Disable' : 'Enable'}
        </button>
        <button class="btn-sm btn-danger" data-action="remove-q" data-id="${q.id}">Remove</button>
      </div>
    </div>
  `).join('');

  el.onclick = async (e) => {
    const btn = e.target.closest('[data-action]');
    if (!btn) return;
    const id = Number(btn.dataset.id);

    if (btn.dataset.action === 'toggle-q') {
      const isActive = btn.dataset.active === 'true';
      try {
        await _spUpdate(_L.questions, id, { IsActive: isActive ? 'No' : 'Yes' });
        const fresh = (await _spFetch(_L.questions, { filter: `TemplateKey eq '${_spStr(templateKey)}'` }))
          .map(_parseQuestion).sort((a, b) => (a.sortOrder - b.sortOrder) || (a.id - b.id));
        _renderQList(el, fresh, templateKey, quizTitle);
      } catch (e2) { alert('Could not update question: ' + e2.message); }
    }

    if (btn.dataset.action === 'remove-q') {
      if (!confirm('Remove this question permanently?')) return;
      try {
        await _spDelete(_L.questions, id);
        const fresh = (await _spFetch(_L.questions, { filter: `TemplateKey eq '${_spStr(templateKey)}'` }))
          .map(_parseQuestion).sort((a, b) => (a.sortOrder - b.sortOrder) || (a.id - b.id));
        _renderQList(el, fresh, templateKey, quizTitle);
      } catch (e2) { alert('Could not remove question: ' + e2.message); }
    }
  };
}


// ─────────────────────────────────────────────────────────────
// DISTRIBUTE TAB
// ─────────────────────────────────────────────────────────────

async function _renderDistribute(container) {
  const [data, links] = await Promise.all([_loadAll(), _loadLinks()]);

  // Saved agent page URL (persisted in localStorage so admin only sets it once)
  const savedAgentUrl = localStorage.getItem('quizAgentPageUrl') || _agentPageUrl();

  container.innerHTML = `
    <div class="section">
      <h2>Agent Quiz Page URL</h2>
      <p class="muted small" style="margin-bottom:10px;">
        This is the SharePoint page where agents take quizzes. Set it once — it will be remembered.
      </p>
      <div class="form-row">
        <input id="dist-agent-url" type="text"
          value="${_e(savedAgentUrl)}"
          placeholder="https://…/SitePages/TESTPAGE13.aspx"
          style="flex:1;">
        <button id="dist-save-url-btn" class="btn-sm btn-ok">Save URL</button>
      </div>
      <p class="muted small" id="dist-url-status" style="margin-top:6px;"></p>
    </div>

    <div class="section">
      <h2>Generate a Quiz Link</h2>
      <p class="muted">
        One link is shared with all agents — each person's result is saved individually.
        Generate a second link at any time without affecting the first; both stay active simultaneously.
      </p>
      <div class="form-row" style="flex-wrap:wrap; gap:8px;">
        <select id="dist-tpl" style="flex:1; min-width:180px;">
          <option value="">All Templates</option>
          ${data.templates.map(t => `<option value="${_e(t.templateKey)}">${_e(t.title)}</option>`).join('')}
        </select>
        <select id="dist-sel" style="flex:2; min-width:200px;">
          <option value="">Select a quiz…</option>
          ${data.quizzes.filter(q => q.isActive && !q.isDeleted).map(q =>
            `<option value="${q.id}" data-tpl="${_e(q.templateKey)}" data-title="${_e(q.title)}">${_e(q.title)}</option>`
          ).join('')}
        </select>
        <input id="dist-label" type="text" placeholder="Label  (e.g. Main Link, Backup)" style="flex:1; min-width:150px;">
        <button id="dist-gen-btn">Generate Link</button>
      </div>
      <div id="dist-output" style="margin-top:16px;"></div>
    </div>

    <div class="section">
      <h2>All Generated Links</h2>
      <div id="dist-all"></div>
    </div>
  `;

  _renderAllLinks(links);

  document.getElementById('dist-tpl').onchange = () => {
    const tplKey = document.getElementById('dist-tpl').value;
    const sel    = document.getElementById('dist-sel');
    sel.innerHTML = '<option value="">Select a quiz…</option>' +
      data.quizzes
        .filter(q => q.isActive && !q.isDeleted && (!tplKey || q.templateKey === tplKey))
        .map(q => `<option value="${q.id}" data-tpl="${_e(q.templateKey)}" data-title="${_e(q.title)}">${_e(q.title)}</option>`)
        .join('');
  };

  document.getElementById('dist-save-url-btn').onclick = () => {
    const url = document.getElementById('dist-agent-url').value.trim();
    if (!url) { alert('Enter the agent page URL.'); return; }
    localStorage.setItem('quizAgentPageUrl', url);
    document.getElementById('dist-url-status').textContent = '✓ Saved — new links will point to this URL.';
  };

  document.getElementById('dist-gen-btn').onclick = async () => {
    const sel      = document.getElementById('dist-sel');
    const quizId   = Number(sel.value);
    const label    = document.getElementById('dist-label').value.trim() || 'Link';
    if (!quizId) { alert('Select a quiz first.'); return; }

    const agentBase = (document.getElementById('dist-agent-url').value.trim()) || _agentPageUrl();
    const tok = _token();
    const url = `${agentBase}${agentBase.includes('?') ? '&' : '?'}token=${tok}`;

    const opt         = sel.options[sel.selectedIndex];
    const templateKey = opt?.dataset?.tpl   || '';
    const quizTitle   = opt?.dataset?.title || '';

    try {
      await _spCreate(_L.links, {
        Title      : tok,
        QuizId     : quizId,
        TemplateKey: templateKey,
        QuizTitle  : quizTitle,
        IsActive   : 'Yes',
        CreatedDate: new Date().toISOString(),
        LinkLabel  : label
      });
    } catch (e) { alert('Could not save link: ' + e.message); return; }

    // Show the new link immediately with a copy button
    document.getElementById('dist-output').innerHTML = `
      <div class="link-box">
        <div style="font-weight:600; margin-bottom:8px;">✓ Link generated — ${_e(label)} &nbsp;·&nbsp; ${_e(quizTitle)}</div>
        <div class="form-row">
          <input id="new-link-url" type="text" value="${_e(url)}" readonly
            style="flex:1; font-size:12px; background:#f3f2f1; cursor:text;">
          <button id="copy-new-link-btn">Copy Link</button>
        </div>
        <p class="muted small" style="margin-top:6px;">
          Share this link with your agents. Each agent takes the quiz once; returning to the same link shows their result.
        </p>
      </div>
    `;
    document.getElementById('copy-new-link-btn').onclick = function () {
      navigator.clipboard.writeText(url)
        .then(() => { this.textContent = 'Copied!'; setTimeout(() => this.textContent = 'Copy Link', 2500); })
        .catch(() => { document.getElementById('new-link-url').select(); alert('Press Ctrl+C to copy.'); });
    };

    _renderAllLinks(await _loadLinks());
  };
}

function _renderAllLinks(links) {
  const el = document.getElementById('dist-all');
  if (!el) return;
  const visible = links.filter(lnk => !lnk.isDeleted);
  if (!visible.length) { el.innerHTML = '<p class="muted">No links generated yet.</p>'; return; }

  const agentBase = localStorage.getItem('quizAgentPageUrl') || _agentPageUrl();

  el.innerHTML = visible.map(lnk => {
    const base = agentBase;
    const url  = `${base}${base.includes('?') ? '&' : '?'}token=${lnk.token}`;
    return `
      <div class="list-item ${lnk.isActive ? '' : 'item-dim'}">
        <div style="flex:1;">
          <strong>${_e(lnk.linkLabel || lnk.token)}</strong>
          <span class="badge ${lnk.isActive ? 'badge-active' : 'badge-draft'}" style="margin-left:6px;">
            ${lnk.isActive ? 'Active' : 'Inactive'}
          </span>
          <span class="muted small"> · ${_e(lnk.quizTitle)}</span>
          <div class="muted small" style="margin-top:3px;">Created: ${_fmtDate(lnk.createdDate)}</div>
          <div class="form-row" style="margin-top:6px;">
            <input class="lnk-url-field" type="text" value="${_e(url)}" readonly
              style="flex:1; font-size:11px; background:#f3f2f1; border:1px solid #e1dfdd; border-radius:4px; padding:4px 6px;">
            <button class="btn-sm copy-lnk-btn" data-url="${_e(url)}">Copy</button>
          </div>
        </div>
        <div class="btn-group" style="margin-left:14px; align-self:flex-start;">
          <button class="btn-sm ${lnk.isActive ? 'btn-warn' : 'btn-ok'}"
            data-action="toggle-lnk" data-id="${lnk.id}" data-active="${lnk.isActive}">
            ${lnk.isActive ? 'Deactivate' : 'Activate'}
          </button>
          <button class="btn-sm btn-danger" data-action="delete-lnk" data-id="${lnk.id}" data-label="${_e(lnk.linkLabel || lnk.token)}">Delete</button>
        </div>
      </div>
    `;
  }).join('');

  el.onclick = async (e) => {
    const copyBtn = e.target.closest('.copy-lnk-btn');
    if (copyBtn) {
      const url = copyBtn.dataset.url;
      navigator.clipboard.writeText(url)
        .then(() => { copyBtn.textContent = 'Copied!'; setTimeout(() => copyBtn.textContent = 'Copy', 2500); })
        .catch(() => { copyBtn.closest('.list-item').querySelector('.lnk-url-field').select(); alert('Press Ctrl+C to copy.'); });
      return;
    }
    const toggleBtn = e.target.closest('[data-action="toggle-lnk"]');
    if (toggleBtn) {
      const id       = Number(toggleBtn.dataset.id);
      const isActive = toggleBtn.dataset.active === 'true';
      try {
        await _spUpdate(_L.links, id, { IsActive: isActive ? 'No' : 'Yes' });
        _renderAllLinks(await _loadLinks());
      } catch (e2) { alert('Could not update link: ' + e2.message); }
    }
    const deleteBtn = e.target.closest('[data-action="delete-lnk"]');
    if (deleteBtn) {
      const id    = Number(deleteBtn.dataset.id);
      const label = deleteBtn.dataset.label;
      if (!confirm(`Delete link "${label}"?\n\nThis link will no longer work for agents.`)) return;
      try {
        await _spUpdate(_L.links, id, { IsActive: 'Deleted' });
        _renderAllLinks(await _loadLinks());
      } catch (e2) { alert('Could not delete link: ' + e2.message); }
    }
  };
}


// ─────────────────────────────────────────────────────────────
// ANALYTICS TAB
// ─────────────────────────────────────────────────────────────

function _showAnswerModal(attempt, quiz, agentMeta) {
  const existing = document.getElementById('quiz-modal-overlay');
  if (existing) existing.remove();

  const m         = agentMeta.get(attempt.agent.toLowerCase());
  const pass      = attempt.score >= PASS_SCORE;
  const responses = attempt.answers?.responses || [];

  const overlay = document.createElement('div');
  overlay.id = 'quiz-modal-overlay';
  overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:9999;display:flex;align-items:flex-start;justify-content:center;padding:40px 16px;overflow-y:auto;';

  overlay.innerHTML = `
    <div style="background:#fff;border-radius:10px;max-width:680px;width:100%;box-shadow:0 8px 32px rgba(0,0,0,.2);padding:28px;position:relative;">
      <button id="modal-close" style="position:absolute;top:14px;right:16px;background:none;border:none;font-size:22px;color:#605e5c;cursor:pointer;padding:0;line-height:1;">✕</button>
      <h2 style="font-size:18px;margin-bottom:4px;">${_e(quiz?.title || attempt.templateKey)}</h2>
      <p style="font-size:13px;color:#605e5c;margin-bottom:16px;">
        ${_e(attempt.agent)}${m?.department ? ` · ${_e(m.department)}` : ''}${m?.location ? ` · ${_e(m.location)}` : ''}
        &nbsp;·&nbsp; Completed: ${_fmtDate(attempt.completedDate)}
      </p>
      <div style="display:inline-block;padding:8px 18px;border-radius:8px;font-size:20px;font-weight:700;margin-bottom:20px;${pass ? 'background:#dff6dd;color:#107c10;' : 'background:#fde7e9;color:#d83b01;'}">
        ${pass ? 'Passed' : 'Failed'} — ${attempt.score}%
      </div>
      ${responses.length
        ? responses.map((r, i) => `
          <div style="margin-bottom:10px;padding:12px 14px;border-radius:7px;border-left:4px solid ${r.isCorrect ? '#107c10' : '#d83b01'};background:${r.isCorrect ? '#f1f8ee' : '#fdf3f2'};">
            <div style="font-weight:600;font-size:13px;margin-bottom:4px;">Q${i + 1}: ${_e(r.questionText)}</div>
            <div style="font-size:13px;">Your answer: <strong>${_e(r.answer || '(blank)')}</strong></div>
            ${!r.isCorrect ? `<div style="font-size:13px;">Correct: <strong>${_e(_fmtCA(r.correctAnswer))}</strong></div>` : ''}
            ${r.reason ? `<div style="font-size:12px;color:#605e5c;margin-top:3px;"><em>${_e(r.reason)}</em></div>` : ''}
            <div style="font-weight:700;margin-top:5px;font-size:13px;color:${r.isCorrect ? '#107c10' : '#d83b01'};">${r.isCorrect ? '✓ Correct' : '✗ Incorrect'}</div>
          </div>`).join('')
        : '<p style="color:#605e5c;">No answer detail stored for this attempt.</p>'
      }
    </div>
  `;

  document.body.appendChild(overlay);
  document.getElementById('modal-close').onclick = () => overlay.remove();
  overlay.onclick = (e) => { if (e.target === overlay) overlay.remove(); };
}

async function _renderAnalytics(container) {
  const [data, dir] = await Promise.all([_loadAll(), _loadAgentDir()]);
  const agentMeta   = new Map(dir.map(a => [a.name.toLowerCase(), a]));
  const deptDisplay = new Map(dir.map(a => [a.department.toLowerCase(), a.department]).filter(([k]) => k));
  const locDisplay  = new Map(dir.map(a => [a.location.toLowerCase(),  a.location ]).filter(([k]) => k));

  const SH = (title, exportId) => `
    <div style="display:flex;justify-content:space-between;align-items:center;border-bottom:2px solid var(--primary);padding-bottom:8px;margin-bottom:14px;">
      <h2 style="font-size:16px;color:var(--text);border:none;padding:0;margin:0;">${title}</h2>
      <button class="btn-sm" id="${exportId}">Export CSV</button>
    </div>`;

  container.innerHTML = `
    <!-- Filters -->
    <div class="section">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;">
        <h2 style="font-size:16px;color:var(--text);border:none;padding:0;margin:0;">Filters</h2>
        <button class="btn-sm btn-warn" id="flt-clear">Clear All</button>
      </div>
      <div class="form-row" style="flex-wrap:wrap; gap:8px;">
        <select id="flt-quiz" style="flex:1; min-width:160px;">
          <option value="">All Quizzes</option>
          ${data.quizzes.map(q => `<option value="${q.id}">${_e(q.title)}</option>`).join('')}
        </select>
        <select id="flt-dept"     style="flex:1; min-width:130px;"><option value="">All Departments</option></select>
        <select id="flt-loc"      style="flex:1; min-width:130px;"><option value="">All Locations</option></select>
        <select id="flt-hiredate" style="flex:1; min-width:150px;"><option value="">Any Hire Date</option></select>
        <select id="flt-creator" style="flex:1; min-width:150px;"><option value="">All Creators</option></select>
        <div style="flex:1; min-width:160px;">
          <input id="flt-agent" type="text" list="flt-agent-list" placeholder="Agent name…" style="width:100%;">
          <datalist id="flt-agent-list"></datalist>
        </div>
      </div>
    </div>

    <!-- Stats -->
    <div class="stats-grid" style="grid-template-columns:repeat(auto-fit,minmax(130px,1fr));">
      <div class="stat"><div class="stat-lbl">Total Attempts</div><div class="stat-val" id="s-total">—</div></div>
      <div class="stat"><div class="stat-lbl">Passed</div><div class="stat-val pass" id="s-pass">—</div></div>
      <div class="stat"><div class="stat-lbl">Failed</div><div class="stat-val fail" id="s-fail">—</div></div>
      <div class="stat"><div class="stat-lbl">Pass Rate</div><div class="stat-val" id="s-rate">—</div></div>
      <div class="stat"><div class="stat-lbl">Avg Score</div><div class="stat-val" id="s-avg">—</div></div>
      <div class="stat"><div class="stat-lbl">Unique Agents</div><div class="stat-val" id="s-agents">—</div></div>
    </div>

    <!-- By Agent -->
    <div class="section">
      ${SH('By Agent <span style="font-size:12px;font-weight:400;color:#605e5c;">(click row to expand attempts · click attempt to view answers)</span>', 'export-agent-sum')}
      <div class="table-wrap">
        <table>
          <thead><tr><th style="width:22px;"></th><th>Agent</th><th>Dept</th><th>Location</th><th>Batch</th><th>Attempts</th><th>Passed</th><th>Failed</th><th>Avg Score</th><th>Best</th></tr></thead>
          <tbody id="tbl-agent-sum"></tbody>
        </table>
      </div>
    </div>

    <!-- Results -->
    <div class="section">
      ${SH('Results <span style="font-size:12px;font-weight:400;color:#605e5c;">(click any row to view answers)</span>', 'export-results')}
      <div class="table-wrap">
        <table>
          <thead><tr><th>Quiz</th><th>Agent</th><th>Dept</th><th>Location</th><th>Batch</th><th>Score</th><th>Result</th><th>Date</th></tr></thead>
          <tbody id="tbl-res"></tbody>
        </table>
      </div>
    </div>

    <!-- By Department -->
    <div class="section">
      ${SH('By Department', 'export-dept')}
      <div class="table-wrap">
        <table>
          <thead><tr><th>Department</th><th>Attempts</th><th>Agents</th><th>Passed</th><th>Failed</th><th>Pass Rate</th><th>Avg Score</th></tr></thead>
          <tbody id="tbl-dept"></tbody>
        </table>
      </div>
    </div>

    <!-- By Location -->
    <div class="section">
      ${SH('By Location', 'export-loc')}
      <div class="table-wrap">
        <table>
          <thead><tr><th>Location</th><th>Attempts</th><th>Agents</th><th>Passed</th><th>Failed</th><th>Pass Rate</th><th>Avg Score</th></tr></thead>
          <tbody id="tbl-loc"></tbody>
        </table>
      </div>
    </div>

    <!-- By Hire Batch -->
    <div class="section">
      ${SH('By Hire Batch', 'export-batch')}
      <div class="table-wrap">
        <table>
          <thead><tr><th>Hire Batch</th><th>Attempts</th><th>Agents</th><th>Passed</th><th>Failed</th><th>Pass Rate</th><th>Avg Score</th></tr></thead>
          <tbody id="tbl-batch"></tbody>
        </table>
      </div>
    </div>

    <!-- Question Difficulty -->
    <div class="section">
      ${SH('Question Difficulty <span style="font-size:12px;font-weight:400;color:#605e5c;">(hardest first)</span>', 'export-qdiff')}
      <div class="table-wrap">
        <table>
          <thead><tr><th>Question</th><th>Quiz</th><th>Created By</th><th>Times Asked</th><th>Correct Rate</th></tr></thead>
          <tbody id="tbl-qdiff"></tbody>
        </table>
      </div>
    </div>
  `;

  const quizMap = new Map(data.quizzes.map(q => [q.id, q]));
  const EMPTY   = (cols) => `<tr><td colspan="${cols}" class="muted tc">No data.</td></tr>`;

  // Populate creator dropdown from non-deleted quizzes
  const creatorOpts = [...new Set(data.quizzes.filter(q => !q.isDeleted && q.createdByName).map(q => q.createdByName))].sort();
  const creatorSel  = document.getElementById('flt-creator');
  if (creatorSel) creatorSel.innerHTML = '<option value="">All Creators</option>' +
    creatorOpts.map(c => `<option value="${_e(c)}">${_e(c)}</option>`).join('');

  let _sorted = [], _deptMap = {}, _locMap = {}, _batchMap = {}, _agentSumMap = {}, _agSumRows = [], _qdRows = [];

  // ── Cascading filter helpers ──
  function _dirFor(skip) {
    const dept     = skip === 'dept'     ? '' : (document.getElementById('flt-dept').value     || '');
    const loc      = skip === 'loc'      ? '' : (document.getElementById('flt-loc').value      || '');
    const hireDate = skip === 'hiredate' ? '' : (document.getElementById('flt-hiredate').value || '');
    return dir.filter(a => {
      if (dept     && a.department.toLowerCase() !== dept)       return false;
      if (loc      && a.location.toLowerCase()   !== loc)        return false;
      if (hireDate && _normDateKey(a.hireDate)   !== hireDate)   return false;
      return true;
    });
  }

  function _rebuildSel(id, rawVals, placeholder, displayFn) {
    const el = document.getElementById(id);
    if (!el) return;
    const cur    = el.value;
    const unique = [...new Set(rawVals.filter(Boolean))].sort().reverse();
    el.innerHTML = [`<option value="">${placeholder}</option>`,
      ...unique.map(v => `<option value="${_e(v)}"${v === cur ? ' selected' : ''}>${_e(displayFn ? displayFn(v) : v)}</option>`)
    ].join('');
    if (cur && el.value !== cur) el.value = '';
  }

  function updateCascade() {
    _rebuildSel('flt-dept',     _dirFor('dept').map(a => a.department.toLowerCase()),    'All Departments', v => deptDisplay.get(v) || v);
    _rebuildSel('flt-loc',      _dirFor('loc').map(a => a.location.toLowerCase()),       'All Locations',   v => locDisplay.get(v)  || v);
    _rebuildSel('flt-hiredate', _dirFor('hiredate').map(a => _normDateKey(a.hireDate)),  'Any Hire Date',   v => _fmtDateShort(v));
    const dl    = document.getElementById('flt-agent-list');
    const names = [...new Set(_dirFor(null).map(a => a.name))].sort();
    if (dl) dl.innerHTML = names.map(n => `<option value="${_e(n)}">`).join('');
  }

  // ── Breakdown helpers ──
  function _groupStats(map, key, attempt) {
    if (!map[key]) map[key] = { attempts: 0, agents: new Set(), passed: 0, scores: [] };
    map[key].attempts++;
    map[key].agents.add(attempt.agent.toLowerCase());
    map[key].scores.push(attempt.score);
    if (attempt.score >= PASS_SCORE) map[key].passed++;
  }

  function _breakdownRow(key, s) {
    const avg  = (s.scores.reduce((x, y) => x + y, 0) / s.scores.length).toFixed(1);
    const rate = ((s.passed / s.attempts) * 100).toFixed(1);
    return `<tr>
      <td>${_e(key)}</td>
      <td>${s.attempts}</td>
      <td>${s.agents.size}</td>
      <td class="pass">${s.passed}</td>
      <td class="fail">${s.attempts - s.passed}</td>
      <td>${rate}%</td>
      <td>${avg}%</td>
    </tr>`;
  }

  function _breakdownCsv(map, sortFn, header) {
    const rows = [header];
    Object.entries(map).sort(sortFn).forEach(([k, s]) => {
      const avg  = (s.scores.reduce((x, y) => x + y, 0) / s.scores.length).toFixed(1);
      const rate = ((s.passed / s.attempts) * 100).toFixed(1);
      rows.push([k, s.attempts, s.agents.size, s.passed, s.attempts - s.passed, `${rate}%`, `${avg}%`]);
    });
    return rows;
  }

  function applyFilters() {
    const qId      = Number(document.getElementById('flt-quiz').value      || 0);
    const dept     = (document.getElementById('flt-dept').value     || '').trim();
    const loc      = (document.getElementById('flt-loc').value      || '').trim();
    const hireDate = (document.getElementById('flt-hiredate').value || '').trim();
    const creator  = (document.getElementById('flt-creator').value  || '').trim();
    const agent    = (document.getElementById('flt-agent').value    || '').trim().toLowerCase();

    const filtered = data.attempts.filter(a => {
      if (qId     && a.quizId !== qId) return false;
      if (creator && quizMap.get(a.quizId)?.createdByName !== creator) return false;
      if (agent   && !a.agent.toLowerCase().includes(agent)) return false;
      const m = agentMeta.get(a.agent.toLowerCase());
      if (dept     && (!m || m.department.toLowerCase() !== dept))         return false;
      if (loc      && (!m || m.location.toLowerCase()   !== loc))          return false;
      if (hireDate && (!m || _normDateKey(m.hireDate)   !== hireDate))     return false;
      return true;
    });

    // ── Stats ──
    const total  = filtered.length;
    const passed = filtered.filter(a => a.score >= PASS_SCORE).length;
    const avg    = total ? (filtered.reduce((s, a) => s + a.score, 0) / total).toFixed(1) : 0;
    const rate   = total ? ((passed / total) * 100).toFixed(1) : 0;
    const unique = new Set(filtered.map(a => a.agent.toLowerCase())).size;
    document.getElementById('s-total').textContent  = total;
    document.getElementById('s-pass').textContent   = passed;
    document.getElementById('s-fail').textContent   = total - passed;
    document.getElementById('s-rate').textContent   = total ? `${rate}%` : '—';
    document.getElementById('s-avg').textContent    = total ? `${avg}%`  : '—';
    document.getElementById('s-agents').textContent = unique || '—';

    // ── Results table ──
    _sorted = filtered.slice().sort((a, b) => new Date(b.completedDate || 0) - new Date(a.completedDate || 0));
    const resTbody = document.getElementById('tbl-res');
    resTbody.innerHTML = _sorted.length
      ? _sorted.map(a => {
          const quiz = quizMap.get(a.quizId);
          const pass = a.score >= PASS_SCORE;
          const m    = agentMeta.get(a.agent.toLowerCase());
          const hb   = _hireDateToBatch(m?.hireDate || '');
          return `<tr data-aid="${a.id}" style="cursor:pointer;" title="Click to view answers">
            <td>${_e(quiz?.title || a.templateKey)}</td>
            <td>${_e(a.agent)}</td>
            <td>${_e(m?.department || '—')}</td>
            <td>${_e(m?.location   || '—')}</td>
            <td>${_e(hb || '—')}</td>
            <td><strong>${a.score}%</strong></td>
            <td class="${pass ? 'pass' : 'fail'}">${pass ? 'Passed' : 'Failed'}</td>
            <td>${_fmtDate(a.completedDate)}</td>
          </tr>`;
        }).join('')
      : EMPTY(8);

    resTbody.querySelectorAll('tr[data-aid]').forEach(row => {
      row.onmouseenter = () => { row.style.background = '#eff6fc'; };
      row.onmouseleave = () => { row.style.background = ''; };
      row.onclick = () => {
        const a = _sorted.find(x => x.id === Number(row.dataset.aid));
        if (a) _showAnswerModal(a, quizMap.get(a.quizId), agentMeta);
      };
    });

    // ── By Department ──
    _deptMap = {};
    filtered.forEach(a => {
      const m = agentMeta.get(a.agent.toLowerCase());
      _groupStats(_deptMap, m?.department || '(Unknown)', a);
    });
    document.getElementById('tbl-dept').innerHTML = Object.keys(_deptMap).length
      ? Object.entries(_deptMap).sort((a, b) => b[1].attempts - a[1].attempts).map(([k, s]) => _breakdownRow(k, s)).join('')
      : EMPTY(7);

    // ── By Location ──
    _locMap = {};
    filtered.forEach(a => {
      const m = agentMeta.get(a.agent.toLowerCase());
      _groupStats(_locMap, m?.location || '(Unknown)', a);
    });
    document.getElementById('tbl-loc').innerHTML = Object.keys(_locMap).length
      ? Object.entries(_locMap).sort((a, b) => b[1].attempts - a[1].attempts).map(([k, s]) => _breakdownRow(k, s)).join('')
      : EMPTY(7);

    // ── By Hire Batch ──
    _batchMap = {};
    filtered.forEach(a => {
      const m   = agentMeta.get(a.agent.toLowerCase());
      const key = _hireDateToBatch(m?.hireDate || '') || '(No batch)';
      _groupStats(_batchMap, key, a);
    });
    document.getElementById('tbl-batch').innerHTML = Object.keys(_batchMap).length
      ? Object.entries(_batchMap).sort((a, b) => b[0].localeCompare(a[0])).map(([k, s]) => _breakdownRow(k, s)).join('')
      : EMPTY(7);

    // ── By Agent ──
    _agentSumMap = {};
    filtered.forEach(a => {
      const m = agentMeta.get(a.agent.toLowerCase());
      if (!_agentSumMap[a.agent]) {
        _agentSumMap[a.agent] = {
          dept: m?.department || '', loc: m?.location || '',
          batch: _hireDateToBatch(m?.hireDate || ''),
          attempts: [], passed: 0, scores: []
        };
      }
      _agentSumMap[a.agent].attempts.push(a);
      _agentSumMap[a.agent].scores.push(a.score);
      if (a.score >= PASS_SCORE) _agentSumMap[a.agent].passed++;
    });
    _agSumRows = Object.entries(_agentSumMap).sort((a, b) => a[0].localeCompare(b[0]));

    const agSumTbody = document.getElementById('tbl-agent-sum');
    if (_agSumRows.length) {
      agSumTbody.innerHTML = _agSumRows.map(([name, s], idx) => {
        const avg  = (s.scores.reduce((x, y) => x + y, 0) / s.scores.length).toFixed(1);
        const best = Math.max(...s.scores);
        const sortedAttempts = s.attempts.slice().sort((a, b) => new Date(b.completedDate || 0) - new Date(a.completedDate || 0));
        const detailRows = sortedAttempts.map(a => {
          const quiz = quizMap.get(a.quizId);
          const pass = a.score >= PASS_SCORE;
          return `<tr data-aaid="${a.id}" style="cursor:pointer;background:#fff;" title="Click to view answers">
            <td colspan="2" style="padding:8px 8px 8px 28px;">${_e(quiz?.title || a.templateKey)}</td>
            <td style="padding:8px 12px;"><strong>${a.score}%</strong></td>
            <td style="padding:8px 12px;" class="${pass ? 'pass' : 'fail'}">${pass ? 'Passed' : 'Failed'}</td>
            <td colspan="6" style="padding:8px 12px;">${_fmtDate(a.completedDate)}</td>
          </tr>`;
        }).join('');
        return `
          <tr class="agent-sum-row" data-agidx="${idx}" style="cursor:pointer;">
            <td style="font-size:11px;color:#605e5c;text-align:center;">▶</td>
            <td><strong>${_e(name)}</strong></td>
            <td>${_e(s.dept || '—')}</td>
            <td>${_e(s.loc  || '—')}</td>
            <td>${_e(s.batch || '—')}</td>
            <td>${s.scores.length}</td>
            <td class="pass">${s.passed}</td>
            <td class="fail">${s.scores.length - s.passed}</td>
            <td>${avg}%</td>
            <td>${best}%</td>
          </tr>
          <tr class="agent-detail-row" data-agdetail="${idx}" style="display:none;background:#f0f6ff;">
            <td colspan="10" style="padding:0;">
              <table style="width:100%;font-size:12px;border-collapse:collapse;">
                <thead><tr style="background:#dce9f7;">
                  <th colspan="2" style="padding:7px 8px 7px 28px;text-align:left;font-size:11px;font-weight:700;color:#605e5c;text-transform:uppercase;letter-spacing:.03em;">Quiz</th>
                  <th style="padding:7px 12px;text-align:left;font-size:11px;font-weight:700;color:#605e5c;text-transform:uppercase;letter-spacing:.03em;">Score</th>
                  <th style="padding:7px 12px;text-align:left;font-size:11px;font-weight:700;color:#605e5c;text-transform:uppercase;letter-spacing:.03em;">Result</th>
                  <th colspan="6" style="padding:7px 12px;text-align:left;font-size:11px;font-weight:700;color:#605e5c;text-transform:uppercase;letter-spacing:.03em;">Date</th>
                </tr></thead>
                <tbody>${detailRows}</tbody>
              </table>
            </td>
          </tr>`;
      }).join('');

      agSumTbody.querySelectorAll('.agent-sum-row').forEach(row => {
        row.onmouseenter = () => { if (row.style.background !== '#eff6fc') row.style.background = '#faf9f8'; };
        row.onmouseleave = () => { const det = agSumTbody.querySelector(`.agent-detail-row[data-agdetail="${row.dataset.agidx}"]`); row.style.background = det?.style.display !== 'none' ? '#eff6fc' : ''; };
        row.onclick = () => {
          const idx    = row.dataset.agidx;
          const detail = agSumTbody.querySelector(`.agent-detail-row[data-agdetail="${idx}"]`);
          const arrow  = row.querySelector('td:first-child');
          const open   = detail.style.display !== 'none';
          detail.style.display = open ? 'none' : '';
          if (arrow) arrow.textContent = open ? '▶' : '▼';
          row.style.background = open ? '' : '#eff6fc';
        };
      });

      agSumTbody.querySelectorAll('tr[data-aaid]').forEach(row => {
        row.onmouseenter = () => { row.style.background = '#fffbe6'; };
        row.onmouseleave = () => { row.style.background = '#fff'; };
        row.onclick = (e) => {
          e.stopPropagation();
          const a = filtered.find(x => x.id === Number(row.dataset.aaid));
          if (a) _showAnswerModal(a, quizMap.get(a.quizId), agentMeta);
        };
      });
    } else {
      agSumTbody.innerHTML = EMPTY(10);
    }

    // ── Question Difficulty ──
    const qdMap = {};
    filtered.forEach(a => {
      const quiz = quizMap.get(a.quizId);
      (a.answers?.responses || []).forEach(r => {
        if (!r.questionText) return;
        if (!qdMap[r.questionText]) qdMap[r.questionText] = { quizTitle: quiz?.title || '', createdByName: quiz?.createdByName || '', total: 0, correct: 0 };
        qdMap[r.questionText].total++;
        if (r.isCorrect) qdMap[r.questionText].correct++;
      });
    });
    _qdRows = Object.entries(qdMap).sort((a, b) => (a[1].correct / a[1].total) - (b[1].correct / b[1].total));
    document.getElementById('tbl-qdiff').innerHTML = _qdRows.length
      ? _qdRows.map(([q, s]) => {
          const pct = ((s.correct / s.total) * 100).toFixed(1);
          const cls = parseFloat(pct) < 50 ? 'fail' : parseFloat(pct) < PASS_SCORE ? '' : 'pass';
          return `<tr>
            <td>${_e(q)}</td>
            <td>${_e(s.quizTitle)}</td>
            <td>${s.createdByName ? _creatorBadge(s.createdByName) : '—'}</td>
            <td>${s.total}</td>
            <td class="${cls}"><strong>${pct}%</strong></td>
          </tr>`;
        }).join('')
      : EMPTY(5);

    // ── CSV wiring (re-bind after each filter so closures capture fresh data) ──
    document.getElementById('export-results').onclick = () => {
      const rows = [['Quiz','Agent','Department','Location','Hire Batch','Score','Result','Date']];
      _sorted.forEach(a => {
        const quiz = quizMap.get(a.quizId);
        const m    = agentMeta.get(a.agent.toLowerCase());
        rows.push([
          quiz?.title || a.templateKey, a.agent,
          m?.department || '', m?.location || '',
          _hireDateToBatch(m?.hireDate || ''),
          `${a.score}%`, a.score >= PASS_SCORE ? 'Passed' : 'Failed',
          _fmtDate(a.completedDate)
        ]);
      });
      _csvDownload('quiz-results.csv', rows);
    };
    document.getElementById('export-dept').onclick  = () =>
      _csvDownload('quiz-by-department.csv', _breakdownCsv(_deptMap,  (a, b) => b[1].attempts - a[1].attempts, ['Department','Attempts','Agents','Passed','Failed','Pass Rate','Avg Score']));
    document.getElementById('export-loc').onclick   = () =>
      _csvDownload('quiz-by-location.csv',   _breakdownCsv(_locMap,   (a, b) => b[1].attempts - a[1].attempts, ['Location','Attempts','Agents','Passed','Failed','Pass Rate','Avg Score']));
    document.getElementById('export-batch').onclick = () =>
      _csvDownload('quiz-by-batch.csv',      _breakdownCsv(_batchMap, (a, b) => b[0].localeCompare(a[0]),      ['Hire Batch','Attempts','Agents','Passed','Failed','Pass Rate','Avg Score']));
    document.getElementById('export-agent-sum').onclick = () => {
      const rows = [['Agent','Department','Location','Hire Batch','Attempts','Passed','Failed','Avg Score','Best Score']];
      _agSumRows.forEach(([name, s]) => {
        const avg = (s.scores.reduce((x, y) => x + y, 0) / s.scores.length).toFixed(1);
        rows.push([name, s.dept, s.loc, s.batch, s.scores.length, s.passed, s.scores.length - s.passed, `${avg}%`, `${Math.max(...s.scores)}%`]);
      });
      _csvDownload('quiz-by-agent.csv', rows);
    };
    document.getElementById('export-qdiff').onclick = () => {
      const rows = [['Question','Quiz','Created By','Times Asked','Correct Rate']];
      _qdRows.forEach(([q, s]) => rows.push([q, s.quizTitle, s.createdByName || '', s.total, `${((s.correct/s.total)*100).toFixed(1)}%`]));
      _csvDownload('quiz-question-difficulty.csv', rows);
    };
  }

  // ── Live filter wiring ──
  ['flt-quiz', 'flt-dept', 'flt-loc', 'flt-hiredate', 'flt-creator'].forEach(id => {
    document.getElementById(id).addEventListener('change', () => { updateCascade(); applyFilters(); });
  });
  document.getElementById('flt-agent').addEventListener('input', applyFilters);
  document.getElementById('flt-clear').addEventListener('click', () => {
    ['flt-quiz', 'flt-dept', 'flt-loc', 'flt-hiredate', 'flt-creator'].forEach(id => { document.getElementById(id).value = ''; });
    document.getElementById('flt-agent').value = '';
    updateCascade();
    applyFilters();
  });

  updateCascade();
  applyFilters();
}


// ═════════════════════════════════════════════════════════════
// AGENT DASHBOARD
// ═════════════════════════════════════════════════════════════

async function _initAgent() {
  const root = document.getElementById('agent-root');
  if (!root) return;
  try {
    const token = _qp('token');
    if (token) await _handleToken(root, token);
    else        await _handleNormal(root);
  } catch (e) {
    root.innerHTML = `<p class="error">Something went wrong. Please refresh the page.<br>
      <small>${_e(e.message)}</small></p>`;
    console.error('[Quiz]', e);
  }
}


// ── TOKEN MODE ──

async function _handleToken(root, token) {
  root.innerHTML = '<p class="muted loading-msg">Validating your quiz link…</p>';

  let linkItems;
  try {
    linkItems = await _spFetch(_L.links, { filter: `Title eq '${_spStr(token)}'` });
  } catch (e) {
    root.innerHTML = '<p class="error">Unable to validate this link right now. Please try again or contact your supervisor.</p>';
    return;
  }

  if (!linkItems.length) {
    root.innerHTML = '<p class="error">This link is not valid. Please contact your supervisor.</p>';
    return;
  }

  const link = _parseLink(linkItems[0]);
  if (!link.isActive) {
    root.innerHTML = '<p class="error">This link has been deactivated. Please contact your supervisor for an updated link.</p>';
    return;
  }

  await _showAgentSelector(root, async (agentName) => {
    await _runQuiz(root, link.quizId, agentName, token);
  });
}


// ── NORMAL MODE (no token — agent browses to the page directly) ──

async function _handleNormal(root) {
  await _showAgentSelector(root, async (agentName) => {
    root.innerHTML = '<p class="muted loading-msg">Loading your quizzes…</p>';

    const [quizItems, dir] = await Promise.all([
      _spFetch(_L.quizzes, { filter: `IsActive eq 'Yes'` }),
      _loadAgentDir()
    ]);
    const quizzes = quizItems.map(_parseQuiz);
    const info      = dir.find(a => a.name.toLowerCase() === agentName.toLowerCase());
    const dept      = (info?.department || '').toLowerCase();
    const loc       = (info?.location   || '').toLowerCase();
    const hireBatch = _hireDateToBatch(info?.hireDate);

    const available = quizzes.filter(q => {
      const deptOk  = !q.targetDepartment || q.targetDepartment.toLowerCase() === dept;
      const locOk   = !q.targetLocation   || q.targetLocation.toLowerCase()   === loc;
      // Batch: if quiz targets a hire cohort, only show to agents hired that month
      const batchOk = !q.batch || !hireBatch || q.batch.toLowerCase() === hireBatch;
      return deptOk && locOk && batchOk;
    });

    if (!available.length) {
      root.innerHTML = `
        <div class="agent-hdr">
          <h2>Welcome, ${_e(agentName)}</h2>
          <button class="btn-sm" onclick="sessionStorage.removeItem('quizAgent');location.reload()">Change Name</button>
        </div>
        <p class="muted">No quizzes are currently available for you. Check back soon or contact your supervisor.</p>
      `;
      return;
    }

    if (available.length === 1) {
      await _runQuiz(root, available[0].id, agentName, null);
      return;
    }

    // Multiple active quizzes — show picker
    root.innerHTML = `
      <div class="agent-hdr">
        <h2>Welcome, ${_e(agentName)}</h2>
        <button class="btn-sm" onclick="sessionStorage.removeItem('quizAgent');location.reload()">Change Name</button>
      </div>
      <p>You have multiple quizzes available. Select one to begin:</p>
      <div id="quiz-pick"></div>
    `;
    const pickEl = document.getElementById('quiz-pick');
    pickEl.innerHTML = available.map(q => `
      <div class="list-item">
        <div>
          <strong>${_e(q.title)}</strong>
          ${q.description ? `<div class="muted small">${_e(q.description)}</div>` : ''}
        </div>
        <button data-qid="${q.id}">Start Quiz →</button>
      </div>
    `).join('');
    pickEl.onclick = async (e) => {
      const btn = e.target.closest('button[data-qid]');
      if (btn) await _runQuiz(root, Number(btn.dataset.qid), agentName, null);
    };
  });
}


// ── SHARED QUIZ RUNNER ──

async function _runQuiz(root, quizId, agentName, token) {
  root.innerHTML = '<p class="muted loading-msg">Loading quiz…</p>';

  const [quizItems, qItems, attItems] = await Promise.all([
    _spFetch(_L.quizzes,  { filter: `ID eq ${quizId}` }),
    _spFetch(_L.questions),
    _spFetch(_L.attempts, { filter: `Agent eq '${_spStr(agentName)}' and QuizId eq ${quizId}` })
  ]);

  const quiz = quizItems.map(_parseQuiz)[0];
  if (!quiz) {
    root.innerHTML = '<p class="error">Quiz not found. Please contact your supervisor.</p>';
    return;
  }

  // Agent already submitted → show their result
  if (attItems.length) {
    const att = _parseAttempt(attItems.sort((a, b) => b.ID - a.ID)[0]);
    _showResults(root, att, quiz, agentName);
    return;
  }

  // Build question list (active questions for this template, in order)
  const questions = qItems.map(_parseQuestion)
    .filter(q => q.templateKey === quiz.templateKey && q.isActive)
    .sort((a, b) => (a.sortOrder - b.sortOrder) || (a.id - b.id));

  if (!questions.length) {
    root.innerHTML = '<p class="error">This quiz has no active questions. Please contact your supervisor.</p>';
    return;
  }

  root.innerHTML = `
    <div class="agent-hdr">
      <div>
        <h2>${_e(quiz.title)}</h2>
        ${quiz.description ? `<p class="muted">${_e(quiz.description)}</p>` : ''}
      </div>
      <button class="btn-sm" onclick="sessionStorage.removeItem('quizAgent');location.reload()">Change Name</button>
    </div>
    <p class="muted"><strong>Agent:</strong> ${_e(agentName)} &nbsp;·&nbsp; <strong>Pass mark:</strong> ${PASS_SCORE}%</p>
    <div id="quiz-progress-wrap" style="position:sticky;top:0;z-index:10;background:#fff;padding:10px 0 8px;margin-bottom:8px;border-bottom:1px solid #e1dfdd;">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
        <span style="font-size:13px;font-weight:600;color:#605e5c;">Progress</span>
        <span id="quiz-progress-txt" style="font-size:13px;font-weight:700;color:#0078d4;">0 of ${questions.length} answered</span>
      </div>
      <div style="height:6px;background:#e1dfdd;border-radius:4px;overflow:hidden;">
        <div id="quiz-progress-bar" style="height:100%;width:0%;background:#0078d4;border-radius:4px;transition:width .2s;"></div>
      </div>
    </div>
    <form id="quiz-form" autocomplete="off">
      ${questions.map((q, i) => `
        <div class="q-block">
          <p class="q-num">Question ${i + 1} of ${questions.length}</p>
          <p class="q-text">${_e(q.questionText)}</p>
          ${q.questionType === 'text'
            ? `<textarea name="q-${q.id}" rows="4" placeholder="Type your answer here…"></textarea>`
            : `${q.correctAnswers.length > 1 ? '<p class="muted small" style="margin-bottom:8px;">Select all that apply.</p>' : ''}` +
              q.options.map(opt => `
                <label class="radio-lbl">
                  <input type="${q.correctAnswers.length > 1 ? 'checkbox' : 'radio'}" name="q-${q.id}" value="${_e(opt)}">
                  <span>${_e(opt)}</span>
                </label>`).join('')
          }
        </div>
      `).join('')}
      <div style="margin-top:20px; padding-top:16px; border-top:1px solid #e1dfdd;">
        <button type="button" id="submit-btn">Submit Answers</button>
        <span class="muted small" style="margin-left:12px;">Answer every question before submitting.</span>
      </div>
    </form>
  `;

  document.getElementById('submit-btn').onclick = async () => {
    await _submitQuiz(root, quiz, questions, agentName, token);
  };

  // Live progress counter
  function _updateProgress() {
    const total     = questions.length;
    const answered  = questions.filter(q => {
      if (q.questionType === 'text') {
        return (document.querySelector(`textarea[name="q-${q.id}"]`)?.value || '').trim() !== '';
      }
      return !!document.querySelector(`input[name="q-${q.id}"]:checked`);
    }).length;
    const pct = total ? (answered / total * 100).toFixed(0) : 0;
    const txt = document.getElementById('quiz-progress-txt');
    const bar = document.getElementById('quiz-progress-bar');
    if (txt) txt.textContent = `${answered} of ${total} answered`;
    if (bar) {
      bar.style.width   = `${pct}%`;
      bar.style.background = answered === total ? '#107c10' : '#0078d4';
    }
  }
  document.getElementById('quiz-form').addEventListener('change', _updateProgress);
  document.getElementById('quiz-form').addEventListener('input',  _updateProgress);
}


// ── SUBMIT ──

async function _submitQuiz(root, quiz, questions, agentName, token) {
  let correct = 0;
  const responses = questions.map(q => {
    let answer    = '';
    let isCorrect = false;
    if (q.questionType === 'text') {
      answer    = (document.querySelector(`textarea[name="q-${q.id}"]`)?.value || '').trim();
      isCorrect = _textAnswersMatch(answer, q.correctAnswer);
    } else {
      const cas = q.correctAnswers || [q.correctAnswer];
      if (cas.length > 1) {
        const selected = [...document.querySelectorAll(`input[name="q-${q.id}"]:checked`)].map(el => el.value).sort();
        const expected = [...cas].sort();
        isCorrect = selected.length === expected.length && selected.every((v, i) => v === expected[i]);
        answer    = selected.join(', ');
      } else {
        answer    = document.querySelector(`input[name="q-${q.id}"]:checked`)?.value || '';
        isCorrect = answer === q.correctAnswer;
      }
    }
    if (isCorrect) correct++;
    return {
      questionId   : q.id,
      questionText : q.questionText,
      answer,
      correctAnswer: q.correctAnswer,
      reason       : q.reason,
      questionType : q.questionType,
      isCorrect
    };
  });

  if (responses.some(r => !r.answer)) {
    alert('Please answer all questions before submitting.');
    return;
  }

  const score        = Math.round((correct / questions.length) * 100);
  const resultStatus = score >= PASS_SCORE ? 'Passed' : 'Failed';
  const batchKey     = _batch();

  const submitBtn = document.getElementById('submit-btn');
  if (submitBtn) { submitBtn.disabled = true; submitBtn.textContent = 'Submitting…'; }

  try {
    await _spCreate(_L.attempts, {
      Title        : `${_slugify(agentName)}-${quiz.id}-${Date.now()}`,
      Agent        : agentName,
      TemplateKey  : quiz.templateKey,
      QuizId       : quiz.id,
      BatchKey     : batchKey,
      Score        : score,
      CompletedDate: new Date().toISOString(),
      TokenUsed    : token || '',
      Answers      : JSON.stringify({ responses, resultStatus })
    });
  } catch (e) {
    if (submitBtn) { submitBtn.disabled = false; submitBtn.textContent = 'Submit Answers'; }
    alert('Unable to save your answers. Please try again.');
    return;
  }

  _showResults(root,
    { score, completedDate: new Date().toISOString(), answers: { responses, resultStatus } },
    quiz, agentName);
}


// ── SHOW RESULTS ──

function _showResults(root, attempt, quiz, agentName) {
  const pass      = attempt.score >= PASS_SCORE;
  const responses = attempt.answers?.responses || [];
  const status    = attempt.answers?.resultStatus || (pass ? 'Passed' : 'Failed');

  root.innerHTML = `
    <div class="agent-hdr">
      <div><h2>${_e(quiz.title)}</h2></div>
      <button class="btn-sm" onclick="sessionStorage.removeItem('quizAgent');location.reload()">Change Name</button>
    </div>
    <p class="muted"><strong>Agent:</strong> ${_e(agentName)} &nbsp;·&nbsp; Completed: ${_fmtDate(attempt.completedDate)}</p>

    <div class="result-card ${pass ? 'result-pass' : 'result-fail'}">
      <div class="result-status">${_e(status)}</div>
      <div class="result-score">${attempt.score}%</div>
      <div class="muted small" style="margin-top:6px;">Pass mark: ${PASS_SCORE}%</div>
    </div>

    ${responses.length ? `
      <h3 style="margin-top:28px; margin-bottom:14px;">Answer Review</h3>
      ${responses.map((r, i) => `
        <div class="fb-block ${r.isCorrect ? 'fb-ok' : 'fb-fail'}">
          <p class="fb-q">Q${i + 1}: ${_e(r.questionText)}</p>
          <p>Your answer: <strong>${_e(r.answer || '(not answered)')}</strong></p>
          ${!r.isCorrect ? `<p>Correct answer: <strong>${_e(_fmtCA(r.correctAnswer))}</strong></p>` : ''}
          ${r.reason ? `<p class="muted"><em>${_e(r.reason)}</em></p>` : ''}
          <p class="${r.isCorrect ? 'pass' : 'fail'} fb-verdict">
            ${r.isCorrect ? '✓ Correct' : '✗ Incorrect'}
          </p>
        </div>`).join('')}
    ` : ''}
  `;
}


// ── AGENT NAME SELECTOR ──

async function _fetchCurrentUserName() {
  // Try page context first (instant, no request)
  const fromCtx = _spDisplayName();
  if (fromCtx) return fromCtx;
  // Fallback: REST API — reliable on all SharePoint modern pages
  try {
    const r = await fetch(`${_site}/_api/web/currentuser?$select=Title`, {
      headers: { Accept: 'application/json;odata=nometadata' },
      credentials: 'same-origin'
    });
    if (r.ok) return (await r.json()).Title || '';
  } catch (_) {}
  return '';
}

async function _showAgentSelector(root, onSelect) {
  const saved = sessionStorage.getItem('quizAgent');
  if (saved) { onSelect(saved); return; }

  root.innerHTML = '<p class="muted loading-msg">Loading…</p>';

  const [dir, spName] = await Promise.all([_loadAgentDir(), _fetchCurrentUserName()]);

  const detected = spName
    ? dir.find(a => a.name.toLowerCase() === spName.toLowerCase()) || null
    : null;

  if (detected) {
    _showAutoDetected(root, detected, dir, onSelect);
  } else {
    _showGuidedSelector(root, dir, onSelect, spName);
  }
}

function _showAutoDetected(root, agent, dir, onSelect) {
  root.innerHTML = `
    <div class="profile-card">
      <div class="profile-welcome">Hey, <strong>${_e(agent.name)}</strong>!</div>
      <div class="profile-rows">
        ${agent.department ? `<div class="profile-row"><span class="profile-lbl">Department</span><span>${_e(agent.department)}</span></div>` : ''}
        ${agent.location   ? `<div class="profile-row"><span class="profile-lbl">Location</span><span>${_e(agent.location)}</span></div>`   : ''}
        ${agent.hireDate   ? `<div class="profile-row"><span class="profile-lbl">Hire Date</span><span>${_fmtDateShort(agent.hireDate)}</span></div>` : ''}
      </div>
      <div style="margin-top:20px;">
        <button id="autodet-confirm">Continue as ${_e(agent.name)} →</button>
      </div>
    </div>
  `;

  document.getElementById('autodet-confirm').onclick = () => {
    sessionStorage.setItem('quizAgent', agent.name);
    onSelect(agent.name);
  };
}

function _showGuidedSelector(root, dir, onSelect, spName) {
  const _detectedSpName = spName || '';

  function agentsFor(dept, loc, dateKey) {
    return dir.filter(a =>
      (!dept    || a.department.toLowerCase() === dept.toLowerCase()) &&
      (!loc     || a.location.toLowerCase()   === loc.toLowerCase())  &&
      (!dateKey || _normDateKey(a.hireDate)   === dateKey)
    );
  }

  function locsFor(dept) {
    return [...new Set(
      dir.filter(a => !dept || a.department.toLowerCase() === dept.toLowerCase())
         .map(a => a.location).filter(Boolean)
    )].sort();
  }

  function datesFor(dept, loc) {
    return [...new Set(
      dir.filter(a =>
        (!dept || a.department.toLowerCase() === dept.toLowerCase()) &&
        (!loc  || a.location.toLowerCase()   === loc.toLowerCase())
      ).map(a => _normDateKey(a.hireDate)).filter(Boolean)
    )].sort().reverse();
  }

  const allDepts = [...new Set(dir.map(a => a.department).filter(Boolean))].sort();

  root.innerHTML = `
    <div class="profile-form">
      <h2 class="profile-form-title">Find Your Name</h2>
      ${_detectedSpName
        ? `<p class="muted" style="margin-bottom:6px; font-size:14px;">
             SharePoint shows you as <strong>${_e(_detectedSpName)}</strong> — no match found in the New Hire list.
             Make sure the name in the list is spelled exactly the same.
           </p>`
        : `<p class="muted" style="margin-bottom:6px; font-size:14px;">
             Could not read your SharePoint login name (this can happen on some pages).
           </p>`
      }
      <p class="muted" style="margin-bottom:20px; font-size:14px;">Use the filters below to find your name.</p>
      <div class="profile-row">
        <span class="profile-lbl">Department</span>
        <select id="gs-dept">
          <option value="">All Departments</option>
          ${allDepts.map(d => `<option value="${_e(d)}">${_e(d)}</option>`).join('')}
        </select>
      </div>
      <div class="profile-row">
        <span class="profile-lbl">Location</span>
        <select id="gs-loc">
          <option value="">All Locations</option>
        </select>
      </div>
      <div class="profile-row">
        <span class="profile-lbl">Hire Date</span>
        <select id="gs-date">
          <option value="">Any Hire Date</option>
        </select>
      </div>
      <div class="profile-row" style="align-items:flex-start;">
        <span class="profile-lbl" style="padding-top:4px;">Your Name</span>
        <select id="gs-name" size="6" style="flex:1;">
          <option value="" disabled>Select filters above first…</option>
        </select>
      </div>
      <button id="gs-confirm" style="margin-top:6px;">Continue →</button>
      ${!dir.length ? '<p class="muted small" style="margin-top:10px;">Directory unavailable — contact your supervisor.</p>' : ''}
    </div>
  `;

  function updateLocs() {
    const dept  = document.getElementById('gs-dept').value;
    const locs  = locsFor(dept);
    const locEl = document.getElementById('gs-loc');
    const prev  = locEl.value;
    locEl.innerHTML = '<option value="">All Locations</option>' +
      locs.map(l => `<option value="${_e(l)}"${l === prev ? ' selected' : ''}>${_e(l)}</option>`).join('');
    updateDates();
  }

  function updateDates() {
    const dept   = document.getElementById('gs-dept').value;
    const loc    = document.getElementById('gs-loc').value;
    const dates  = datesFor(dept, loc);
    const dateEl = document.getElementById('gs-date');
    const prev   = dateEl.value;
    dateEl.innerHTML = '<option value="">Any Hire Date</option>' +
      dates.map(dk => `<option value="${_e(dk)}"${dk === prev ? ' selected' : ''}>${_fmtDateShort(dk)}</option>`).join('');
    updateNames();
  }

  function updateNames() {
    const dept    = document.getElementById('gs-dept').value;
    const loc     = document.getElementById('gs-loc').value;
    const dateKey = document.getElementById('gs-date').value;
    const matches = agentsFor(dept, loc, dateKey);
    const nameEl  = document.getElementById('gs-name');
    nameEl.innerHTML = matches.length
      ? matches.map(a => `<option value="${_e(a.name)}">${_e(a.name)}</option>`).join('')
      : '<option value="" disabled>No matches — try fewer filters</option>';
  }

  document.getElementById('gs-dept').onchange  = updateLocs;
  document.getElementById('gs-loc').onchange   = updateDates;
  document.getElementById('gs-date').onchange  = updateNames;
  updateLocs();

  document.getElementById('gs-confirm').onclick = () => {
    const nameEl   = document.getElementById('gs-name');
    const selected = (nameEl?.value || '').trim();
    if (!selected) { alert('Please select your name from the list.'); return; }
    sessionStorage.setItem('quizAgent', selected);
    onSelect(selected);
  };
}


// ─────────────────────────────────────────────────────────────
// AUTO-INIT  (detects which page and runs the right dashboard)
// ─────────────────────────────────────────────────────────────

function _init() {
  if (document.getElementById('admin-root')) _initAdmin().catch(e => console.error('[Quiz Admin]', e));
  else if (document.getElementById('agent-root')) _initAgent().catch(e => console.error('[Quiz Agent]', e));
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', _init);
} else {
  _init();
}
