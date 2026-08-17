// ── State ──
let selectedFiles = [];
let selectedRefFiles = [];
let analysisResult = null;
let fileGroups = {}; // {filename: groupName} — same group = same bidder
let isAnalyzing = false;

// ── Analyzing lock (guard against mid-run state changes) ──
function _unloadGuard(e) { e.preventDefault(); e.returnValue = ''; }
function setAnalyzing(on) {
  isAnalyzing = on;
  document.querySelector('.app-container').classList.toggle('analyzing', on);
  if (on) window.addEventListener('beforeunload', _unloadGuard);
  else window.removeEventListener('beforeunload', _unloadGuard);
}

// ── DOM refs ──
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const fileList = document.getElementById('fileList');
const refUploadArea = document.getElementById('refUploadArea');
const refFileInput = document.getElementById('refFileInput');
const refFileList = document.getElementById('refFileList');
const btnAnalyze = document.getElementById('btnAnalyze');
const btnClear = document.getElementById('btnClear');
const btnDownload = document.getElementById('btnDownload');
const progressPanel = document.getElementById('progressPanel');
const progressFill = document.getElementById('progressFill');
const progressBar = progressFill.parentElement;  // the track (.progress-bar)
const progressText = document.getElementById('progressText');
const progressSteps = document.getElementById('progressSteps');
const resultsSection = document.getElementById('resultsSection');

// ── Bid File Selection ──
uploadArea.addEventListener('click', () => fileInput.click());
uploadArea.addEventListener('keydown', e => {
  if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); }
});
uploadArea.addEventListener('dragover', e => { e.preventDefault(); uploadArea.classList.add('drag-over'); });
uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('drag-over'));
uploadArea.addEventListener('drop', e => {
  e.preventDefault();
  uploadArea.classList.remove('drag-over');
  addFiles(e.dataTransfer.files, 'bid');
});
fileInput.addEventListener('change', e => addFiles(e.target.files, 'bid'));

// ── Reference File Selection ──
refUploadArea.addEventListener('click', () => refFileInput.click());
refUploadArea.addEventListener('keydown', e => {
  if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); refFileInput.click(); }
});
refUploadArea.addEventListener('dragover', e => { e.preventDefault(); refUploadArea.classList.add('drag-over'); });
refUploadArea.addEventListener('dragleave', () => refUploadArea.classList.remove('drag-over'));
refUploadArea.addEventListener('drop', e => {
  e.preventDefault();
  refUploadArea.classList.remove('drag-over');
  addFiles(e.dataTransfer.files, 'ref');
});
refFileInput.addEventListener('change', e => addFiles(e.target.files, 'ref'));

const VALID_EXTS = ['.docx', '.doc', '.pdf', '.txt'];

function escapeHtml(str) {
  // Escapes quotes too: results are interpolated into value="..." and
  // other double-quoted attributes, where a bare " would break out.
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function formatSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

// ── Progress bar helpers ──

// Highest bar percentage ever reached — the bar must never shrink.
var _progressMaxPct = 0;

function _barSet(pct) {
  pct = Math.max(0, Math.min(100, pct));
  if (pct > _progressMaxPct) _progressMaxPct = pct;
  progressFill.style.width = _progressMaxPct + '%';
}

function _ensureExtractStep(detail) {
  // Find or create the "提取文字" progress step so the user sees a
  // dedicated indicator for the (potentially long) extraction phase.
  var existing = progressSteps.querySelectorAll('.progress-step');
  var found = null;
  existing.forEach(function(el) {
    if (el.getAttribute('data-step') === 'extract') found = el;
  });
  if (!found) {
    found = document.createElement('span');
    found.className = 'progress-step';
    found.setAttribute('data-step', 'extract');
    progressSteps.appendChild(found);
  }
  found.textContent = detail || '提取文字';
  return found;
}

var _extractFileName = '';
var _extractStepEl = null;
var _fileIndex = 1;
var _totalFiles = 1;
var _fileShare = 24;

function startProgress() {
  _progressMaxPct = 0;
  _analysisFailed = false;
  _fileIndex = 1;
  _totalFiles = 1;
  _fileShare = 24;
  progressBar.classList.remove('extracting');
  progressPanel.classList.remove('error');
  progressPanel.style.display = 'block';
  progressFill.style.width = '0%';
  progressText.textContent = '上传中...';
  progressSteps.innerHTML = '<span class="progress-step active" data-step="upload">上传文件</span>';
  _extractStepEl = null;
}

function updateProgress(event) {
  // Entering the analysis phase: stop the extraction sweep. The sweep is
  // a transform-based overlay, so it never touched `width`; the bar's real
  // (monotonic) position is already correct - we only drop the class.
  // The similarity step is the slow O(n^2) difflib phase; a single large
  // file yields no inner events, so keep the sweep overlay alive so the
  // bar isn't dead-still for seconds. All other analysis steps are fast and
  // run clean (no sweep).
  if (event.step === 'similarity') {
    progressBar.classList.add('extracting');
  } else {
    progressBar.classList.remove('extracting');
  }
  _barSet(event.percent);
  progressText.textContent = event.label;

  // Mark the extraction step as completed (if it was created)
  if (_extractStepEl) {
    _extractStepEl.classList.remove('active');
    _extractStepEl.classList.add('done');
  }

  if (event.detail) {
    var existing = progressSteps.querySelectorAll('.progress-step');
    existing.forEach(function(el) { el.classList.remove('active'); });

    var found = null;
    existing.forEach(function(el) {
      if (el.getAttribute('data-step') === event.step) { found = el; }
    });
    if (!found) {
      found = document.createElement('span');
      found.className = 'progress-step';
      found.setAttribute('data-step', event.step);
      progressSteps.appendChild(found);
    }
    found.textContent = event.detail;
    found.classList.add('active');
  }
}

function updateExtractProgress(event) {
  if (event.phase === 'start') {
    _extractFileName = event.file || '';
    _extractStepEl = _ensureExtractStep('提取文字: ' + _extractFileName);
    var steps = progressSteps.querySelectorAll('.progress-step');
    steps.forEach(function(el) {
      if (el.getAttribute('data-step') === 'upload') {
        el.classList.remove('active'); el.classList.add('done');
      }
    });
    _extractStepEl.classList.remove('done');
    _extractStepEl.classList.add('active');
    // Each file owns an equal slice of the 1-25% extraction band. _barSet
    // is monotonic (it tracks the high-water mark), so file N starts where
    // file N-1 ended and the bar never moves backward across files.
    _fileIndex = event.fileIndex || 1;
    _totalFiles = event.totalFiles || 1;
    _fileShare = 24 / _totalFiles;  // extraction occupies the 1-25% band
    var fileStartPct = 1 + (_fileIndex - 1) * _fileShare;
    _barSet(fileStartPct);
    progressBar.classList.add('extracting');
    progressText.textContent = '提取文字: ' + _extractFileName;
    return;
  }

  if (event.phase === 'pdf_page') {
    var fileFraction = event.total > 0 ? (event.current / event.total) : 0;
    var fileStartPct = 1 + (_fileIndex - 1) * _fileShare;
    var realPct = Math.min(fileStartPct + fileFraction * _fileShare, 25);
    _barSet(realPct);
    progressText.textContent = '提取文字: ' + _extractFileName + ' (' + event.current + '/' + event.total + ' 页)';
    return;
  }

  if (event.phase === 'pdf_early_stop' || event.phase === 'pdf_done') {
    // Park the bar at this file's end of band. The sweep stays on for the
    // next file, or until the analysis phase begins in updateProgress.
    var fileEndPct = Math.min(1 + _fileIndex * _fileShare, 25);
    _barSet(fileEndPct);
    if (event.detail) {
      progressText.textContent = event.detail;
    } else {
      progressText.textContent = '文字提取完成: ' + _extractFileName + ' (' + (event.current || '?') + ' 页)';
    }
    return;
  }
}

function showWarning(event) {
  var panel = document.getElementById('warningPanel');
  if (!panel) {
    panel = document.createElement('div');
    panel.id = 'warningPanel';
    panel.className = 'warning-panel';
    var head = document.createElement('div');
    head.className = 'warning-panel-head';
    head.innerHTML = '<span>⚠ 分析提示</span>';
    var close = document.createElement('button');
    close.className = 'warning-panel-close';
    close.title = '关闭提示';
    close.innerHTML = '&times;';
    close.onclick = function() { panel.remove(); };
    head.appendChild(close);
    panel.appendChild(head);
    var resultsSection = document.getElementById('resultsSection');
    if (resultsSection) {
      resultsSection.parentNode.insertBefore(panel, resultsSection);
    }
  }
  var existing = panel.querySelectorAll('.warning-msg');
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].textContent === event.message) return;
  }
  var msg = document.createElement('div');
  msg.className = 'warning-msg';
  msg.textContent = '⚠ ' + event.message;
  panel.appendChild(msg);
}

// Marks the progress panel as failed: red bar + error text, no "分析完成".
var _analysisFailed = false;
function failProgress(msg) {
  _analysisFailed = true;
  progressBar.classList.remove('extracting');
  progressPanel.classList.add('error');
  progressText.textContent = msg || '分析失败';
  var active = progressSteps.querySelectorAll('.progress-step.active');
  active.forEach(function(el) { el.classList.remove('active'); });
}

function finishProgress() {
  if (_analysisFailed) {
    // Keep the red failure state up a bit longer than the success state.
    setTimeout(function() { progressPanel.style.display = 'none'; }, 5000);
    return;
  }
  _barSet(100);
  progressBar.classList.remove('extracting');
  progressText.textContent = '分析完成';
  var steps = progressSteps.querySelectorAll('.progress-step');
  steps.forEach(function(el) { el.classList.remove('active'); el.classList.add('done'); });
  setTimeout(function() { progressPanel.style.display = 'none'; }, 2000);
}

function addFiles(files, type) {
  const validFiles = Array.from(files).filter(f =>
    VALID_EXTS.some(ext => f.name.toLowerCase().endsWith(ext))
  );
  if (validFiles.length === 0) { alert('请选择 .docx / .doc / .pdf / .txt 格式的文件'); return; }

  const target = type === 'ref' ? selectedRefFiles : selectedFiles;
  validFiles.forEach(f => {
    if (!target.find(sf => sf.name === f.name && sf.size === f.size)) {
      target.push(f);
    }
  });
  renderAllFileLists();
  updateButtons();
}

function renderAllFileLists() {
  // Bid files with group inputs
  if (selectedFiles.length === 0) {
    fileList.innerHTML = '';
  } else {
    let html = '<div style="display:flex;flex-direction:column;gap:8px;">';
    selectedFiles.forEach((f, i) => {
      const key = f.name + '_' + f.size;
      if (!(key in fileGroups)) {
        const defaultGroup = f.name.replace(/[（(]?(商务|技术|投标|响应)[部分卷册文件]*[）)]?/g, '')
          .replace(/\.(docx|doc|pdf|txt)$/i, '').trim() || f.name;
        fileGroups[key] = defaultGroup;
      }
      const group = fileGroups[key] || '';
      html += '<div class="file-group-row">';
      html += '<span class="file-tag" style="flex:1;min-width:0;">';
      html += '<span style="word-break:break-all;">' + escapeHtml(f.name) + ' (' + formatSize(f.size) + ')</span>';
      html += '<span class="remove-btn" title="移除" onclick="removeFile(' + i + ',\'bid\')">&times;</span>';
      html += '</span>';
      // data-key goes through escapeHtml (which escapes quotes), and the
      // handler reads this.dataset.key - no inline string concatenation of
      // raw filenames into JS code.
      html += '<input class="group-input" value="' + escapeHtml(group) + '" placeholder="投标人名称" data-key="' + escapeHtml(key) + '" onchange="updateGroup(this.dataset.key, this.value)" title="相同名称的文件将合并为一个投标人"/>';
      html += '</div>';
    });
    html += '</div>';
    if (selectedFiles.length >= 2) {
      const groups = new Set();
      selectedFiles.forEach(f => { const k = f.name + '_' + f.size; groups.add(fileGroups[k] || ''); });
      html += '<p style="font-size:11px;color:var(--text-muted);margin-top:6px;">将合并为 <strong>' + groups.size + '</strong> 个投标人</p>';
    }
    fileList.innerHTML = html;
  }

  // Reference files
  if (selectedRefFiles.length === 0) {
    refFileList.innerHTML = '<span style="font-size:12px;color:#9ca3af;">（未选择 — 查重时将不扣除模板内容）</span>';
  } else {
    let refHtml = '';
    selectedRefFiles.forEach((f, i) => {
      refHtml += '<span class="file-tag" style="background:#f0fdf4;color:#16a34a;">';
      refHtml += '<span>' + escapeHtml(f.name) + ' (' + formatSize(f.size) + ')</span>';
      refHtml += '<span class="remove-btn" title="移除" onclick="removeFile(' + i + ',\'ref\')">&times;</span>';
      refHtml += '</span>';
    });
    refFileList.innerHTML = refHtml;
  }
}

function updateGroup(key, value) {
  try {
    fileGroups[key] = value.trim();
  } catch(e) {}
  renderAllFileLists();
}

function removeFile(idx, type) {
  if (isAnalyzing) return;
  const target = type === 'ref' ? selectedRefFiles : selectedFiles;
  target.splice(idx, 1);
  if (type !== 'ref') {
    // Drop group entries for files that are no longer selected.
    var validKeys = new Set(selectedFiles.map(f => f.name + '_' + f.size));
    Object.keys(fileGroups).forEach(k => { if (!validKeys.has(k)) delete fileGroups[k]; });
  }
  renderAllFileLists();
  updateButtons();
  if (selectedFiles.length === 0 && selectedRefFiles.length === 0) {
    resultsSection.style.display = 'none';
    analysisResult = null;
    clearTabCounts();
  }
  fileInput.value = '';
  refFileInput.value = '';
}

btnClear.addEventListener('click', () => {
  if (isAnalyzing) return;
  selectedFiles = [];
  selectedRefFiles = [];
  fileGroups = {};
  renderAllFileLists();
  updateButtons();
  resultsSection.style.display = 'none';
  analysisResult = null;
  clearTabCounts();
  fileInput.value = '';
  refFileInput.value = '';
  // Clear all warnings
  var warnPanel = document.getElementById('warningPanel');
  if (warnPanel) warnPanel.remove();
});

function updateButtons() {
  btnAnalyze.disabled = selectedFiles.length < 2 || isAnalyzing;
  var textEl = document.getElementById('btnAnalyzeText');
  if (!textEl) return;
  if (isAnalyzing) {
    textEl.textContent = '分析中...';
  } else if (selectedFiles.length >= 2) {
    const extra = selectedRefFiles.length > 0 ? ` + ${selectedRefFiles.length}份模板` : '';
    textEl.textContent = '开始分析' + extra;
  } else if (selectedFiles.length > 0) {
    textEl.textContent = '请至少上传 2 份标书';
  } else {
    textEl.textContent = '请至少上传 2 份标书文件';
  }
}

// ── Analyze (streaming progress with NDJSON) ──
btnAnalyze.addEventListener('click', async () => {
  if (isAnalyzing) return;
  if (selectedFiles.length < 2) {
    alert('请至少上传2份标书文件(.docx/.doc/.pdf/.txt)');
    return;
  }

  setAnalyzing(true);
  updateButtons();
  // Clear previous warnings before new analysis
  var prevWarn = document.getElementById('warningPanel');
  if (prevWarn) prevWarn.remove();
  startProgress();

  try {
    const formData = new FormData();
    selectedFiles.forEach(f => {
      formData.append('files', f);
      const key = f.name + '_' + f.size;
      formData.append('file_groups', fileGroups[key] || f.name);
    });
    selectedRefFiles.forEach(f => formData.append('ref_files', f));

    const resp = await fetch('/api/analyze_stream', {
      method: 'POST',
      body: formData
    });

    if (!resp.ok) {
      const errData = await resp.json().catch(() => ({}));
      failProgress('分析失败: ' + (errData.error || '服务器错误'));
      alert('分析失败: ' + (errData.error || '服务器错误'));
      return;
    }

    const reader = resp.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split('\n');
      buffer = lines.pop();

      for (const line of lines) {
        if (!line.trim()) continue;
        try {
          const event = JSON.parse(line);
          if (event.type === 'progress') {
            updateProgress(event);
          } else if (event.type === 'extract') {
            updateExtractProgress(event);
          } else if (event.type === 'warning') {
            // Backend sends {messages: [...]} array (plural) — show one by one
            if (event.messages && event.messages.length > 0) {
              event.messages.forEach(function(msg) {
                if (msg) showWarning({ code: event.code || 'warning', message: msg });
              });
            } else if (event.message) {
              showWarning(event);
            }
          } else if (event.type === 'result') {
            analysisResult = event.data;
            // Show warnings from result if any
            if (event.data._warnings && event.data._warnings.length > 0) {
              event.data._warnings.forEach(function(w) {
                showWarning({ code: 'no_text', message: w });
              });
            }
            resultsSection.style.display = 'block';
            renderAllTabs();
            resultsSection.scrollIntoView({ behavior: 'smooth' });
          } else if (event.type === 'error') {
            failProgress('分析失败: ' + event.message);
            alert('分析失败: ' + event.message);
          }
        } catch (e) {
          console.warn('NDJSON parse error:', e);
        }
      }
    }
  } catch (err) {
    console.error('Analysis failed:', err);
    failProgress('分析失败，请确认服务器已启动');
    alert('分析失败，请确认服务器已启动: ' + err.message);
  } finally {
    finishProgress();
    setAnalyzing(false);
    updateButtons();
  }
});

// ── Tab Switching ──
function switchToTab(tabId) {
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
  var btn = document.querySelector('[data-tab="' + tabId + '"]');
  if (btn) btn.classList.add('active');
  var content = document.getElementById(tabId);
  if (content) content.classList.add('active');
}
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => switchToTab(btn.dataset.tab));
});

// ── Render All ──
function renderAllTabs() {
  if (!analysisResult) return;
  renderVerdict();
  renderMetadata();
  renderPersonnel();
  renderSimilarity();
  renderPricing();
  updateTabCounts();
}

// ── Tab count badges ──
function _setTabCount(id, count, danger) {
  var el = document.getElementById(id);
  if (!el) return;
  if (count > 0) {
    el.textContent = count;
    el.classList.remove('hidden');
    el.classList.toggle('danger', !!danger);
  } else {
    el.classList.add('hidden');
  }
}

function updateTabCounts() {
  if (!analysisResult) return;

  // 元数据：一致项数量（排除 info 级）；存在 high 级则标红
  var meta = analysisResult.metadata || {};
  var metaMatches = meta.matches || [];
  _setTabCount('tabCountMetadata',
    metaMatches.filter(m => m.severity !== 'info').length,
    metaMatches.some(m => m.severity === 'high'));

  // 人员：高/致命交叉命中数
  var p = analysisResult.personnel || {};
  _setTabCount('tabCountPersonnel',
    (p.cross_matches || []).filter(m => m.severity === 'high' || m.severity === 'critical').length,
    true);

  // 查重：可见匹配段（高风险 + 疑似）；存在高风险则标红
  var s = analysisResult.text_similarity || {};
  var subCnt = 0, susCnt = 0;
  (s.pair_results || []).forEach(pr => {
    subCnt += pr.substantial_count || 0;
    susCnt += pr.suspicious_count || 0;
  });
  _setTabCount('tabCountSimilarity', subCnt + susCnt, subCnt > 0);

  // 报价：发现条数；存在完全一致报价则标红
  var pr = analysisResult.pricing || {};
  var samePrice = Object.values(pr.comparison || {}).some(c => c && c._same_all === true);
  _setTabCount('tabCountPricing', (pr.findings || []).length, samePrice);
}

function clearTabCounts() {
  ['tabCountMetadata', 'tabCountPersonnel', 'tabCountSimilarity', 'tabCountPricing']
    .forEach(id => { var el = document.getElementById(id); if (el) el.classList.add('hidden'); });
}

// ── Verdict Tab ──
function renderVerdict() {
  const v = analysisResult.verdict;
  const banner = document.getElementById('verdictBanner');
  const conclusion = v.conclusion;
  const level = v.conclusion_level || 'warning';
  let cls = 'warning', icon = '';

  if (level === 'high') { cls = 'suspect'; icon = '⚠️ '; }
  else if (level === 'medium') { cls = 'suspicious'; icon = '🔍 '; }
  else if (level === 'low') { cls = 'clean'; icon = '✅ '; }
  else if (level === 'uncertain') { cls = 'uncertain'; icon = '❓ '; }

  banner.className = 'verdict-banner ' + cls;
  banner.textContent = icon + '判定结论: ' + conclusion;

  // Show reference docs info
  const refDocs = analysisResult.ref_docs || [];
  if (refDocs.length > 0) {
    banner.textContent += ' (已扣除' + refDocs.length + '份模板文档)';
  }

  // ── Score overview cards ──
  const score = v.score || 0;
  const maxScore = v.max_score || 100;
  const pct = Math.min(100, Math.round(score / maxScore * 100));
  let scoreColor = '#16a34a', scoreBg = '#dcfce7';
  if (level === 'high') { scoreColor = '#dc2626'; scoreBg = '#fef2f2'; }
  else if (level === 'medium') { scoreColor = '#d97706'; scoreBg = '#fff7ed'; }
  else if (level === 'uncertain') { scoreColor = '#9ca3af'; scoreBg = '#f3f4f6'; }

  // Count clause stats
  var satisfiedCount = 0, uncertainCount = 0, notCount = 0;
  v.clauses.forEach(c => {
    if (c.satisfied === true) satisfiedCount++;
    else if (c.satisfied === false) notCount++;
    else uncertainCount++;
  });

    document.getElementById('verdictSummary').innerHTML = `
    <div class="score-overview">
      <div class="score-card">
        <div class="score-card-header">综合风险评分</div>
        <div class="score-big">
          <span class="score-num" style="color:${scoreColor};">${score}</span>
          <span class="score-unit">/ ${maxScore}</span>
        </div>
        <div class="score-bar-wrap">
          <div class="score-bar-fill" style="width:${pct}%;background:${scoreColor};"></div>
          <div class="score-bar-mark" style="left:15%;"></div>
          <div class="score-bar-mark" style="left:50%;"></div>
        </div>
        <div class="score-ticks">
          <span class="tick" style="left:0%;">0</span>
          <span class="tick tick-warn" style="left:15%;">15 可疑</span>
          <span class="tick tick-danger" style="left:50%;">50 高度嫌疑</span>
          <span class="tick" style="left:100%;">${maxScore}</span>
        </div>
      </div>
      <div class="score-card">
        <div class="score-card-header">条款命中统计</div>
        <div class="clause-stats" style="font-size:13px;">
          满足 <b class="danger">${satisfiedCount}</b> · 无法判断 <b class="warn">${uncertainCount}</b> · 不满足 <b class="success">${notCount}</b>
        </div>
        <div style="display:flex;gap:6px;flex-wrap:wrap;margin-top:12px;">
          ${v.synergy_bonus > 0 ? `<span class="bonus-badge soft">软协同 +${v.synergy_bonus}</span>` : ''}
          ${v.hard_synergy_bonus > 0 ? `<span class="bonus-badge hard">硬证据协同 +${v.hard_synergy_bonus}</span>` : ''}
          ${!(v.synergy_bonus > 0) && !(v.hard_synergy_bonus > 0) ? '<span style="font-size:12px;color:var(--text-muted);">无协同加分</span>' : ''}
        </div>
      </div>
    </div>
    <details class="rules-panel">
      <summary style="cursor:pointer;font-weight:600;color:var(--text-secondary);">📋 综合判定评分规则</summary>
      <div style="margin-top:8px;line-height:1.8;">
        <p class="rules-block-title">评分依据：《招标投标法实施条例》第四十条</p>
        <table class="rules-table">
          <thead><tr><th>条款</th><th>权重</th><th>说明</th></tr></thead>
          <tbody>
            <tr><td>第（一）项</td><td>50分</td><td>同一单位或个人编制 - 硬证据：WPS ID、授权代表=创建者、最后修改人同一</td></tr>
            <tr><td>第（二）项</td><td>25分</td><td>同一人办理投标 - 硬证据：授权代表姓名相同 / 联系电话相同 / 身份证号相同（命中即强）</td></tr>
            <tr><td>第（三）项</td><td>15分</td><td>项目管理人员相同 - 硬证据：人员高度重叠(≥50%) 或同名项目管理成员(项目经理/技术负责人等)</td></tr>
            <tr><td>第（四）项-a</td><td>5分</td><td>投标文件异常一致 - 软证据：辅助参考</td></tr>
            <tr><td>第（四）项-b</td><td>4分</td><td>报价异常一致或呈规律性差异（相同报价可达"强"） - 软证据：辅助参考</td></tr>
          </tbody>
        </table>
        <p class="rules-block-title">证据强度系数：</p>
        <ul class="rules-list">
          <li>强 = 权重 × 1.0（满分）</li>
          <li>中 = 权重 × 0.3（多项间接证据）</li>
          <li>弱 = 权重 × 0.15（单条间接证据，已过滤默认模板/通病版本号）</li>
          <li>无法判断 / 无 = 0</li>
          <li>软协同加分：第（四）项-a 和 -b 同时为"强" 时 +1分</li>
          <li>硬证据协同加分：第（一）项 与 第（二）项 同为"强" 时 +5分（文档同源+投标事宜同人双重确认）</li>
          <li>总分上限 100分</li>
        </ul>
        <p class="rules-block-title">综合结论阈值：</p>
        <ul class="rules-list">
          <li>≥ 50分 -&gt; <span style="color:#dc2626;font-weight:600;">⚠️ 存在围标串标高度嫌疑</span></li>
          <li>15–49分 -&gt; <span style="color:#d97706;font-weight:600;">🔍 存在可疑情形，建议进一步核查</span></li>
          <li>&lt; 15分 -&gt; <span style="color:#16a34a;font-weight:600;">✅ 未发现明显围标串标异常</span></li>
          <li>单份文件或全维度无法判断 -&gt; <span style="color:#9ca3af;font-weight:600;">❓ 数据不足，无法做出完整判定</span></li>
        </ul>
      </div>
    </details>
  `;

  // Clause → detail tab mapping
  var clauseTabs = {
    '第（一）项': 'tab-metadata',
    '第（二）项': 'tab-personnel',
    '第（三）项': 'tab-personnel',
    '第（四）项-a': 'tab-similarity',
    '第（四）项-b': 'tab-pricing',
  };

  let html = '';
  v.clauses.forEach(c => {
    const satisfied = c.satisfied;
    let cardCls = 'uncertain', tagCls = 'tag-uncertain', tagText = '无法判断';
    if (satisfied === true) { cardCls = 'satisfied'; tagCls = 'tag-satisfied'; tagText = '满足'; }
    else if (satisfied === false) { cardCls = 'not-satisfied'; tagCls = 'tag-not'; tagText = '不满足'; }

    // Clause number from string: "第（一）项" → "一"
    var clauseNum = c.clause.replace('第（', '').replace('）项', '').replace('项-', '').replace('项', '');
    var targetTab = clauseTabs[c.clause] || '';
    html += `<div class="clause-card ${cardCls}"${targetTab ? ` onclick="switchToTab('${targetTab}')" style="cursor:pointer;"` : ''}>
      <div class="clause-header">
        <span class="clause-index">${escapeHtml(clauseNum)}</span>
        <strong style="font-size:14px;">${escapeHtml(c.description)}</strong>
        <span class="clause-tag ${tagCls}">${tagText}</span>
        ${c._score !== undefined ? `<span class="clause-score-badge">+${c._score}分</span>` : ''}
        ${targetTab ? `<span style="font-size:10px;color:var(--accent-dim);margin-left:4px;">详情 →</span>` : ''}
      </div>`;

    if (c.evidence && c.evidence.length > 0) {
      html += '<ul class="clause-evidence">';
      c.evidence.forEach(e => { html += `<li>${escapeHtml(String(e))}</li>`; });
      html += '</ul>';
    }
    if (c.evidence_level) {
      const lvlCls = c.evidence_level === '强' ? 'level-strong'
        : c.evidence_level === '中' ? 'level-medium'
        : c.evidence_level === '弱' ? 'level-weak'
        : 'level-none';
      html += `<span class="evidence-level ${lvlCls}">证据强度: ${c.evidence_level}</span>`;
    }
    html += '</div>';
  });
  document.getElementById('clausesGrid').innerHTML = html;
}

// ── Metadata Tab ──
function renderMetadata() {
  const meta = analysisResult.metadata;

  // ── 1. 元数据一致项 (prominent, first) ──
  let mhtml = '';
  if (meta.matches.length > 0) {
    mhtml += '<div class="section-title">🔍 元数据一致项</div>';
    mhtml += '<div class="meta-match-grid">';
    meta.matches.forEach(m => {
      const isInfo = m.severity === 'info';
      const icon = isInfo ? 'ℹ️' : (m.severity === 'high' ? '⚠️' : '📋');
      const cls = isInfo ? 'meta-info' : (m.severity === 'high' ? 'meta-critical' : 'meta-normal');
      const badgeCls = isInfo ? 'badge-info' : (m.severity === 'high' ? 'badge-high' : 'badge-medium');
      mhtml += `<div class="meta-match-card ${cls}">
        <div class="meta-match-icon">${icon}</div>
        <div class="meta-match-body">
          <div class="meta-match-field">${escapeHtml(m.field)}</div>
          <div class="meta-match-value">${escapeHtml(String(m.value).substring(0, 200))}</div>
          ${m.pair ? `<div class="meta-match-pair">📄 ${escapeHtml(m.pair)}</div>` : ''}
        </div>
        <span class="meta-match-badge ${badgeCls}">${m.verdict}</span>
      </div>`;
    });
    mhtml += '</div>';
  } else {
    mhtml += '<div class="meta-empty"><span>✅</span> 未发现元数据一致项</div>';
  }

  // Time findings
  meta.findings.forEach(f => {
    if (f && f.trim() && !f.includes('一致')) mhtml += `<p style="font-size:13px;color:#64748b;padding:4px 0;">📅 ${escapeHtml(f)}</p>`;
  });
  document.getElementById('metadataMatches').innerHTML = mhtml;

  // ── 2. 文件元数据详情 (collapsed by default) ──
  let fhtml = '<div class="section-title" style="margin-top:20px;cursor:pointer;" onclick="document.getElementById(\'metaDetailBody\').style.display = document.getElementById(\'metaDetailBody\').style.display === \'none\' ? \'block\' : \'none\'">📋 文件元数据详情 ▾</div>';
  fhtml += '<div id="metaDetailBody" style="display:none;">';
  meta.files.forEach(f => {
    const shortName = f.name.length > 40 ? f.name.substring(0, 40) + '...' : f.name;
    fhtml += `<h3 class="section-subtitle tinted"><span class="file-chip">${escapeHtml(shortName)}</span></h3>`;
    fhtml += '<table class="data-table"><thead><tr><th>属性</th><th>值</th></tr></thead><tbody>';
    const rows = [
      ['文件类型', f._error ? '读取异常' : (f.name.endsWith('.pdf') ? 'PDF' : f.name.endsWith('.doc') ? 'DOC(旧版)' : f.name.endsWith('.txt') ? 'TXT(纯文本)' : 'DOCX')],
      ['创建者', f.creator], ['最后保存者', f.last_modified_by],
      ['创建时间', f.created], ['修改时间', f.modified],
      ['修订次数', f.revision], ['编辑时长(分钟)', f.total_edit_time],
      ['页数', f.pages], ['字数', f.words],
      ['应用程序', f.application], ['模板', f.template],
      ['WPS版本', f.KSOProductBuildVer], ['WPS保存记录', f.KSOTemplateDocerSaveRecord],
      ['ICV', f.ICV],
    ];
    rows.forEach(([label, val]) => {
      if (val) fhtml += `<tr><td>${label}</td><td style="word-break:break-all;">${escapeHtml(String(val).substring(0, 300))}</td></tr>`;
    });
    fhtml += '</tbody></table>';
  });

  const refDocs = analysisResult.ref_docs || [];
  if (refDocs.length > 0) {
    fhtml += `<p style="margin-top:8px;font-size:13px;color:#64748b;">📂 模板参考: ${refDocs.map(r => escapeHtml(r)).join(', ')}</p>`;
  }
  fhtml += '</div>';
  document.getElementById('metadataTables').innerHTML = fhtml;
}

// ── Personnel Tab ──
function renderPersonnel() {
  const p = analysisResult.personnel;
  if (!p) return;

  // Role display labels
  const ROLE_LABELS = {
    'legal_rep': '法定代表人', 'authorized_rep': '授权代表',
    'project_manager': '项目经理', 'tech_lead': '技术负责人',
    'bid_contact': '投标联系人', 'team_member': '团队成员',
    'signatory': '签署人', 'other': '其他人员'
  };

  let html = '';
  let hasAnyData = false;

  p.files.forEach(f => {
    const rows = [
      ['公司名称', f.company_name],
      ['法定代表人', f.legal_rep], ['授权代表', f.authorized_rep],
      ['身份证号', f.id_number], ['联系电话', f.phone],
      ['联系地址', f.address], ['响应日期', f.response_date],
    ];
    const filled = rows.filter(r => r[1]);

    // Check for project team members
    const allPersons = f.all_persons || [];
    const teamMembers = allPersons.filter(function(p) {
      return ['project_manager', 'tech_lead', 'team_member', 'bid_contact'].indexOf(p.role) >= 0;
    });

    if (filled.length === 0 && teamMembers.length === 0) return;

    hasAnyData = true;
    html += '<h3 class="section-subtitle">' + escapeHtml(f.name) + '</h3>';

    // Basic info table
    if (filled.length > 0) {
      html += '<table class="data-table"><thead><tr><th>属性</th><th>值</th></tr></thead><tbody>';
      filled.forEach(function(row) {
        html += '<tr><td>' + row[0] + '</td><td>' + escapeHtml(String(row[1])) + '</td></tr>';
      });
      html += '</tbody></table>';
    }

    // Project team members table
    if (teamMembers.length > 0) {
      html += '<h4 class="section-subtitle">项目团队成员</h4>';
      html += '<table class="data-table"><thead><tr><th>姓名</th><th>角色</th><th>置信度</th></tr></thead><tbody>';
      teamMembers.forEach(function(m) {
        var roleLabel = ROLE_LABELS[m.role] || m.role;
        var conf = m.confidence ? Math.round(m.confidence * 100) + '%' : '-';
        html += '<tr><td>' + escapeHtml(m.name) + '</td><td>' + roleLabel + '</td><td>' + conf + '</td></tr>';
      });
      html += '</tbody></table>';
    }
  });

  if (!hasAnyData) {
    html = '<p class="empty-note">未从标书中提取到人员信息（法定代表人、授权代表、项目成员等）</p>';
  }
  document.getElementById('personnelTables').innerHTML = html;

  // Cross-match findings
  var mhtml = '';
  if (p.cross_matches && p.cross_matches.length > 0) {
    p.cross_matches.forEach(function(m) {
      var sevClass = 'badge-medium';
      var sevLabel = '一般';
      if (m.severity === 'high') { sevClass = 'badge-high'; sevLabel = '严重'; }
      if (m.severity === 'critical') { sevClass = 'badge-critical'; sevLabel = '致命'; }
      if (m.severity === 'info') { sevClass = 'badge-info'; sevLabel = '信息'; }
      mhtml += '<div class="match-card">';
      mhtml += '<span class="match-badge ' + sevClass + '">' + sevLabel + '</span>';
      mhtml += '<strong>' + escapeHtml(m.type) + '</strong>';
      mhtml += '<p class="match-detail">' + escapeHtml(m.detail) + '</p>';
      mhtml += '</div>';
    });
  }

  if (p.findings && p.findings.length > 0) {
    if (mhtml) mhtml += '<div style="margin-top:12px;"></div>';
    mhtml += '<ul class="finding-list">';
    p.findings.forEach(function(f) { mhtml += '<li>' + escapeHtml(f) + '</li>'; });
    mhtml += '</ul>';
  }

  if (!mhtml && !hasAnyData) {
    document.getElementById('personnelMatches').innerHTML = '';
  } else if (!mhtml) {
    document.getElementById('personnelMatches').innerHTML = '<p class="empty-note">未发现人员交叉异常</p>';
  } else {
    document.getElementById('personnelMatches').innerHTML = mhtml;
  }
}

// ── Similarity Tab ──
let _allMatchRefs = []; // flat list of all matches for modal navigation
let _currentSimFilter = 'all'; // chip-selector filter state

function renderSimilarity() {
  if (!analysisResult) return;
  const s = analysisResult.text_similarity;
  const filter = _currentSimFilter;

  let totalMatches = 0, totalSubstantial = 0, totalSuspicious = 0, totalTemplate = s.template_matches || 0;
  s.pair_results.forEach(p => {
    totalMatches += p.total_matches;
    totalSubstantial += (p.substantial_count || 0);
    totalSuspicious += (p.suspicious_count || 0);
  });

  // Chip-selector cards (click = set filter + re-render)
  const chip = (value, num, cls, label) =>
    `<div class="stat-card${filter===value?' active':''}" onclick="_currentSimFilter='${value}';renderSimilarity();">
      <div class="stat-num${cls?' '+cls:''}">${num}</div>
      <div class="stat-label">${label}</div>
    </div>`;

  document.getElementById('similaritySummary').innerHTML = `
    <div class="smart-toggle-row">
      <div class="summary-stat" style="margin:0;">
        ${chip('all', totalMatches, '', '总匹配段落数')}
        ${chip('substantial', totalSubstantial, 'danger', '🔴 可能高风险异常')}
        ${chip('suspicious', totalSuspicious, 'warn', '🟡 疑似模板')}
        ${chip('template', totalTemplate, 'success', '⚪ 已过滤模板')}
      </div>
      <button class="smart-toggle-btn" onclick="toggleAllPairs()" id="btnSmartToggle">▸ 展开全部</button>
    </div>
    <ul class="finding-list">${s.findings.map(f => `<li>${escapeHtml(f)}</li>`).join('')}</ul>
        <details class="rules-panel" style="margin-top:10px;">
      <summary style="cursor:pointer;font-weight:600;color:var(--text-secondary);">📋 风险分级评分规则</summary>
      <div style="margin-top:8px;line-height:1.8;">
        <p class="rules-block-title">过滤链（依次执行）：</p>
        <ol class="rules-list">
          <li>参照文件匹配 - 用户上传的招标文件中出现过的段落 -&gt; <span style="color:#16a34a;">⚪ 已过滤</span></li>
          <li>规则库匹配 - 签字/盖章/日期、公告措辞、法律条款、格式声明、编号等 -&gt; <span style="color:#16a34a;">⚪ 已过滤</span></li>
          <li>全局共现检测 - 同一段落在 ≥ max(3, 50%文件数) 份标书中出现 -&gt; <span style="color:#16a34a;">⚪ 已过滤</span></li>
        </ol>
        <p class="rules-block-title">实质性评分（4维度加权，满分1.0）：</p>
        <table class="rules-table">
          <thead><tr><th>维度</th><th>权重</th><th>说明</th></tr></thead>
          <tbody>
            <tr><td>段落长度</td><td>30%</td><td>越长越可能为独立编制内容（200字满分）</td></tr>
            <tr><td>技术/业务术语密度</td><td>30%</td><td>含型号、参数、专业术语等具体信息</td></tr>
            <tr><td>数值/编号特异性</td><td>25%</td><td>含具体金额、百分比、日期、版本号等</td></tr>
            <tr><td>句式模板化程度（反向）</td><td>15%</td><td>"应当/必须/不得"密度高且无具体信息则扣分</td></tr>
          </tbody>
        </table>
        <p class="rules-block-title">评分阈值：</p>
        <ul class="rules-list">
          <li>≥ 0.6 -&gt; <span style="color:#dc2626;font-weight:600;">🔴 可能高风险异常</span> - 参与围串标结论判定</li>
          <li>0.3–0.6 -&gt; <span style="color:#d97706;font-weight:600;">🟡 疑似模板</span> - 展示但降级，不参与判定</li>
          <li>&lt; 0.3 -&gt; <span style="color:#16a34a;font-weight:600;">⚪ 已过滤模板</span> - 自动扣除</li>
        </ul>
      </div>
    </details>
  `;

  // Build flat list for modal nav
  _allMatchRefs = [];

  // ── Pair overview cards ──
  let overview = '<div class="pair-overview-grid">';
  s.pair_results.forEach((pr, pi) => {
    const subCnt = pr.substantial_count || 0;
    const susCnt = pr.suspicious_count || 0;
    const tplCnt = pr.template_count || 0;
    overview += `<div class="pair-overview-card" onclick="scrollToPair(${pi})">
      <div class="pair-overview-header">对比 ${pi + 1}</div>
      <div style="font-size:12px;color:#666;margin:4px 0;">${shortenName(pr.file1)} ↔ ${shortenName(pr.file2)}</div>
      <div style="display:flex;gap:8px;font-size:12px;">
        ${subCnt > 0 ? `<span style="color:#dc2626;font-weight:600;">${subCnt}可能高风险</span>` : ''}
        ${susCnt > 0 ? `<span style="color:#d97706;font-weight:600;">${susCnt}疑似</span>` : ''}
        <span style="color:#16a34a;">${tplCnt}模板</span>
        <span style="color:#888;">${pr.total_matches}总计</span>
      </div>
    </div>`;
  });
  overview += '</div>';
  document.getElementById('similaritySummary').innerHTML += `
    <div id="pairOverview">${overview}</div>`;

  const PAGE_SIZE = 10;

  // ── Per-pair detail sections ──
  let dhtml = '';
  s.pair_results.forEach((pr, pairIdx) => {
    const pairId = `pair-${pairIdx}`;

    const substantialMatches = pr.matches.filter(m => m.risk_level === 'substantial');
    const suspiciousMatches = pr.matches.filter(m => m.risk_level === 'suspicious');
    const templateMatches = pr.matches.filter(m => m.risk_level === 'template');

    // Backward compat: if risk_level not present, fall back to abnormal flag
    if (substantialMatches.length === 0 && suspiciousMatches.length === 0 && templateMatches.length === 0) {
      pr.matches.forEach(m => {
        if (m.abnormal) {
          substantialMatches.push(m);
        } else {
          templateMatches.push(m);
        }
      });
    }

    const hasSubstantial = substantialMatches.length > 0;
    const hasSuspicious = suspiciousMatches.length > 0;
    const hasTemplate = templateMatches.length > 0;

    const showSubstantial = filter === 'all' || filter === 'substantial' || filter === 'abnormal';
    const showSuspicious = filter === 'all' || filter === 'suspicious';
    const showTemplate = filter === 'all' || filter === 'template';

    const subCnt = pr.substantial_count || 0;
    const susCnt = pr.suspicious_count || 0;
    const tplCnt = pr.template_count || 0;

    // Auto-expand: open pair if there's any visible content
    const hasVisibleContent = (showSubstantial && hasSubstantial) ||
                              (showSuspicious && hasSuspicious) ||
                              (showTemplate && hasTemplate);
    const openByDefault = hasVisibleContent ? 'block' : 'none';
    const toggleIcon = hasVisibleContent ? '▼' : '▶';

    dhtml += `<div class="pair-section" id="${pairId}">
      <div class="pair-header" onclick="togglePair('${pairId}')">
        <span class="pair-toggle" id="${pairId}-toggle">${toggleIcon}</span>
        <span class="pair-title">对比 ${pairIdx + 1}: ${shortenName(pr.file1, 15)} ↔ ${shortenName(pr.file2, 15)}</span>
        <span class="pair-stats">
          ${subCnt > 0 ? `<span style="color:#dc2626;">${subCnt}可能高风险</span>` : ''}
          ${susCnt > 0 ? `<span style="color:#d97706;margin-left:8px;">${susCnt}疑似</span>` : ''}
          <span style="color:#16a34a;margin-left:8px;">${tplCnt}模板</span>
          <span style="color:#888;margin-left:8px;">${pr.total_matches}总计</span>
        </span>
      </div>
      <div class="pair-body" id="${pairId}-body" style="display:${openByDefault};">`;

    // Render helper
    const renderPaginated = (matches, label, colorClass, bgStyle, wrapInDetails) => {
      if (matches.length === 0) return '';
      const bodyHtml = (matchList) => {
        let h = '';
        const visible = matchList.slice(0, PAGE_SIZE);
        const hidden = matchList.slice(PAGE_SIZE);
        visible.forEach(m => {
          const refIdx = _allMatchRefs.length;
          const rl = m.risk_level || (m.abnormal ? 'abnormal' : 'template');
          _allMatchRefs.push({ match: m, file1: pr.file1, file2: pr.file2, type: rl });
          h += `<div class="text-match-item" style="border-left:3px solid ${colorClass};">
            <div class="text-match-header">
              <span class="text-match-num" style="${bgStyle}">#${m.index}</span>
              <span class="text-match-length">${m.length}字</span>
              ${(m.reasons||[]).map(r => `<span class="text-match-reason">${escapeHtml(r)}</span>`).join('')}
              ${m.score !== undefined ? `<span class="text-match-score" style="font-size:11px;color:#888;">[评分:${m.score}]</span>` : ''}
              <button class="match-locate-btn" onclick="openMatchModal(${refIdx})">📍 定位</button>
            </div>
            <div class="text-match-content">${escapeHtml(m.text.substring(0, 200))}${m.text.length > 200 ? '...' : ''}</div>
          </div>`;
        });
        if (hidden.length > 0) {
          const labelId = label.replace(/[^a-z0-9一-鿿]/g,'');
          h += `<div id="${pairId}-more-${labelId}" style="display:none;">`;
          hidden.forEach(m => {
            const refIdx = _allMatchRefs.length;
            const rl = m.risk_level || (m.abnormal ? 'abnormal' : 'template');
            _allMatchRefs.push({ match: m, file1: pr.file1, file2: pr.file2, type: rl });
            h += `<div class="text-match-item" style="border-left:3px solid ${colorClass};">
              <div class="text-match-header">
                <span class="text-match-num" style="${bgStyle}">#${m.index}</span>
                <span class="text-match-length">${m.length}字</span>
                ${m.score !== undefined ? `<span class="text-match-score" style="font-size:11px;color:#888;">[评分:${m.score}]</span>` : ''}
                <button class="match-locate-btn" onclick="openMatchModal(${refIdx})">📍 定位</button>
              </div>
              <div class="text-match-content">${escapeHtml(m.text.substring(0, 200))}${m.text.length > 200 ? '...' : ''}</div>
            </div>`;
          });
          h += '</div>';
          h += `<button class="btn btn-sm btn-outline" style="margin-top:4px;"
            onclick="toggleMore('${pairId}-more-${labelId}', this)">显示全部 ${hidden.length} 项</button>`;
        }
        return h;
      };

      if (wrapInDetails) {
        return `<details style="margin-top:6px;" open>
          <summary style="font-weight:600;color:${colorClass};cursor:pointer;padding:4px 0;">${label} (${matches.length}处)</summary>
          <div style="margin-top:4px;">${bodyHtml(matches)}</div>
        </details>`;
      }
      return `<p style="font-weight:600;color:${colorClass};margin:8px 0 4px;">${label} (${matches.length}处):</p>` + bodyHtml(matches);
    };

    if (showSubstantial) {
      dhtml += renderPaginated(substantialMatches, '🔴 可能高风险异常段落（高风险）', '#dc2626', '', false);
    }
    if (showSuspicious) {
      dhtml += renderPaginated(suspiciousMatches, '🟡 疑似模板段落（已降级）', '#d97706', 'background:#fffbeb;color:#92400e;', false);
    }
    if (showTemplate) {
      // Wrap template in collapsible <details> when filter is 'all', normal otherwise
      const wrapTpl = filter === 'all';
      dhtml += renderPaginated(templateMatches, '⚪ 已过滤模板内容', '#16a34a', 'background:#f0fdf4;color:#16a34a;', wrapTpl);
    }
    if (filter === 'substantial' && !hasSubstantial) dhtml += '<p style="color:#888;padding:8px 0;">无可能高风险异常段落</p>';
    if (filter === 'suspicious' && !hasSuspicious) dhtml += '<p style="color:#888;padding:8px 0;">无疑似模板段落</p>';
    if (filter === 'template' && !hasTemplate) dhtml += '<p style="color:#888;padding:8px 0;">无已过滤模板段落</p>';

    dhtml += '</div></div>';
  });
  document.getElementById('similarityDetails').innerHTML = dhtml;

  // Update smart toggle button text based on current state
  updateSmartToggleBtn();
}

// ── Pricing Tab ──
function renderPricing() {
  if (!analysisResult) return;
  const p = analysisResult.pricing;
  if (!p) return;

  // ── Total Price Comparison ──
  let tableHtml = '';
  if (p.files && p.files.length > 0) {
    tableHtml += '<div class="section-title">💰 总价对比</div>';
    tableHtml += '<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>报价项</th>';
    p.files.forEach(f => {
      const s = (f.name || '').length > 20 ? (f.name || '').substring(0, 20) + '...' : (f.name || '');
      tableHtml += '<th>' + escapeHtml(s) + '</th>';
    });
    tableHtml += '</tr></thead><tbody>';

    var priceRows = [
      { label: '含税总价（元）', key: 'totalPriceInTax' },
      { label: '不含税总价（元）', key: 'totalPrice' },
      { label: '税率', key: 'taxRate' },
      { label: '收入（元）', key: 'revenue' },
      { label: '成本（元）', key: 'cost' },
    ];
    priceRows.forEach(function(row) {
      if (p.files.some(function(f) { return f[row.key] != null; })) {
        tableHtml += '<tr><td>' + row.label + '</td>';
        p.files.forEach(function(f) {
          var v = f[row.key];
          tableHtml += '<td>' + (v != null ? (typeof v === 'number' ? v.toLocaleString() : escapeHtml(String(v))) : '<span style="color:#9ca3af;">—</span>') + '</td>';
        });
        tableHtml += '</tr>';
      }
    });
    tableHtml += '</tbody></table></div>';

    if (p.comparison && Object.keys(p.comparison).length > 0) {
      tableHtml += '<div style="margin-top:12px;font-size:13px;">';
      Object.entries(p.comparison).forEach(function(entry) {
        var label = entry[0];
        var data = entry[1];
        var sameAll = data._same_all;
        var vals = Object.entries(data).filter(function(kv) { return kv[0] !== '_same_all'; }).map(function(kv) { return kv[1]; }).filter(function(v) { return v != null; });
        if (sameAll === true && vals.length >= 2) {
          tableHtml += '<p class="price-compare-line same">📌 <strong>' + escapeHtml(label) + '</strong>: 全部一致 (' + vals[0].toLocaleString() + ' 元)</p>';
        } else if (sameAll === false && vals.length >= 2) {
          tableHtml += '<p class="price-compare-line diff">⚠️ <strong>' + escapeHtml(label) + '</strong>: 存在差异 — ' + vals.map(function(v) { return v.toLocaleString(); }).join(' / ') + ' 元</p>';
        }
      });
      tableHtml += '</div>';
    }
  } else {
    tableHtml += '<p class="empty-note">暂无报价数据</p>';
  }
  document.getElementById('pricingTable').innerHTML = tableHtml;

  // ── Sub-item Comparison ──
  var subHtml = '';
  if (p.subItemCompare && p.subItemCompare.length > 0) {
    subHtml += '<div class="section-title" style="margin-top:20px;">📊 分项明细比对</div>';
    p.subItemCompare.forEach(function(c, ci) {
      subHtml += '<div class="subitem-card">';
      subHtml += '<strong style="font-size:14px;">' + escapeHtml(c.name || ('分项 ' + (ci + 1))) + '</strong>';
      if (c.items && c.items.length > 1) {
        // Check if any item has manufacturer/model extras
        var hasExtras = c.items.some(function(it) { return it.extras && it.extras['厂家/型号']; });
        var hasType = c.items.some(function(it) { return it.type; });

        var headerCols = '<th>投标人</th>';
        if (hasType) headerCols += '<th>来源</th>';
        if (hasExtras) headerCols += '<th>厂家/型号</th>';
        headerCols += '<th>单价</th><th>数量</th><th>不含税总价</th><th>含税总价</th><th>税率</th>';

        subHtml += '<table class="data-table" style="margin-top:8px;"><thead><tr>' + headerCols + '</tr></thead><tbody>';
        c.items.forEach(function(it) {
          var sf = (it.file || '').length > 25 ? (it.file || '').substring(0, 25) + '...' : (it.file || '');
          subHtml += '<tr>';
          subHtml += '<td>' + escapeHtml(sf) + '</td>';
          if (hasType) subHtml += '<td>' + escapeHtml(it.type || '—') + '</td>';
          if (hasExtras) subHtml += '<td>' + escapeHtml((it.extras && it.extras['厂家/型号']) || '—') + '</td>';
          subHtml += '<td>' + (it.unitPrice != null ? it.unitPrice.toLocaleString() : '—') + '</td>';
          subHtml += '<td>' + (it.count != null ? it.count : '—') + '</td>';
          subHtml += '<td>' + (it.totalPrice != null ? it.totalPrice.toLocaleString() : '—') + '</td>';
          subHtml += '<td>' + (it.totalPriceInTax != null ? it.totalPriceInTax.toLocaleString() : '—') + '</td>';
          subHtml += '<td>' + (it.tax != null ? escapeHtml(String(it.tax)) : '—') + '</td>';
          subHtml += '</tr>';
        });
        subHtml += '</tbody></table>';
      }
      if (c.findings && c.findings.length > 0) {
        subHtml += '<ul class="finding-list" style="margin-top:8px;">';
        c.findings.forEach(function(f) { subHtml += '<li>' + escapeHtml(f) + '</li>'; });
        subHtml += '</ul>';
      }
      subHtml += '</div>';
    });
  }
  document.getElementById('subItemTable').innerHTML = subHtml;

  // ── Findings ──
  var fhtml = '';
  if (p.findings && p.findings.length > 0) {
    fhtml += '<div class="section-title" style="margin-top:16px;">📝 分析发现</div>';
    fhtml += '<ul class="finding-list">';
    p.findings.forEach(function(f) { fhtml += '<li>' + escapeHtml(f) + '</li>'; });
    fhtml += '</ul>';
  }
  document.getElementById('pricingFindings').innerHTML = fhtml;
}

function toggleAllPairs() {
  // Smart toggle: if all are collapsed, expand all; otherwise collapse all
  const bodies = document.querySelectorAll('.pair-body');
  const allExpanded = Array.from(bodies).every(el => el.style.display === 'block');
  const target = allExpanded ? 'none' : 'block';
  const icon = allExpanded ? '▶' : '▼';
  bodies.forEach(el => el.style.display = target);
  document.querySelectorAll('.pair-toggle').forEach(el => el.textContent = icon);
  updateSmartToggleBtn();
}

function updateSmartToggleBtn() {
  const btn = document.getElementById('btnSmartToggle');
  if (!btn) return;
  const bodies = document.querySelectorAll('.pair-body');
  if (bodies.length === 0) { btn.textContent = '▸ 展开全部'; return; }
  const allExpanded = Array.from(bodies).every(el => el.style.display === 'block');
  const allCollapsed = Array.from(bodies).every(el => el.style.display === 'none');
  if (allExpanded) {
    btn.textContent = '▴ 折叠全部';
  } else {
    btn.textContent = '▸ 展开全部';
  }
}
function toggleMore(id, btn) {
  const el = document.getElementById(id);
  if (!el) return;
  if (el.style.display === 'none') {
    el.style.display = 'block';
    btn.textContent = '收起';
  } else {
    el.style.display = 'none';
    const total = el.querySelectorAll('.text-match-item').length;
    btn.textContent = `显示全部 ${total} 项`;
  }
}

function shortenName(name, max) {
  max = max || 18;
  if (!name) return '';
  return name.length > max ? name.substring(0, max) + '...' : name;
}

function scrollToPair(idx) {
  const el = document.getElementById(`pair-${idx}`);
  if (el) {
    el.scrollIntoView({ behavior: 'smooth', block: 'start' });
    // Expand
    const body = document.getElementById(`pair-${idx}-body`);
    const toggle = document.getElementById(`pair-${idx}-toggle`);
    if (body && body.style.display === 'none') {
      body.style.display = 'block';
      if (toggle) toggle.textContent = '▼';
    }
    updateSmartToggleBtn();
  }
}

function togglePair(pairId) {
  const body = document.getElementById(pairId + '-body');
  const toggle = document.getElementById(pairId + '-toggle');
  if (!body) return;
  if (body.style.display === 'none') {
    body.style.display = 'block';
    if (toggle) toggle.textContent = '▼';
  } else {
    body.style.display = 'none';
    if (toggle) toggle.textContent = '▶';
  }
  updateSmartToggleBtn();
}

// ── Match Modal ──
let _currentMatchIdx = -1;

function openMatchModal(idx) {
  _currentMatchIdx = idx;
  renderModalMatch();
  document.getElementById('matchModal').style.display = 'flex';
}

document.getElementById('btnCloseModal').addEventListener('click', () => {
  document.getElementById('matchModal').style.display = 'none';
});
document.getElementById('matchModal').addEventListener('click', e => {
  if (e.target === document.getElementById('matchModal')) {
    document.getElementById('matchModal').style.display = 'none';
  }
});

document.getElementById('btnPrevMatch').addEventListener('click', () => {
  if (_allMatchRefs.length === 0) return;
  _currentMatchIdx = (_currentMatchIdx - 1 + _allMatchRefs.length) % _allMatchRefs.length;
  renderModalMatch();
});
document.getElementById('btnNextMatch').addEventListener('click', () => {
  if (_allMatchRefs.length === 0) return;
  _currentMatchIdx = (_currentMatchIdx + 1) % _allMatchRefs.length;
  renderModalMatch();
});

// Keyboard nav
document.addEventListener('keydown', e => {
  // Escape closes whichever modal is open (match modal takes priority)
  if (e.key === 'Escape') {
    if (document.getElementById('matchModal').style.display === 'flex') {
      document.getElementById('btnCloseModal').click();
      return;
    }
    if (document.getElementById('historyModal').style.display === 'flex') {
      document.getElementById('btnCloseHistory').click();
      return;
    }
    return;
  }
  if (document.getElementById('matchModal').style.display !== 'flex') return;
  if (e.key === 'ArrowLeft') document.getElementById('btnPrevMatch').click();
  if (e.key === 'ArrowRight') document.getElementById('btnNextMatch').click();
});

function renderModalMatch() {
  const ref = _allMatchRefs[_currentMatchIdx];
  if (!ref || !ref.match) return;

  const m = ref.match;
  const matchText = m.text || '';
  document.getElementById('modalTitle').textContent =
    `第${m.index || '?'}项匹配 (${m.length || 0}字) — ${ref.type === 'abnormal' ? '⚠ 异常一致' : '✅ 模板匹配'}`;
  document.getElementById('matchCounter').textContent =
    `${_currentMatchIdx + 1} / ${_allMatchRefs.length}`;
  document.getElementById('diffLabel1').textContent = ref.file1 || '';
  document.getElementById('diffLabel2').textContent = ref.file2 || '';
  document.getElementById('matchReason').innerHTML = (m.reasons || []).map(r => `<span class="text-match-reason">${escapeHtml(r)}</span>`).join(' ');
  document.getElementById('diffContent1').innerHTML = renderContextWithHighlight(m.ctx1 || matchText, matchText, m.length);
  document.getElementById('diffContent2').innerHTML = renderContextWithHighlight(m.ctx2 || matchText, matchText, m.length);
}

// Server-side matching (see _build_normalized_map / _FOLD_TABLE in app.py)
// drops whitespace + invisible chars and folds case / full-width / CJK
// punctuation, so two extractions that differ only in those forms still
// match. We mirror that normalization here to locate the highlight span in
// EITHER pane's raw context, then map back to raw indices so the displayed
// text stays faithful to the source document.
var _SKIP_RE = /[\s\u0085\u001c\u001d\u001e\u001f­​‌‍﻿]/;
var _FOLD = {};
(function () {
  for (var o = 0xFF01; o < 0xFF5F; o++) {            // full-width !-~ -> ascii
    var d = String.fromCharCode(o - 0xFEE0);
    if (d >= 'A' && d <= 'Z') d = d.toLowerCase();   // fold letters to lowercase
    _FOLD[o] = d;
  }
  for (var c = 65; c <= 90; c++) { _FOLD[c] = String.fromCharCode(c + 32); }
  var punct = {
    '，': ',', '。': '.', '、': ',', '；': ';', '：': ':', '？': '?', '！': '!',
    '‘': "'", '’': "'", '‚': "'", '“': '"', '”': '"', '„': '"',
    '（': '(', '）': ')', '【': '[', '】': ']', '〔': '[', '〕': ']',
    '《': '<', '》': '>', '〈': '<', '〉': '>', '『': '[', '』': ']',
    '–': '-', '—': '-', '―': '-', '−': '-', '…': '.', '·': '.',
    '～': '~', '％': '%', '＋': '+', '×': 'x', '÷': '/', '￥': '¥'
  };
  for (var k in punct) { _FOLD[k.charCodeAt(0)] = punct[k]; }
})();
function _foldChar(c) { return _FOLD[c.charCodeAt(0)] || c; }
function _buildNormMap(text) {
  var chars = [], pos = [];
  for (var i = 0; i < text.length; i++) {
    var c = text.charAt(i);
    if (_SKIP_RE.test(c)) continue;
    chars.push(_foldChar(c)); pos.push(i);
  }
  return { s: chars.join(''), p: pos };
}
function renderContextWithHighlight(ctx, matchText, fullLen) {
  if (!ctx) return escapeHtml(matchText || '');
  if (!matchText) return escapeHtml(ctx);
  var cm = _buildNormMap(ctx);
  var needle = _buildNormMap(matchText).s;
  if (needle.length === 0) return escapeHtml(ctx);
  var want = (typeof fullLen === 'number' && fullLen > needle.length) ? fullLen : needle.length;
  var start = cm.s.indexOf(needle);
  if (start < 0) {
    // Fallback: the tail may have drifted between extractions; match the head.
    var head = needle.slice(0, Math.min(16, needle.length));
    start = cm.s.indexOf(head);
  }
  if (start < 0) {
    // Fallback 2 (near-duplicate matches): the two extractions differ by
    // scattered edits, so no single run of the needle exists in this pane.
    // Anchor the highlight on the longest exact substring shared between the
    // needle and the normalized ctx (rolling DP), then extend to `want`
    // chars around the anchor so the whole match region lights up.
    var bestStart = -1, bestLen = 0, nl = needle.length, cl = cm.s.length;
    var dp = new Array(cl + 1).fill(0);
    for (var i2 = 1; i2 <= nl; i2++) {
      var prev = 0;
      for (var j2 = 1; j2 <= cl; j2++) {
        var tmp = dp[j2];
        if (needle.charAt(i2 - 1) === cm.s.charAt(j2 - 1)) {
          dp[j2] = prev + 1;
          if (dp[j2] > bestLen) { bestLen = dp[j2]; bestStart = j2 - dp[j2]; }
        } else { dp[j2] = 0; }
        prev = tmp;
      }
    }
    if (bestLen >= 8 && bestStart >= 0) {
      start = bestStart;
      // Extend left/right around the anchor to cover `want` chars (clamped).
      var left = Math.min(start, Math.floor((want - bestLen) / 2));
      start = start - left;
    }
  }
  if (start < 0) return escapeHtml(ctx);
  var end = Math.min(start + want, cm.s.length);
  if (end <= start) end = start + 1;
  var rawStart = cm.p[start];
  var rawEnd = cm.p[end - 1] + 1;
  return escapeHtml(ctx.substring(0, rawStart)) +
         '<mark class="match-highlight">' + escapeHtml(ctx.substring(rawStart, rawEnd)) + '</mark>' +
         escapeHtml(ctx.substring(rawEnd));
}

// ── History ──
async function openHistory() {
  document.getElementById('historyModal').style.display = 'flex';
  document.getElementById('historyBody').innerHTML = '<p style="text-align:center;padding:32px;color:var(--text-muted);">加载中...</p>';

  try {
    const resp = await fetch('/api/history');
    const entries = await resp.json();

    if (!entries.length) {
      document.getElementById('historyBody').innerHTML =
        '<p style="text-align:center;padding:48px;color:var(--text-muted);">暂无历史记录</p>';
      return;
    }

    let html = '<div style="display:flex;flex-direction:column;gap:8px;">';
    entries.forEach(e => {
      const eid = e.id;
      if (!eid) return;  // skip invalid/legacy entries without a valid id
      const vText = e.verdict || '';
      let verdictCls = 'success';
      if (vText.includes('高度嫌疑')) verdictCls = 'danger';
      else if (vText.includes('可疑') || vText.includes('核查')) verdictCls = 'warning';
      else if (vText.includes('数据不足') || vText.includes('无法')) verdictCls = 'muted';
      let filesStr = (e.bid_files || []).slice(0, 3).map(f => (f || '').substring(0, 20) + (f.length > 20 ? '...' : '')).join(', ');
      if (e.bid_files.length > 3) filesStr += ` 等${e.bid_files.length}份`;

      html += `<div class="history-item">
        <div class="history-main" onclick="loadHistory('${eid}')">
          <div class="history-header">
            <span class="history-time">${escapeHtml(e.time)}</span>
            <span class="history-verdict ${verdictCls}">${escapeHtml(e.verdict)}</span>
          </div>
          <div class="history-files">📄 ${escapeHtml(filesStr)}</div>
          <div class="history-stats">
            <span>${e.total_pairs}组比对</span>
            <span style="color:var(--danger);">${e.abnormal_matches}异常</span>
            <span style="color:var(--success);">${e.template_matches}模板</span>
            ${e.ref_count > 0 ? `<span>📂 ${e.ref_count}份参考</span>` : ''}
          </div>
        </div>
        <button class="history-delete" title="删除" onclick="event.stopPropagation();deleteHistory('${eid}')">×</button>
      </div>`;
    });
    html += '</div>';
    document.getElementById('historyBody').innerHTML = html;
  } catch (err) {
    document.getElementById('historyBody').innerHTML =
      '<p style="text-align:center;padding:32px;color:var(--danger);">加载失败</p>';
  }
}

async function loadHistory(id) {
  if (!id) return;
  document.getElementById('historyModal').style.display = 'none';
  try {
    const resp = await fetch('/api/history/' + id);
    if (!resp.ok) throw new Error('请求失败: ' + resp.status);
    const data = await resp.json();
    if (data.error) { alert(data.error); return; }
    analysisResult = data;
    resultsSection.style.display = 'block';
    try {
      renderAllTabs();
    } catch (e) {
      console.error('Render error:', e);
      // Try rendering individually
      try { renderVerdict(); } catch(e) {}
      try { renderMetadata(); } catch(e) {}
      try { renderPersonnel(); } catch(e) {}
      try { renderSimilarity(); } catch(e) {}
      try { renderPricing(); } catch(e) {}
    }
    resultsSection.scrollIntoView({ behavior: 'smooth' });
  } catch (err) {
    console.error('History load error:', err);
    alert('加载失败: ' + (err.message || '未知错误'));
  }
}

async function deleteHistory(id) {
  if (!id) return;
  if (!confirm('确定删除此记录？')) return;
  const resp = await fetch('/api/history/' + id, { method: 'DELETE' });
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    alert(err.error || '删除失败，请刷新后重试');
  }
  openHistory();
}

// ── Download Report ──
btnDownload.addEventListener('click', async () => {
  if (!analysisResult) {
    alert('请先完成分析，再下载报告');
    return;
  }

  const originalText = btnDownload.innerHTML;
  btnDownload.disabled = true;
  btnDownload.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/></svg> 生成报告中...';

  try {
    const resp = await fetch('/api/report', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ analysis: analysisResult })
    });

    if (!resp.ok) {
      const errData = await resp.json().catch(() => null);
      throw new Error(errData?.error || `服务器错误 (${resp.status})`);
    }

    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = '围串标分析报告_' + new Date().toISOString().slice(0, 10).replace(/-/g, '') + '.docx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch (err) {
    console.error('Report download failed:', err);
    alert('报告生成失败: ' + (err.message || '未知错误'));
  } finally {
    btnDownload.disabled = false;
    btnDownload.innerHTML = originalText;
  }
});

document.getElementById('btnCloseHistory').addEventListener('click', () => {
  document.getElementById('historyModal').style.display = 'none';
});
document.getElementById('historyModal').addEventListener('click', e => {
  if (e.target === document.getElementById('historyModal')) {
    document.getElementById('historyModal').style.display = 'none';
  }
});

// ══════════════════════════════════════════════════════════
// 数据统计页 (Stats View)
// ══════════════════════════════════════════════════════════

const LEVEL_META = {
  high:      { label: '高度嫌疑', color: '#dc2626' },
  medium:    { label: '可疑',     color: '#d97706' },
  low:       { label: '未见异常', color: '#16a34a' },
  uncertain: { label: '数据不足', color: '#94a3b8' },
};

let _statsData = null;        // last fetched /api/stats payload
let _statsAnimated = false;   // skip entry animations on resize re-render

function _prefersReducedMotion() {
  return window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
}

// ── View switching ──
function switchView(name) {
  const isStats = name === 'stats';
  document.getElementById('view-analyze').hidden = isStats;
  document.getElementById('view-stats').hidden = !isStats;
  document.getElementById('navLinkAnalyze').classList.toggle('active', !isStats);
  document.getElementById('navLinkStats').classList.toggle('active', isStats);
  window.scrollTo({ top: 0 });
  if (isStats) refreshStats();
  if ((location.hash === '#stats') !== isStats) {
    history.replaceState(null, '', isStats ? '#stats' : '#analyze');
  }
}
// Deep-link support: /#stats opens the stats view on load
if (location.hash === '#stats') switchView('stats');

// Sticky nav gains a shadow once the page scrolls
window.addEventListener('scroll', function() {
  const nav = document.getElementById('siteNav');
  if (nav) nav.classList.toggle('scrolled', window.scrollY > 4);
}, { passive: true });

document.getElementById('btnRefreshStats').addEventListener('click', () => refreshStats(true));

// Re-layout the px-based trend chart on viewport changes (no refetch)
let _statsResizeTimer = null;
window.addEventListener('resize', function() {
  if (document.getElementById('view-stats').hidden || !_statsData) return;
  clearTimeout(_statsResizeTimer);
  _statsResizeTimer = setTimeout(function() {
    _statsAnimated = true;
    renderStats(_statsData);
  }, 200);
});

async function refreshStats(force) {
  const body = document.getElementById('statsBody');
  if (force || !_statsData) {
    body.innerHTML = '<div class="stats-error" style="color:var(--text-muted);">加载中...</div>';
  }
  try {
    const resp = await fetch('/api/stats');
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    _statsData = await resp.json();
    _statsAnimated = false;  // fresh fetch -> play entry animations
    renderStats(_statsData);
    document.getElementById('statsUpdated').textContent =
      '更新于 ' + new Date().toLocaleTimeString('zh-CN', { hour12: false });
  } catch (err) {
    console.error('Stats fetch failed:', err);
    body.innerHTML = '<div class="stats-error">统计加载失败，请确认服务器已启动</div>';
  }
}

// ── Number count-up ──
function _countUp(el, target, decimals) {
  const fmt = function(v) {
    return v.toLocaleString('zh-CN', {
      minimumFractionDigits: decimals || 0,
      maximumFractionDigits: decimals || 0,
    });
  };
  if (_prefersReducedMotion() || !target) { el.textContent = fmt(target || 0); return; }
  const dur = 900, t0 = performance.now();
  (function frame(t) {
    const p = Math.min(1, (t - t0) / dur);
    const eased = 1 - Math.pow(1 - p, 3);
    el.textContent = fmt(target * eased);
    if (p < 1) requestAnimationFrame(frame);
  })(t0);
}

// ── Render whole stats page ──
function renderStats(s) {
  const body = document.getElementById('statsBody');

  if (!s.total_analyses) {
    body.innerHTML = `
      <div class="stats-empty">
        <div class="stats-empty-icon">
          <svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/><path d="M3 3v18h18"/></svg>
        </div>
        <h3>暂无分析数据</h3>
        <p>完成首次围串标分析后，此处将展示风险分布、评分趋势与维度统计。</p>
        <button class="btn btn-primary" onclick="switchView('analyze')">前往分析</button>
      </div>`;
    return;
  }

  const total = s.total_analyses;
  const high = s.verdict_counts.high || 0;
  const highPct = Math.round(high / total * 100);

  // ── KPI cards ──
  body.innerHTML = `
    <div class="kpi-grid">
      <div class="kpi-card" style="--kpi-icon-bg:#eff4ff;--kpi-icon-fg:#1d4ed8;--kpi-glow:rgba(37,99,235,.07);">
        <div class="kpi-icon"><svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg></div>
        <div class="kpi-num" id="kpiAnalyses">0</div>
        <div class="kpi-label">累计分析</div>
        <div class="kpi-note accent">覆盖 ${s.total_documents.toLocaleString()} 份标书</div>
      </div>
      <div class="kpi-card" style="--kpi-icon-bg:#eef2ff;--kpi-icon-fg:#4f46e5;--kpi-glow:rgba(79,70,229,.07);">
        <div class="kpi-icon"><svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M13 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V9z"/><polyline points="13 2 13 9 20 9"/></svg></div>
        <div class="kpi-num" id="kpiDocs">0</div>
        <div class="kpi-label">分析标书总数</div>
        <div class="kpi-note">${s.total_pairs.toLocaleString()} 组两两比对</div>
      </div>
      <div class="kpi-card" style="--kpi-icon-bg:#fef2f2;--kpi-icon-fg:#dc2626;--kpi-glow:rgba(220,38,38,.07);">
        <div class="kpi-icon"><svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg></div>
        <div class="kpi-num" id="kpiHigh" style="color:#dc2626;">0</div>
        <div class="kpi-label">高风险案件</div>
        <div class="kpi-note danger">占全部分析 ${highPct}%</div>
      </div>
      <div class="kpi-card" style="--kpi-icon-bg:#fffbeb;--kpi-icon-fg:#d97706;--kpi-glow:rgba(217,119,6,.08);">
        <div class="kpi-icon"><svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2a10 10 0 100 20 10 10 0 000-20z"/><path d="M12 6v6l4 2"/></svg></div>
        <div class="kpi-num" id="kpiAvg">0</div>
        <div class="kpi-label">平均风险评分</div>
        <div class="kpi-note">满分 100 · 单次最高 ${_maxScoreSeen(s)} 分</div>
      </div>
      <div class="kpi-card" style="--kpi-icon-bg:#f0fdf4;--kpi-icon-fg:#16a34a;--kpi-glow:rgba(22,163,74,.07);">
        <div class="kpi-icon"><svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg></div>
        <div class="kpi-num" id="kpiAbnormal" style="color:#d97706;">0</div>
        <div class="kpi-label">累计异常匹配段</div>
        <div class="kpi-note success">另过滤模板段 ${s.total_template_matches.toLocaleString()} 处</div>
      </div>
    </div>

    <div class="stats-charts">
      <div class="chart-card">
        <div class="chart-card-title">结论分布</div>
        <div class="chart-card-sub">按综合判定结论统计 ${total} 次分析</div>
        <div class="donut-wrap">
          ${_donutSVG(s.verdict_counts, total)}
          <div class="donut-legend">
            ${['high', 'medium', 'low', 'uncertain'].map(function(k) {
              const cnt = s.verdict_counts[k] || 0;
              const pct = total ? Math.round(cnt / total * 100) : 0;
              return `<div class="legend-item">
                <span class="legend-dot" style="background:${LEVEL_META[k].color};"></span>
                <span class="legend-name">${LEVEL_META[k].label}</span>
                <span class="legend-count">${cnt}</span>
                <span class="legend-pct">${pct}%</span>
              </div>`;
            }).join('')}
          </div>
        </div>
      </div>
      <div class="chart-card wide">
        <div class="chart-card-title">风险评分趋势</div>
        <div class="chart-card-sub">按分析时序排列 · 虚线为 15 分（可疑）与 50 分（高度嫌疑）阈值</div>
        <div class="trend-wrap" id="trendWrap"></div>
        <div class="trend-legend">
          ${['high', 'medium', 'low', 'uncertain'].map(k =>
            `<span><i style="background:${LEVEL_META[k].color};"></i>${LEVEL_META[k].label}</span>`).join('')}
        </div>
      </div>
    </div>

    <div class="chart-card">
      <div class="chart-card-title">风险维度涉及情况</div>
      <div class="chart-card-sub">出现对应维度风险线索的案件数（条形）与累计命中条数（右侧）</div>
      <div class="dim-bars" id="dimBars"></div>
    </div>

    <div class="chart-card">
      <div class="chart-card-title">最近分析记录</div>
      <div class="chart-card-sub">点击「查看」载入完整分析结果</div>
      <div style="margin-top:14px;">${_recentTableHTML(s.recent)}</div>
    </div>
  `;

  // KPI count-ups
  _countUp(document.getElementById('kpiAnalyses'), s.total_analyses, 0);
  _countUp(document.getElementById('kpiDocs'), s.total_documents, 0);
  _countUp(document.getElementById('kpiHigh'), high, 0);
  _countUp(document.getElementById('kpiAvg'), s.avg_score || 0, 1);
  _countUp(document.getElementById('kpiAbnormal'), s.total_abnormal_matches, 0);

  _renderTrend(s.scores_over_time || []);
  _renderDimBars(s);

  // Donut draw-in: swap the zero-length dasharray for the real one after
  // the initial paint so the CSS transition animates the segments.
  if (!_statsAnimated) {
    requestAnimationFrame(function() {
      requestAnimationFrame(function() {
        document.querySelectorAll('.donut-seg').forEach(function(c) {
          c.setAttribute('stroke-dasharray', c.getAttribute('data-final'));
        });
      });
    });
  }
}

function _maxScoreSeen(s) {
  let max = 0;
  (s.scores_over_time || []).forEach(function(p) {
    if (typeof p.score === 'number' && p.score > max) max = p.score;
  });
  return max % 1 === 0 ? max : max.toFixed(1);
}

// ── Donut chart (pure SVG) ──
function _donutSVG(counts, total) {
  const R = 74, C = 2 * Math.PI * R;
  const keys = ['high', 'medium', 'low', 'uncertain'].filter(k => (counts[k] || 0) > 0);
  const gap = keys.length > 1 ? 3 : 0;  // px gap between segments
  let startFrac = 0, segs = '';
  keys.forEach(function(k) {
    const frac = (counts[k] || 0) / total;
    const len = Math.max(frac * C - gap, 0.5);
    const startDeg = startFrac * 360 - 90;
    const finalDash = len.toFixed(2) + ' ' + (C - len).toFixed(2);
    segs += `<circle class="donut-seg" cx="100" cy="100" r="${R}"
      stroke="${LEVEL_META[k].color}" stroke-width="26"
      stroke-dasharray="${_statsAnimated ? finalDash : '0 ' + C.toFixed(2)}"
      data-final="${finalDash}"
      transform="rotate(${startDeg.toFixed(2)} 100 100)"></circle>`;
    startFrac += frac;
  });
  return `<svg class="donut-svg" viewBox="0 0 200 200" role="img" aria-label="结论分布环图">
    ${segs}
    <text class="donut-center-num" x="100" y="98" text-anchor="middle">${total}</text>
    <text class="donut-center-label" x="100" y="116" text-anchor="middle">次分析</text>
  </svg>`;
}

// ── Trend chart (px-coordinate SVG so tooltips map 1:1) ──
function _renderTrend(points) {
  const wrap = document.getElementById('trendWrap');
  if (!wrap) return;
  const pts = points.filter(function(p) { return typeof p.score === 'number'; });
  if (pts.length === 0) {
    wrap.innerHTML = '<p class="empty-note">暂无评分数据</p>';
    return;
  }

  wrap.innerHTML = '<svg class="trend-svg" id="trendSvg"></svg><div class="trend-tooltip" id="trendTooltip"></div>';
  const svg = document.getElementById('trendSvg');
  const W = Math.max(svg.clientWidth || 0, 320);
  const H = Math.max(svg.clientHeight || 0, 236);
  svg.setAttribute('width', W);
  svg.setAttribute('height', H);

  const padL = 36, padR = 16, padT = 16, padB = 28;
  const innerW = W - padL - padR, innerH = H - padT - padB;
  const x = function(i) {
    return pts.length === 1 ? padL + innerW / 2 : padL + innerW * i / (pts.length - 1);
  };
  const y = function(score) { return padT + innerH * (1 - score / 100); };

  let grid = '', labels = '';
  [0, 25, 50, 75, 100].forEach(function(v) {
    const gy = y(v);
    grid += `<line class="trend-grid" x1="${padL}" y1="${gy}" x2="${W - padR}" y2="${gy}"></line>`;
    labels += `<text class="trend-axis-label" x="${padL - 8}" y="${gy + 3}" text-anchor="end">${v}</text>`;
  });
  // Threshold guides (same cutoffs as the score card on the verdict tab)
  let guides = '';
  [[15, '#d97706'], [50, '#dc2626']].forEach(function(t) {
    guides += `<line class="trend-threshold" x1="${padL}" y1="${y(t[0])}" x2="${W - padR}" y2="${y(t[0])}" stroke="${t[1]}" opacity=".45"></line>`;
  });
  // X labels: first, last, and a few evenly-spaced in between
  const maxLabels = Math.min(6, pts.length);
  for (let i = 0; i < maxLabels; i++) {
    const idx = maxLabels === 1 ? 0 : Math.round(i * (pts.length - 1) / (maxLabels - 1));
    const t = (pts[idx].time || '').slice(5, 10);  // MM-DD
    labels += `<text class="trend-axis-label" x="${x(idx)}" y="${H - 8}" text-anchor="middle">${escapeHtml(t)}</text>`;
  }

  const coords = pts.map(function(p, i) { return [x(i), y(p.score)]; });
  const lineD = coords.map(function(c, i) {
    return (i === 0 ? 'M' : 'L') + c[0].toFixed(1) + ' ' + c[1].toFixed(1);
  }).join(' ');
  const areaD = lineD +
    ` L ${coords[coords.length - 1][0].toFixed(1)} ${y(0)}` +
    ` L ${coords[0][0].toFixed(1)} ${y(0)} Z`;

  let dots = '';
  pts.forEach(function(p, i) {
    dots += `<circle class="trend-dot" data-idx="${i}" cx="${coords[i][0].toFixed(1)}" cy="${coords[i][1].toFixed(1)}" r="4.5" stroke="${LEVEL_META[p.level] ? LEVEL_META[p.level].color : '#94a3b8'}"></circle>`;
  });

  svg.innerHTML = `
    <defs>
      <linearGradient id="trendAreaGrad" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%" stop-color="rgba(37,99,235,.22)"/>
        <stop offset="100%" stop-color="rgba(37,99,235,0)"/>
      </linearGradient>
      <linearGradient id="trendLineGrad" x1="${padL}" y1="0" x2="${W - padR}" y2="0" gradientUnits="userSpaceOnUse">
        <stop offset="0%" stop-color="#2563eb"/>
        <stop offset="100%" stop-color="#6366f1"/>
      </linearGradient>
    </defs>
    ${grid}${guides}${labels}
    <path class="trend-area${_statsAnimated ? ' shown' : ''}" d="${areaD}" fill="url(#trendAreaGrad)"></path>
    <path class="trend-line${_statsAnimated ? ' shown' : ''}" d="${lineD}" pathLength="1"
      fill="none" stroke="url(#trendLineGrad)" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"></path>
    ${dots}
  `;

  // Draw-in + tooltips
  if (!_statsAnimated) {
    requestAnimationFrame(function() {
      requestAnimationFrame(function() {
        const line = svg.querySelector('.trend-line');
        const area = svg.querySelector('.trend-area');
        if (line) line.classList.add('shown');
        if (area) area.classList.add('shown');
      });
    });
  }

  const tip = document.getElementById('trendTooltip');
  svg.querySelectorAll('.trend-dot').forEach(function(dot) {
    dot.addEventListener('mouseenter', function() {
      const p = pts[+dot.getAttribute('data-idx')];
      if (!p) return;
      const lvl = LEVEL_META[p.level] || LEVEL_META.uncertain;
      tip.innerHTML = `<div class="tt-time">${escapeHtml(p.time || '')}</div>` +
        `<div class="tt-score" style="color:${lvl.color === '#94a3b8' ? '#cbd5e1' : lvl.color};">${p.score}<span style="font-size:11px;color:#94a3b8;"> / 100</span></div>` +
        `<div class="tt-verdict ${p.level}">${lvl.label}</div>`;
      tip.style.left = (+dot.getAttribute('cx')) + 'px';
      tip.style.top = (+dot.getAttribute('cy')) + 'px';
      tip.classList.add('show');
    });
    dot.addEventListener('mouseleave', function() { tip.classList.remove('show'); });
  });
}

// ── Dimension bars ──
function _renderDimBars(s) {
  const holder = document.getElementById('dimBars');
  if (!holder) return;
  const total = s.total_analyses || 0;
  const DIMS = [
    { key: 'metadata',  cls: 'meta',  name: '元数据一致',  sub: '创建者/修改者/机器码等' },
    { key: 'personnel', cls: 'person', name: '人员交叉',   sub: '同名/同手机号/同身份证' },
    { key: 'similarity', cls: 'sim',  name: '文本查重',   sub: '高风险异常一致段落' },
    { key: 'pricing',   cls: 'price', name: '报价异常',   sub: '报价一致或规律性差异' },
  ];
  holder.innerHTML = DIMS.map(function(d) {
    const hits = s.dimension_hits[d.key] || 0;
    const totals = s.dimension_totals[d.key] || 0;
    const pct = total ? Math.round(hits / total * 100) : 0;
    return `<div class="bar-row">
      <div class="bar-row-top">
        <span class="bar-name">${d.name}<span class="bar-sub">${d.sub} · 累计 ${totals.toLocaleString()} 条</span></span>
        <span class="bar-count"><b>${hits}</b> / ${total} 件 (${pct}%)</span>
      </div>
      <div class="bar-track"><div class="bar-fill ${d.cls}" data-pct="${pct}"></div></div>
    </div>`;
  }).join('');
  const apply = function() {
    holder.querySelectorAll('.bar-fill').forEach(function(el) {
      el.style.width = el.getAttribute('data-pct') + '%';
    });
  };
  if (!_statsAnimated) {
    requestAnimationFrame(function() { requestAnimationFrame(apply); });
  } else {
    apply();
  }
}

// ── Recent records table ──
function _recentTableHTML(recent) {
  if (!recent || recent.length === 0) {
    return '<p class="empty-note">暂无记录</p>';
  }
  let rows = '';
  recent.forEach(function(e) {
    if (!e.id) return;
    const lvl = LEVEL_META[e.level] || LEVEL_META.uncertain;
    const score = (typeof e.score === 'number') ? (e.score % 1 === 0 ? e.score : e.score.toFixed(1)) : '-';
    rows += `<tr>
      <td style="white-space:nowrap;font-family:var(--font-mono);font-size:11.5px;">${escapeHtml(e.time || '')}</td>
      <td style="text-align:center;">${e.bid_count || 0} 份</td>
      <td class="score-cell ${e.level}">${score}</td>
      <td><span class="level-chip ${e.level}">${lvl.label}</span></td>
      <td style="color:var(--text-secondary);">${escapeHtml(e.verdict || '')}</td>
      <td style="text-align:right;"><button class="recent-open-btn" onclick="openHistoryRecord('${escapeHtml(e.id)}')">查看</button></td>
    </tr>`;
  });
  return `<table class="data-table" style="margin-bottom:0;">
    <thead><tr><th>时间</th><th style="text-align:center;">标书数</th><th>评分</th><th>结论</th><th>判定详情</th><th></th></tr></thead>
    <tbody>${rows}</tbody>
  </table>`;
}

// Load a history record from the stats page, then jump back to the analysis view
async function openHistoryRecord(id) {
  switchView('analyze');
  await loadHistory(id);
}
