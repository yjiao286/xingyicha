// ── State ──
let selectedFiles = [];
let selectedRefFiles = [];
let analysisResult = null;
let fileGroups = {}; // {filename: groupName} — same group = same bidder

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
const progressText = document.getElementById('progressText');
const progressSteps = document.getElementById('progressSteps');
const resultsSection = document.getElementById('resultsSection');

// ── Bid File Selection ──
uploadArea.addEventListener('click', () => fileInput.click());
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
refUploadArea.addEventListener('dragover', e => { e.preventDefault(); refUploadArea.classList.add('drag-over'); });
refUploadArea.addEventListener('dragleave', () => refUploadArea.classList.remove('drag-over'));
refUploadArea.addEventListener('drop', e => {
  e.preventDefault();
  refUploadArea.classList.remove('drag-over');
  addFiles(e.dataTransfer.files, 'ref');
});
refFileInput.addEventListener('change', e => addFiles(e.target.files, 'ref'));

const VALID_EXTS = ['.docx', '.doc', '.pdf'];

function escapeHtml(str) {
  const div = document.createElement('div');
  div.appendChild(document.createTextNode(str));
  return div.innerHTML;
}

function formatSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function startProgress() {
  progressPanel.style.display = 'block';
  progressFill.style.width = '0%';
  progressText.textContent = '上传中...';
  progressSteps.innerHTML = '<span class="progress-step active">上传文件</span>';
}

function updateProgress(event) {
  progressFill.style.width = event.percent + '%';
  progressText.textContent = event.label;

  // Reset all active states
  var existing = progressSteps.querySelectorAll('.progress-step');
  existing.forEach(function(el) { el.classList.remove('active'); });

  if (event.detail) {
    // Find or create step
    var found = null;
    existing.forEach(function(el) {
      if (el.getAttribute('data-step') === event.step) found = el;
    });
    if (!found) {
      found = document.createElement('span');
      found.className = 'progress-step';
      found.setAttribute('data-step', event.step);
      found.textContent = event.detail;
      progressSteps.appendChild(found);
    }
    found.classList.add('active');
  }
}

function finishProgress() {
  progressFill.style.width = '100%';
  progressText.textContent = '分析完成';
  var steps = progressSteps.querySelectorAll('.progress-step');
  steps.forEach(function(el) { el.classList.remove('active'); el.classList.add('done'); });
  setTimeout(function() { progressPanel.style.display = 'none'; }, 2000);
}

function addFiles(files, type) {
  const validFiles = Array.from(files).filter(f =>
    VALID_EXTS.some(ext => f.name.toLowerCase().endsWith(ext))
  );
  if (validFiles.length === 0) { alert('请选择 .docx / .doc / .pdf 格式的文件'); return; }

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
          .replace(/\.(docx|doc|pdf)$/i, '').trim() || f.name;
        fileGroups[key] = defaultGroup;
      }
      const group = fileGroups[key] || '';
      html += '<div class="file-group-row">';
      html += '<span class="file-tag" style="flex:1;min-width:0;">';
      html += '<span style="word-break:break-all;">' + escapeHtml(f.name) + ' (' + formatSize(f.size) + ')</span>';
      html += '<span class="remove-btn" title="移除" onclick="removeFile(' + i + ',\'bid\')">&times;</span>';
      html += '</span>';
      html += '<input class="group-input" value="' + escapeHtml(group) + '" placeholder="投标人名称" onchange="updateGroup(\'' + key.replace(/'/g, "\\'") + '\', this.value)" title="相同名称的文件将合并为一个投标人"/>';
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
  const target = type === 'ref' ? selectedRefFiles : selectedFiles;
  target.splice(idx, 1);
  renderAllFileLists();
  updateButtons();
  if (selectedFiles.length === 0 && selectedRefFiles.length === 0) {
    resultsSection.style.display = 'none';
    analysisResult = null;
  }
  fileInput.value = '';
  refFileInput.value = '';
}

btnClear.addEventListener('click', () => {
  selectedFiles = [];
  selectedRefFiles = [];
  renderAllFileLists();
  updateButtons();
  resultsSection.style.display = 'none';
  analysisResult = null;
  fileInput.value = '';
  refFileInput.value = '';
});

function updateButtons() {
  btnAnalyze.disabled = selectedFiles.length < 2;
  if (selectedFiles.length < 2 && selectedFiles.length > 0) {
    btnAnalyze.innerHTML = '<span class="btn-icon">&#9881;</span> 请至少上传2份标书';
  } else if (selectedFiles.length >= 2) {
    const extra = selectedRefFiles.length > 0 ? ` + ${selectedRefFiles.length}份模板` : '';
    btnAnalyze.innerHTML = '<span class="btn-icon">&#9881;</span> 开始分析' + extra;
  } else {
    btnAnalyze.innerHTML = '<span class="btn-icon">&#9881;</span> 请至少上传2份标书文件';
  }
}

// ── Analyze (streaming progress with NDJSON) ──
btnAnalyze.addEventListener('click', async () => {
  if (selectedFiles.length < 2) {
    alert('请至少上传2份标书文件(.docx/.doc/.pdf)');
    return;
  }

  btnAnalyze.disabled = true;
  btnAnalyze.innerHTML = '<span class="btn-icon">&#9881;</span> 分析中...';
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
      const errData = await resp.json();
      alert('分析失败: ' + (errData.error || '服务器错误'));
      finishProgress();
      updateButtons();
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
          } else if (event.type === 'result') {
            analysisResult = event.data;
            resultsSection.style.display = 'block';
            renderAllTabs();
            resultsSection.scrollIntoView({ behavior: 'smooth' });
          } else if (event.type === 'error') {
            alert('分析失败: ' + event.message);
          }
        } catch (e) {
          console.warn('NDJSON parse error:', e);
        }
      }
    }
  } catch (err) {
    console.error('Analysis failed:', err);
    alert('分析失败，请确认服务器已启动: ' + err.message);
  } finally {
    finishProgress();
    updateButtons();
  }
});

// ── Tab Switching ──
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(btn.dataset.tab).classList.add('active');
  });
});

// ── Render All ──
function renderAllTabs() {
  if (!analysisResult) return;
  renderVerdict();
  renderMetadata();
  renderPersonnel();
  renderSimilarity();
  renderPricing();
}

// ── Verdict Tab ──
function renderVerdict() {
  const v = analysisResult.verdict;
  const banner = document.getElementById('verdictBanner');
  const conclusion = v.conclusion;
  let cls = 'warning', icon = '';
  if (conclusion.includes('高度嫌疑')) { cls = 'suspect'; icon = '⚠️ '; }
  else if (conclusion.includes('未发现')) { cls = 'clean'; icon = '✅ '; }
  banner.className = 'verdict-banner ' + cls;
  banner.textContent = icon + '判定结论: ' + conclusion;

  // Show reference docs info
  const refDocs = analysisResult.ref_docs || [];
  if (refDocs.length > 0) {
    banner.textContent += ' (已扣除' + refDocs.length + '份模板文档)';
  }

  let html = '';
  v.clauses.forEach(c => {
    const satisfied = c.satisfied;
    let cardCls = 'uncertain', tagCls = 'tag-uncertain', tagText = '无法判断';
    if (satisfied === true) { cardCls = 'satisfied'; tagCls = 'tag-satisfied'; tagText = '满足'; }
    else if (satisfied === false) { cardCls = 'not-satisfied'; tagCls = 'tag-not'; tagText = '不满足'; }

    html += `<div class="clause-card ${cardCls}">
      <div class="clause-header">
        <strong>${c.clause}</strong>
        <span class="clause-tag ${tagCls}">${tagText}</span>
      </div>
      <p style="font-size:14px;margin-bottom:6px;">${escapeHtml(c.description)}</p>`;

    if (c.evidence && c.evidence.length > 0) {
      html += '<ul class="clause-evidence">';
      c.evidence.forEach(e => { html += `<li>${escapeHtml(String(e))}</li>`; });
      html += '</ul>';
    }
    if (c.evidence_level) {
      const lvlCls = c.evidence_level === '强' ? 'level-strong' : c.evidence_level === '中' ? 'level-medium' : 'level-none';
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
      const icon = m.severity === 'high' ? '⚠️' : '📋';
      const cls = m.severity === 'high' ? 'meta-critical' : 'meta-normal';
      mhtml += `<div class="meta-match-card ${cls}">
        <div class="meta-match-icon">${icon}</div>
        <div class="meta-match-body">
          <div class="meta-match-field">${escapeHtml(m.field)}</div>
          <div class="meta-match-value">${escapeHtml(String(m.value).substring(0, 200))}</div>
          ${m.pair ? `<div class="meta-match-pair">📄 ${escapeHtml(m.pair)}</div>` : ''}
        </div>
        <span class="meta-match-badge ${m.severity === 'high' ? 'badge-high' : 'badge-medium'}">${m.verdict}</span>
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
    fhtml += `<h3 style="margin:10px 0 6px;font-size:14px;color:#4361ee;">${escapeHtml(shortName)}</h3>`;
    fhtml += '<table class="data-table"><thead><tr><th>属性</th><th>值</th></tr></thead><tbody>';
    const rows = [
      ['文件类型', f._error ? '读取异常' : (f.name.endsWith('.pdf') ? 'PDF' : f.name.endsWith('.doc') ? 'DOC(旧版)' : 'DOCX')],
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
    html += '<h3 style="margin-bottom:8px;">' + escapeHtml(f.name) + '</h3>';

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
      html += '<h4 style="margin:12px 0 6px;">项目团队成员</h4>';
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
    html = '<p style="color:#888;text-align:center;padding:32px;">未从标书中提取到人员信息（法定代表人、授权代表、项目成员等）</p>';
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
      mhtml += '<div class="match-card">';
      mhtml += '<span class="match-badge ' + sevClass + '">' + sevLabel + '</span>';
      mhtml += '<strong>' + escapeHtml(m.type) + '</strong>';
      mhtml += '<p style="margin-top:4px;font-size:14px;">' + escapeHtml(m.detail) + '</p>';
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
    document.getElementById('personnelMatches').innerHTML = '<p style="color:#888;">未发现人员交叉异常</p>';
  } else {
    document.getElementById('personnelMatches').innerHTML = mhtml;
  }
}

// ── Similarity Tab ──
let _allMatchRefs = []; // flat list of all matches for modal navigation

function renderSimilarity() {
  if (!analysisResult) return;
  const s = analysisResult.text_similarity;

  const filterEl = document.querySelector('input[name="simFilter"]:checked');
  const filter = filterEl ? filterEl.value : 'all';

  let totalMatches = 0, totalAbnormal = 0, totalTemplate = s.template_matches || 0;
  s.pair_results.forEach(p => {
    totalMatches += p.total_matches;
    totalAbnormal += p.abnormal_count;
  });

  document.getElementById('similaritySummary').innerHTML = `
    <div class="summary-stat">
      <div class="stat-card" style="cursor:pointer" onclick="document.querySelector('input[value=all]').click();renderSimilarity();">
        <div class="stat-num">${totalMatches}</div>
        <div class="stat-label">总匹配段落数</div>
      </div>
      <div class="stat-card" style="cursor:pointer" onclick="document.querySelector('input[value=abnormal]').click();renderSimilarity();">
        <div class="stat-num danger">${totalAbnormal}</div>
        <div class="stat-label">异常一致段落数</div>
      </div>
      <div class="stat-card" style="cursor:pointer" onclick="document.querySelector('input[value=template]').click();renderSimilarity();">
        <div class="stat-num" style="color:#16a34a;">${totalTemplate}</div>
        <div class="stat-label">模板匹配(已扣除)</div>
      </div>
    </div>
    <ul class="finding-list">${s.findings.map(f => `<li>${escapeHtml(f)}</li>`).join('')}</ul>
  `;

  // Build flat list for modal nav
  _allMatchRefs = [];

  // ── Pair overview cards ──
  let overview = '<div class="pair-overview-grid">';
  s.pair_results.forEach((pr, pi) => {
    overview += `<div class="pair-overview-card" onclick="scrollToPair(${pi})" style="cursor:pointer;">
      <div class="pair-overview-header">对比 ${pi + 1}</div>
      <div style="font-size:12px;color:#666;margin:4px 0;">${shortenName(pr.file1)} ↔ ${shortenName(pr.file2)}</div>
      <div style="display:flex;gap:8px;font-size:12px;">
        <span style="color:#dc2626;font-weight:600;">${pr.abnormal_count}异常</span>
        <span style="color:#16a34a;">${pr.template_count || 0}模板</span>
        <span style="color:#888;">${pr.total_matches}总计</span>
      </div>
    </div>`;
  });
  overview += '</div>';
  document.getElementById('similaritySummary').innerHTML += `
    <div id="pairOverview">${overview}</div>
    <div style="margin-top:8px;">
      <button class="btn btn-sm btn-outline" onclick="expandAllPairs()">展开全部</button>
      <button class="btn btn-sm btn-outline" onclick="collapseAllPairs()" style="margin-left:4px;">折叠全部</button>
    </div>`;

  const PAGE_SIZE = 5;

  // ── Per-pair detail sections (collapsed by default, paginated) ──
  let dhtml = '';
  s.pair_results.forEach((pr, pairIdx) => {
    const pairId = `pair-${pairIdx}`;
    const abnormalMatches = pr.matches.filter(m => m.abnormal);
    const templateMatches = pr.matches.filter(m => !m.abnormal);
    const showAbnormal = filter === 'all' || filter === 'abnormal';
    const showTemplate = filter === 'all' || filter === 'template';

    dhtml += `<div class="pair-section" id="${pairId}">
      <div class="pair-header" onclick="togglePair('${pairId}')">
        <span class="pair-toggle" id="${pairId}-toggle">▶</span>
        <span class="pair-title">对比 ${pairIdx + 1}: ${shortenName(pr.file1, 15)} ↔ ${shortenName(pr.file2, 15)}</span>
        <span class="pair-stats">
          <span style="color:#dc2626;">${pr.abnormal_count}异常</span>
          <span style="color:#16a34a;margin-left:8px;">${pr.template_count || 0}模板</span>
          <span style="color:#888;margin-left:8px;">${pr.total_matches}总计</span>
        </span>
      </div>
      <div class="pair-body" id="${pairId}-body" style="display:none;">`;

    // Render helper: show first N items + "show more" button
    const renderPaginated = (matches, label, colorClass, bgStyle) => {
      if (matches.length === 0) return '';
      let html = `<p style="font-weight:600;color:${colorClass};margin:8px 0 4px;">${label} (${matches.length}处):</p>`;
      const visible = matches.slice(0, PAGE_SIZE);
      const hidden = matches.slice(PAGE_SIZE);

      visible.forEach(m => {
        const refIdx = _allMatchRefs.length;
        _allMatchRefs.push({ match: m, file1: pr.file1, file2: pr.file2, type: m.abnormal ? 'abnormal' : 'template' });
        html += `<div class="text-match-item" style="border-left:3px solid ${colorClass};">
          <div class="text-match-header">
            <span class="text-match-num" style="${bgStyle}">#${m.index}</span>
            <span class="text-match-length">${m.length}字</span>
            ${(m.reasons||[]).map(r => `<span class="text-match-reason">${escapeHtml(r)}</span>`).join('')}
            <button class="match-locate-btn" onclick="openMatchModal(${refIdx})">📍 定位</button>
          </div>
          <div class="text-match-content">${escapeHtml(m.text.substring(0, 200))}${m.text.length > 200 ? '...' : ''}</div>
        </div>`;
      });

      if (hidden.length > 0) {
        html += `<div id="${pairId}-more-${label.replace(/[^a-z]/g,'')}" style="display:none;">`;
        hidden.forEach(m => {
          const refIdx = _allMatchRefs.length;
          _allMatchRefs.push({ match: m, file1: pr.file1, file2: pr.file2, type: m.abnormal ? 'abnormal' : 'template' });
          html += `<div class="text-match-item" style="border-left:3px solid ${colorClass};">
            <div class="text-match-header">
              <span class="text-match-num" style="${bgStyle}">#${m.index}</span>
              <span class="text-match-length">${m.length}字</span>
              <button class="match-locate-btn" onclick="openMatchModal(${refIdx})">📍 定位</button>
            </div>
            <div class="text-match-content">${escapeHtml(m.text.substring(0, 200))}${m.text.length > 200 ? '...' : ''}</div>
          </div>`;
        });
        html += '</div>';
        html += `<button class="btn btn-sm btn-outline" style="margin-top:4px;"
          onclick="toggleMore('${pairId}-more-${label.replace(/[^a-z]/g,'')}', this)">显示全部 ${hidden.length} 项</button>`;
      }
      return html;
    };

    if (showAbnormal) {
      dhtml += renderPaginated(abnormalMatches, '异常一致段落', '#dc2626', '');
    }
    if (showTemplate) {
      dhtml += renderPaginated(templateMatches, '模板匹配段落', '#16a34a', 'background:#f0fdf4;color:#16a34a;');
    }
    if (filter === 'abnormal' && abnormalMatches.length === 0) dhtml += '<p style="color:#888;">无异常一致段落</p>';
    if (filter === 'template' && templateMatches.length === 0) dhtml += '<p style="color:#888;">无模板匹配段落</p>';

    dhtml += '</div></div>';
  });
  document.getElementById('similarityDetails').innerHTML = dhtml;
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
          tableHtml += '<p>📌 <strong>' + escapeHtml(label) + '</strong>: 全部一致 (' + vals[0].toLocaleString() + ' 元)</p>';
        } else if (sameAll === false && vals.length >= 2) {
          tableHtml += '<p>⚠️ <strong>' + escapeHtml(label) + '</strong>: 存在差异 — ' + vals.map(function(v) { return v.toLocaleString(); }).join(' / ') + ' 元</p>';
        }
      });
      tableHtml += '</div>';
    }
  } else {
    tableHtml += '<p style="color:#9ca3af;text-align:center;padding:20px;">暂无报价数据</p>';
  }
  document.getElementById('pricingTable').innerHTML = tableHtml;

  // ── Sub-item Comparison ──
  var subHtml = '';
  if (p.subItemCompare && p.subItemCompare.length > 0) {
    subHtml += '<div class="section-title" style="margin-top:20px;">📊 分项明细比对</div>';
    p.subItemCompare.forEach(function(c, ci) {
      subHtml += '<div style="margin-bottom:14px;border:1px solid var(--border-light);border-radius:8px;padding:12px;">';
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

function expandAllPairs() {
  document.querySelectorAll('.pair-body').forEach(el => el.style.display = 'block');
  document.querySelectorAll('.pair-toggle').forEach(el => el.textContent = '▼');
}
function collapseAllPairs() {
  document.querySelectorAll('.pair-body').forEach(el => el.style.display = 'none');
  document.querySelectorAll('.pair-toggle').forEach(el => el.textContent = '▶');
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
  if (document.getElementById('matchModal').style.display !== 'flex') return;
  if (e.key === 'ArrowLeft') document.getElementById('btnPrevMatch').click();
  if (e.key === 'ArrowRight') document.getElementById('btnNextMatch').click();
  if (e.key === 'Escape') document.getElementById('btnCloseModal').click();
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
  document.getElementById('diffContent1').innerHTML = renderContextWithHighlight(m.ctx1 || matchText, matchText);
  document.getElementById('diffContent2').innerHTML = renderContextWithHighlight(m.ctx2 || matchText, matchText);
}

function renderContextWithHighlight(ctx, matchText) {
  if (!ctx) return escapeHtml(matchText || '');
  if (!matchText) return escapeHtml(ctx);

  // Find match text position in context
  var idx = ctx.indexOf(matchText);
  if (idx === -1) {
    // Try with trimmed match text
    var trimmed = matchText.replace(/^[\s\n\r]+|[\s\n\r]+$/g, '');
    idx = ctx.indexOf(trimmed);
    if (idx !== -1) matchText = trimmed;
  }
  if (idx === -1) {
    // Try finding first 10 chars of match as fallback
    var short = matchText.substring(0, Math.min(20, matchText.length));
    idx = ctx.indexOf(short);
    if (idx !== -1) {
      // Extend to include full match if possible
      matchText = ctx.substring(idx, Math.min(ctx.length, idx + matchText.length));
    }
  }

  if (idx === -1) return escapeHtml(ctx);

  var before = ctx.substring(0, idx);
  var after = ctx.substring(idx + matchText.length);
  return escapeHtml(before) + '<mark class="match-highlight">' + escapeHtml(matchText) + '</mark>' + escapeHtml(after);
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
      const verdictCls = (e.verdict || '').includes('高度嫌疑') ? 'danger' : 'success';
      let filesStr = (e.bid_files || []).slice(0, 3).map(f => (f || '').substring(0, 20) + (f.length > 20 ? '...' : '')).join(', ');
      if (e.bid_files.length > 3) filesStr += ` 等${e.bid_files.length}份`;

      html += `<div class="history-item">
        <div class="history-main" onclick="loadHistory('${e.id}')">
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
        <button class="history-delete" title="删除" onclick="event.stopPropagation();deleteHistory('${e.id}')">×</button>
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
  if (!confirm('确定删除此记录？')) return;
  await fetch('/api/history/' + id, { method: 'DELETE' });
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
