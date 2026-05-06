let workspace = null;
let selectedFiles = [];

const $ = (id) => document.getElementById(id);
const html = (s) => String(s ?? '').replace(/[&<>]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[ch]));

$('files').addEventListener('change', (e) => {
  selectedFiles = Array.from(e.target.files || []);
  renderFileList();
});

function renderFileList(){
  const box = $('fileList');
  if(!selectedFiles.length){ box.innerHTML = ''; return; }
  box.innerHTML = selectedFiles.map((f,i)=>`<div class="file-pill"><span>${html(f.name)} · ${(f.size/1024).toFixed(1)} KB</span><button type="button" class="danger" onclick="removeFile(${i})">Delete</button></div>`).join('');
}
window.removeFile = (i) => { selectedFiles.splice(i,1); const dt = new DataTransfer(); selectedFiles.forEach(f=>dt.items.add(f)); $('files').files = dt.files; renderFileList(); };

$('intakeForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  $('error').textContent='';
  const btn = $('analyzeBtn'); btn.disabled=true; btn.textContent='Analyzing files and data...';
  try{
    const fd = new FormData();
    fd.append('note', $('note').value || '');
    fd.append('journal', $('journal').value || '');
    fd.append('study_design', $('studyDesign').value || '');
    selectedFiles.forEach(f => fd.append('files', f));
    const res = await fetch('/api/analyze', { method:'POST', body:fd });
    if(!res.ok){ const er = await res.json().catch(()=>({detail:res.statusText})); throw new Error(er.detail || 'Analysis failed'); }
    workspace = await res.json();
    renderWorkspace();
  } catch(err){ $('error').textContent = err.message; }
  finally{ btn.disabled=false; btn.textContent='Analyze Study'; }
});

function renderWorkspace(){
  $('workspace').classList.remove('hidden'); $('nav').classList.remove('hidden');
  $('projectStatus').innerHTML = `<span class="eyebrow">Project status</span><strong>Workspace built</strong><p>${workspace.audit.rows} rows · ${workspace.audit.columns} variables · ${html(workspace.journal)}</p>`;
  renderFeasibility(); renderAudit(); renderAnalysis(); renderManuscript(); renderExport();
  document.querySelectorAll('#nav button').forEach(b => b.onclick = () => scrollToSection(b.dataset.target));
  scrollToSection('feasibility');
}
function scrollToSection(id){
  document.querySelectorAll('#nav button').forEach(b => b.classList.toggle('active', b.dataset.target===id));
  $(id).scrollIntoView({behavior:'smooth', block:'start'});
}

function renderFeasibility(){
  const r = workspace.research, a = workspace.audit;
  $('feasibility').innerHTML = `
    <span class="eyebrow">Step 1</span><h2>Feasibility review</h2>
    <div class="cards">
      <div class="card"><span>Rows</span><strong>${a.rows}</strong><span>records detected</span></div>
      <div class="card"><span>Variables</span><strong>${a.columns}</strong><span>columns detected</span></div>
      <div class="card"><span>Patient ID candidate</span><strong>${html(a.patient_id_candidate || 'Not detected')}</strong><span>${a.duplicate_id_count ?? '—'} duplicate IDs</span></div>
    </div>
    <div class="ok"><strong>Recommended primary question</strong><br>${html(r.recommended_primary_question)}</div>
    <div class="card"><strong>Hypothesis</strong><p>${html(r.hypothesis)}</p></div>
    <h3>Secondary / exploratory questions</h3>${r.secondary_questions.map(q=>`<span class="tag">${html(q)}</span>`).join('')}
    <h3>Needs caution or is not feasible yet</h3>${r.not_feasible_or_needs_caution.map(x=>`<div class="warning">${html(x)}</div>`).join('')}
    <h3>Detected candidates</h3>
    <p><strong>Exposure candidates:</strong> ${(r.detected_exposure_candidates||[]).map(x=>`<span class="tag">${html(x)}</span>`).join('') || 'None detected'}</p>
    <p><strong>Outcome candidates:</strong> ${(r.detected_outcome_candidates||[]).map(x=>`<span class="tag">${html(x)}</span>`).join('') || 'None detected'}</p>`;
}

function renderAudit(){
  const a = workspace.audit;
  $('audit').innerHTML = `
    <span class="eyebrow">Step 2</span><h2>Dataset audit</h2>
    <h3>File inventory</h3><div class="scroll">${table(workspace.files)}</div>
    <h3>Variable classification</h3>
    <div class="cards">${Object.entries(a.variable_classification).map(([k,v])=>`<div class="card"><span>${html(k.replaceAll('_',' '))}</span><strong>${v.length}</strong><p>${v.slice(0,8).map(x=>`<span class="tag">${html(x)}</span>`).join('')}</p></div>`).join('')}</div>
    <h3>Missingness</h3><div class="scroll">${table(a.missingness.slice(0,50))}</div>
    <h3>Missingness figure</h3><div class="figure">${workspace.analysis.missingness_figure_svg || ''}</div>
    <h3>Plausibility checks</h3>${a.plausibility_warnings.map(w=>`<div class="${w.includes('No major')?'ok':'warning'}">${html(w)}</div>`).join('')}`;
}

function renderAnalysis(){
  const an = workspace.analysis;
  $('analysis').innerHTML = `
    <span class="eyebrow">Step 3</span><h2>Analysis workbench</h2>
    <p class="note">The app auto-selected a possible comparison, but you can override it. All p-values are exploratory and require verification before submission.</p>
    <div class="mini-form">
      <div><label>Exposure / group variable</label><select id="groupVar">${an.group_candidates.map(c=>`<option ${c===an.auto_group_variable?'selected':''}>${html(c)}</option>`).join('')}</select></div>
      <div><label>Numeric outcome variable</label><select id="outcomeVar">${an.numeric_candidates.map(c=>`<option ${c===an.auto_outcome_variable?'selected':''}>${html(c)}</option>`).join('')}</select></div>
      <button class="primary" onclick="runComparison()">Run selected comparison</button>
    </div>
    <div id="comparisonBox">${renderComparison(an.primary_comparison)}</div>
    <h3>Exploratory Table 1</h3><div class="scroll">${table(an.table1.slice(0,80))}</div>
    <h3>Paired pre/post screening</h3>${an.paired_screening.length ? `<div class="scroll">${table(an.paired_screening)}</div>` : '<div class="warning">No obvious paired baseline/follow-up columns detected from names. Rename columns clearly or verify manually.</div>'}
    <h3>Regression feasibility</h3><div class="${an.regression_feasibility.status==='Potentially feasible'?'ok':'warning'}">${html(an.regression_feasibility.status)} — ${html(an.regression_feasibility.reason)}</div>`;
}

window.runComparison = async function(){
  const group_variable = $('groupVar').value, outcome_variable = $('outcomeVar').value;
  const res = await fetch(`/api/compare/${workspace.id}`, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({group_variable,outcome_variable})});
  const comp = await res.json();
  workspace.analysis.primary_comparison = comp;
  workspace.analysis.auto_group_variable = group_variable;
  workspace.analysis.auto_outcome_variable = outcome_variable;
  $('comparisonBox').innerHTML = renderComparison(comp);
  renderManuscript(); renderExport();
}

function renderComparison(comp){
  if(!comp) return '<div class="warning">No valid primary comparison was auto-generated. Select variables above and run comparison.</div>';
  if(comp.error) return `<div class="warning">${html(comp.error)}</div>`;
  return `<div class="card"><strong>Primary comparison</strong><p>${html(comp.outcome_variable)} by ${html(comp.group_variable)} · ${html(comp.test)} · exploratory p=${html(comp.p_value_exploratory)}</p><div class="scroll">${table(comp.summary)}</div><div class="figure">${comp.figure_svg || ''}</div><p class="note">${html(comp.warning)}</p></div>`;
}

function renderManuscript(){
  const r = workspace.research;
  $('manuscript').innerHTML = `
    <span class="eyebrow">Step 4</span><h2>Manuscript strategy and builder</h2>
    <p class="note">Generate only after reviewing the analysis. The app will not invent references and will label missing information.</p>
    <div class="grid two">
      <div class="card"><strong>App-recommended direction</strong><p>${html(r.recommended_primary_question)}</p><button class="secondary" onclick="generateManuscript('')">Generate from recommended direction</button></div>
      <div class="card"><strong>User-selected direction</strong><textarea id="customDirection" rows="5" placeholder="Edit or type the exact manuscript direction you want..."></textarea><button class="secondary" onclick="generateManuscript(document.getElementById('customDirection').value)">Generate from my direction</button></div>
    </div>
    <div id="manuscriptDraft">${workspace.manuscript ? renderManuscriptDraft(workspace.manuscript) : '<div class="warning">No manuscript generated yet.</div>'}</div>`;
}
window.generateManuscript = async function(direction){
  const box = $('manuscriptDraft'); box.innerHTML = '<div class="warning">Generating data-grounded manuscript package...</div>';
  const res = await fetch(`/api/manuscript/${workspace.id}`, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({direction, use_ai:true})});
  if(!res.ok){ box.innerHTML='<div class="warning">Manuscript generation failed.</div>'; return; }
  workspace.manuscript = await res.json();
  box.innerHTML = renderManuscriptDraft(workspace.manuscript); renderExport();
}
function renderManuscriptDraft(ms){
  return `<h3>Generated manuscript package</h3>${Object.entries(ms.sections).map(([h,t])=>`<div class="manuscript-section"><strong>${html(h)}</strong>\n\n${html(t)}</div>`).join('')}`;
}

function renderExport(){
  $('export').innerHTML = `
    <span class="eyebrow">Step 5</span><h2>Export</h2>
    <div class="cards">
      <div class="card"><strong>Workspace JSON</strong><span>Full audit, analysis, and generated content.</span><div class="actions"><a class="secondary" href="/api/download/workspace/${workspace.id}">Download workspace</a></div></div>
      <div class="card"><strong>Word manuscript</strong><span>DOCX package with manuscript sections and computed tables.</span><div class="actions">${workspace.manuscript ? `<a class="secondary" href="/api/download/manuscript/${workspace.id}">Download DOCX manuscript</a>` : '<button class="secondary" onclick="scrollToSection(\'manuscript\')">Generate manuscript first</button>'}</div></div>
      <div class="card"><strong>Unresolved checklist</strong><span>${workspace.research.unresolved_items.length} items require user confirmation before submission.</span></div>
    </div>`;
}

function table(rows){
  if(!rows || !rows.length) return '<p class="note">No rows available.</p>';
  const cols = [...new Set(rows.flatMap(r => Object.keys(r).filter(k=>k!=='figure_svg')))];
  return `<table><thead><tr>${cols.map(c=>`<th>${html(c)}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${cols.map(c=>`<td>${html(r[c] ?? '')}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
}

$('resetBtn').onclick = () => { workspace=null; selectedFiles=[]; $('files').value=''; $('fileList').innerHTML=''; $('workspace').classList.add('hidden'); $('nav').classList.add('hidden'); $('projectStatus').innerHTML='<span class="eyebrow">Project status</span><strong>Waiting for study intake</strong><p>Upload data and notes to begin.</p>'; window.scrollTo({top:0,behavior:'smooth'}); };
