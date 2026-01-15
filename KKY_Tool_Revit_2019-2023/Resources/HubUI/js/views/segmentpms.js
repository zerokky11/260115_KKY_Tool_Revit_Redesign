import { clear, div, toast, setBusy, showExcelSavedDialog, chooseExcelMode } from '../core/dom.js';
import { createRvtTable, renderRvtRows, getRvtName } from './rvtTable.js';
import { ProgressDialog } from '../core/progress.js';
import { post, onHost } from '../core/bridge.js';

const LS_RVT_LIST = 'kky_segmentpms_rvt_list';
const SUGGEST_SCORE_THRESHOLD = 70;
let progressHideTimer = null;
const PROGRESS_STAGE_WEIGHT = {
  open: 0,
  start: 0,
  extract: 0.33,
  route: 0.66,
  save: 0.75,
  excel_init: 0.75,
  excel_write: 0.9,
  excel_save: 0.95,
  autofit: 0.98,
  finish: 1,
  done: 1,
  error: 1
};
const PROGRESS_STAGE_TITLE = {
  open: 'RVT 준비 중',
  start: 'RVT 준비 중',
  extract: 'Segment 추출 중',
  route: 'PMS 매핑 적용 중',
  save: '결과 저장 중',
  excel_init: '엑셀 준비 중',
  excel_write: '엑셀 작성 중',
  excel_save: '엑셀 내보내기 중',
  autofit: 'AutoFit 적용 중',
  finish: '검토 마무리 중',
  done: '검토 완료',
  error: '오류 발생'
};
const PROGRESS_STAGE_DETAIL = {
  open: 'Revit 파일을 여는 중입니다.',
  start: 'Revit 파일을 여는 중입니다.',
  extract: 'Segment 데이터를 추출하고 있습니다.',
  route: 'PMS 룰과 매핑을 준비하고 있습니다.',
  save: '결과를 저장하고 있습니다.',
  excel_init: '엑셀 워크북을 준비하고 있습니다.',
  excel_write: '엑셀 데이터를 작성하고 있습니다.',
  excel_save: '파일을 저장하는 중입니다.',
  autofit: 'AutoFit을 적용하고 있습니다.',
  finish: '처리가 곧 완료됩니다.',
  done: '모든 파일 처리가 완료되었습니다.',
  error: '진행 중 오류가 발생했습니다.'
};
let progressPrevPct = 0;

function loadRvtList() {
  try {
    const arr = JSON.parse(localStorage.getItem(LS_RVT_LIST) || '[]');
    if (Array.isArray(arr)) return arr;
  } catch { }
  return [];
}

function saveRvtList(list) {
  localStorage.setItem(LS_RVT_LIST, JSON.stringify(list || []));
}

export function renderSegmentPms(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);
  const top = document.querySelector('#topbar-root .topbar') || document.querySelector('.topbar'); if (top) top.classList.add('hub-topbar');

  const state = {
    rvtList: loadRvtList(),
    rvtChecked: new Set(loadRvtList()),
    extractLoaded: false,
    extractSummary: '',
    extractPath: '',
    pmsLoaded: false,
    pmsOpts: [],
    groups: [],
    suggestions: new Map(), // key: groupKey -> {cls, segment}
    selections: new Map(), // groupKey -> {cls, segment, source}
    results: null,
    busy: false,
    progressTimer: null
  };

  const page = div('feature-shell segmentpms-page');
  const header = div('feature-header');
  const heading = div('feature-heading');
  heading.innerHTML = `<span class="feature-kicker">PipeType - PMS</span><h2 class="feature-title">Segment 매핑/검증</h2><p class="feature-sub">추출(Excel)과 PMS를 분리하여 그룹 단위 매핑 후 비교합니다.</p>`;
  header.append(heading);
  page.append(header);

  /* Extract section */
  const extractSection = div('section segmentpms-extract');
  const exHeader = document.createElement('div'); exHeader.className = 'section-header';
  exHeader.innerHTML = '<h3>1단계: 추출 (RVT → Excel)</h3>';
  const exActions = div('segmentpms-actions-row');
  const btnAddRvt = cardBtn('RVT 파일 추가', () => post('segmentpms:rvt-pick-files', {}));
  const btnAddFolder = cardBtn('폴더 선택', () => post('segmentpms:rvt-pick-folder', {}));
  const btnRemoveSel = cardBtn('선택 제거', removeCheckedRvt);
  const btnClearAll = cardBtn('등록 목록 비우기', () => { state.rvtList = []; state.rvtChecked.clear(); persistRvt(); renderRvtList(); updateButtons(); });
  const btnExtract = cardBtn('추출 시작', onExtract);
  const btnSaveExtract = cardBtn('엑셀 내보내기', () => chooseExcelMode((mode) => post('segmentpms:save-extract', { excelMode: mode || 'fast' })));
  exActions.append(btnAddRvt, btnAddFolder, btnRemoveSel, btnClearAll, btnExtract, btnSaveExtract);
  exHeader.append(exActions);

  const { table: rvtTable, tbody: rvtBody, master: rvtMaster } = createRvtTable();
  const rvtBox = div('segmentpms-rvtlist');
  rvtBox.append(rvtTable);
  const extractInfo = div('segmentpms-summary'); extractInfo.textContent = '추출 상태: 미실행';
  const extractLoadRow = div('segmentpms-actions-row');
  const btnLoadExtract = cardBtn('추출 결과 불러오기', () => post('segmentpms:load-extract', {}));
  extractLoadRow.append(btnLoadExtract);
  extractSection.append(exHeader, rvtBox, extractLoadRow, extractInfo);
  page.append(extractSection);

  /* Check section */
  const checkSection = div('section segmentpms-check');
  const chHeader = document.createElement('div'); chHeader.className = 'section-header';
  chHeader.innerHTML = '<h3>2단계: 검토 (추출 Excel + PMS)</h3>';
  const chActions = div('segmentpms-actions-row');
  const btnRegisterPms = cardBtn('PMS 등록/업데이트', () => { setBusy(true, 'PMS 불러오는 중'); state.busy = true; updateButtons(); post('segmentpms:register-pms', {}); });
  const btnTemplate = cardBtn('PMS 양식 추출하기', () => { setBusy(true, 'PMS 양식 저장 중'); state.busy = true; updateButtons(); post('segmentpms:pms-template', {}); });
  const btnPrepare = cardBtn('매핑 준비', onPrepareMapping);
  const btnRun = cardBtn('검토 시작', onRun);
  const btnSave = cardBtn('엑셀 내보내기', () => {
    if (!state.results) { toast('저장할 결과가 없습니다.', 'err'); return; }
    chooseExcelMode((mode) => post('segmentpms:save-result', { excelMode: mode || 'fast' }));
  });
  chActions.append(btnLoadExtract, btnRegisterPms, btnTemplate, btnPrepare, btnRun, btnSave);
  chHeader.append(chActions);
  const pmsGuide = div('segmentpms-summary');
  pmsGuide.textContent = "처음 사용자는 ‘PMS 양식 추출하기’로 샘플을 내려받아 동일 형식으로 작성하세요.";
  checkSection.append(chHeader, pmsGuide);

  const groupTable = document.createElement('table'); groupTable.className = 'segmentpms-table';
  groupTable.innerHTML = '<thead><tr><th>Revit Segment 그룹</th><th>사용처</th><th>PMS Segment</th><th>추천</th></tr></thead><tbody></tbody>';
  const groupBody = groupTable.querySelector('tbody');
  checkSection.append(groupTable);

  const resInfo = div('segmentpms-summary'); resInfo.textContent = '결과 없음';
  checkSection.append(resInfo);
  page.append(checkSection);

  target.append(page);

  renderRvtList();
  updateButtons();
  onHost(handleHost);

  function persistRvt() { saveRvtList(state.rvtList); }

  function renderRvtList() {
    const allChecked = state.rvtList.length > 0 && state.rvtList.every(f => state.rvtChecked.has(f));
    rvtMaster.checked = allChecked;
    rvtMaster.disabled = state.rvtList.length === 0;
    rvtMaster.onchange = () => {
      if (rvtMaster.checked) state.rvtChecked = new Set(state.rvtList);
      else state.rvtChecked.clear();
      renderRvtList();
      updateButtons();
    };
    const rows = state.rvtList.map((p, idx) => ({
      checked: state.rvtChecked.has(p),
      index: idx + 1,
      name: getRvtName(p),
      path: p,
      title: p,
      onToggle: (checked) => {
        if (checked) state.rvtChecked.add(p);
        else state.rvtChecked.delete(p);
        updateButtons();
      }
    }));
    renderRvtRows(rvtBody, rows);
    updateButtons();
  }

  function removeCheckedRvt() {
    if (!state.rvtChecked.size) { toast('제거할 항목을 선택하세요.', 'err'); return; }
    state.rvtList = state.rvtList.filter(p => !state.rvtChecked.has(p));
    state.rvtChecked.clear(); persistRvt(); renderRvtList(); updateButtons();
  }

  function onExtract() {
    const targets = state.rvtList.filter(p => state.rvtChecked.has(p));
    if (!targets.length) { toast('추출할 RVT를 선택하세요.', 'err'); return; }
    state.busy = true; updateButtons();
    post('segmentpms:extract', { files: targets });
  }

  function onPrepareMapping() {
    if (!state.extractLoaded) { toast('추출 결과를 먼저 불러오세요.', 'err'); return; }
    setBusy(true, '매핑 준비'); state.busy = true; updateButtons();
    post('segmentpms:prepare-mapping', {});
  }

  function buildGroupTable() {
    groupBody.innerHTML = '';
    state.selections.clear();
    if (!state.groups.length) {
      const tr = document.createElement('tr'); const td1 = document.createElement('td'); td1.colSpan = 4; td1.textContent = '추출 결과와 PMS를 불러와 매핑을 준비하세요.'; tr.append(td1); groupBody.append(tr); updateButtons(); return;
    }
    state.groups.forEach(g => {
      const tr = document.createElement('tr');
      tr.append(td(g.displayKey || g.groupKey));
      tr.append(td(g.usageSummary || ''));

      const pmsSel = document.createElement('select');
      fillPmsOptions(pmsSel);
      const suggLabel = document.createElement('small'); suggLabel.className = 'segmentpms-suggest';
      const localSuggestion = suggestPms(g.displayKey || g.groupKey, state.pmsOpts);
      const backendSuggestion = normalizeSuggestion(state.suggestions.get(g.groupKey));
      const groupSuggestion = normalizeSuggestion(g.suggestedSegmentKey ? { cls: g.suggestedClass, segment: g.suggestedSegmentKey, score: g.score || g.Score } : null);
      const applySuggestion = () => {
        const sug = pickBestSuggestion([backendSuggestion, groupSuggestion, localSuggestion]);
        if (sug && sug.segment && (sug.score || 0) >= SUGGEST_SCORE_THRESHOLD) {
          pmsSel.value = `${sug.cls}|||${sug.segment}`;
          suggLabel.textContent = '추천 적용';
          commitSelection(g.groupKey, pmsSel.value, 'Suggest');
        } else {
          suggLabel.textContent = '';
          commitSelection(g.groupKey, pmsSel.value, 'Manual');
        }
      };
      applySuggestion();

      pmsSel.onchange = () => {
        commitSelection(g.groupKey, pmsSel.value, 'Manual');
        suggLabel.textContent = '';
      };

      const tdPms = document.createElement('td'); tdPms.append(pmsSel);
      const tdSug = document.createElement('td'); tdSug.append(suggLabel);
      tr.append(tdPms, tdSug);
      groupBody.append(tr);
    });
    updateButtons();
  }

  function commitSelection(groupKey, pmsVal, source) {
    const val = String(pmsVal || '');
    const parts = val.split('|||');
    const cls = parts[0] || '';
    const seg = parts[1] || '';
    state.selections.set(groupKey, { groupKey, cls, segment: seg, source: source || 'Manual' });
  }

  function fillPmsOptions(sel) {
    const oldVal = sel.value;
    sel.innerHTML = '';
    sel.append(new Option('(선택)', ''));
    const added = new Set();
    state.pmsOpts.forEach(o => {
      const key = `${o.cls}|||${o.segment}`;
      if (added.has(key)) return;
      added.add(key);
      sel.append(new Option(o.label, key));
    });
    sel.value = oldVal;
  }

  function onRun() {
    if (!state.extractLoaded) { toast('추출 데이터를 먼저 불러오세요.', 'err'); return; }
    if (!state.pmsLoaded) { toast('PMS를 등록하세요.', 'err'); return; }
    const groups = [...state.selections.values()];
    state.results = null;
    setBusy(true, '검토 실행'); state.busy = true; updateButtons();
    post('segmentpms:run', { groups });
  }

  function paintResults(payload) {
    state.results = payload || { hasResult: true };
    const total = payload?.totalCount ?? (Array.isArray(payload?.compare) ? payload.compare.length : 0);
    resInfo.textContent = total > 0 ? `총 ${total}건의 결과가 준비되었습니다. 엑셀 내보내기 후 확인하세요.` : '결과 없음';
    toast('검토가 완료되었습니다. 엑셀로 내보내세요.', 'ok');
  }

  function updateButtons() {
    btnRemoveSel.disabled = state.busy || state.rvtChecked.size === 0;
    btnClearAll.disabled = state.busy || state.rvtList.length === 0;
    btnExtract.disabled = state.busy || state.rvtChecked.size === 0;
    btnSaveExtract.disabled = state.busy || !state.extractLoaded;
    btnLoadExtract.disabled = state.busy;
    btnRegisterPms.disabled = state.busy;
    btnTemplate.disabled = state.busy;
    btnPrepare.disabled = state.busy || !state.extractLoaded;
    btnRun.disabled = state.busy || !state.extractLoaded || !state.pmsLoaded;
    btnSave.disabled = state.busy || !state.results;
  }

  function handleHost(msg) {
    if (!msg || !msg.ev) return;
    switch (msg.ev) {
      case 'segmentpms:rvt-picked-files':
      case 'segmentpms:rvt-picked-folder': {
        const files = Array.isArray(msg.payload?.paths) ? msg.payload.paths : [];
        files.forEach(f => { if (!state.rvtList.some(x => x.toLowerCase() === String(f).toLowerCase())) state.rvtList.push(f); });
        state.rvtChecked = new Set(files);
        persistRvt(); renderRvtList(); updateButtons();
        break;
      }
      case 'segmentpms:extract-saved':
        setBusy(false); state.busy = false;
        state.extractLoaded = true;
        state.extractSummary = msg.payload?.summary || '';
        state.extractPath = msg.payload?.path || '';
        state.results = null;
        extractInfo.textContent = `추출 완료: ${state.extractSummary} (${state.extractPath})`;
        toast('추출을 완료했습니다.', 'ok');
        post('segmentpms:prepare-mapping', {});
        updateButtons();
        break;
      case 'segmentpms:extract-loaded':
        setBusy(false); state.busy = false;
        state.extractLoaded = true;
        state.extractSummary = msg.payload?.summary || '';
        state.extractPath = msg.payload?.path || '';
        state.results = null;
        extractInfo.textContent = `추출 로드: ${state.extractSummary} (${state.extractPath})`;
        state.groups = msg.payload?.groups || [];
        state.pmsOpts = msg.payload?.pms || state.pmsOpts;
        state.suggestions = buildSuggestionMap(msg.payload?.suggestions || []);
        buildGroupTable();
        updateButtons();
        break;
      case 'segmentpms:mapping-ready':
        setBusy(false); state.busy = false;
        state.groups = msg.payload?.groups || [];
        state.pmsOpts = msg.payload?.pms || state.pmsOpts;
        state.suggestions = buildSuggestionMap(msg.payload?.suggestions || []);
        buildGroupTable();
        updateButtons();
        break;
      case 'segmentpms:pms-registered':
        setBusy(false); state.busy = false;
        state.pmsLoaded = true;
        state.pmsOpts = msg.payload?.options || [];
        state.suggestions = buildSuggestionMap(msg.payload?.suggestions || []);
        fillPmsOptions(document.createElement('select'));
        if (state.extractLoaded) post('segmentpms:prepare-mapping', {});
        toast('PMS를 등록했습니다.', 'ok');
        updateButtons();
        break;
      case 'segmentpms:pms-template-saved':
        setBusy(false); state.busy = false;
        showExcelSavedDialog('PMS 양식을 저장했습니다.', msg.payload?.path, (p) => {
          const target = p || msg.payload?.path;
          if (!target) { toast('열 수 있는 경로가 없습니다.', 'err'); return; }
          post('excel:open', { path: target });
        });
        updateButtons();
        break;
      case 'segmentpms:result':
        setBusy(false); state.busy = false;
        paintResults(msg.payload || {});
        updateButtons();
        break;
      case 'segmentpms:progress':
        handleProgress(msg.payload || {});
        break;
      case 'segmentpms:saved':
        showExcelSavedDialog('결과를 저장했습니다.', msg.payload?.path, (p) => {
          const target = p || msg.payload?.path;
          if (!target) { toast('열 수 있는 경로가 없습니다.', 'err'); return; }
          post('excel:open', { path: target });
        });
        break;
      case 'segmentpms:error':
        setBusy(false); state.busy = false;
        toast(msg.payload?.message || '오류가 발생했습니다.', 'err');
        updateButtons();
        break;
      default: break;
    }
  }
}

function buildSuggestionMap(list) {
  const map = new Map();
  if (!Array.isArray(list)) return map;
  list.forEach(s => {
    const key = s.groupKey || s.GroupKey || s.segmentKey || s.SegmentKey || s.file;
    if (!key) return;
    const rawScore = Number(s.score ?? s.Score ?? 0);
    const score = Number.isFinite(rawScore) ? rawScore : 0;
    map.set(String(key), { pmsClass: s.pmsClass || s.PmsClass, pmsSegmentKey: s.pmsSegmentKey || s.PmsSegmentKey, score });
  });
  return map;
}

function handleProgress(payload) {
  if (progressHideTimer) { clearTimeout(progressHideTimer); progressHideTimer = null; }
  if (!payload) { ProgressDialog.hide(); progressPrevPct = 0; return; }

  const stage = normalizeStage(payload.stage || payload.phase);
  const total = Number(payload.total ?? payload.fileTotal) || 0;
  const index = Number(payload.index ?? payload.fileIndex ?? payload.current) || 0;
  const currentRows = Number(payload.current) || index;
  const percent = computeWeightedPercent(stage, total, index, payload.percent, currentRows);

  const file = payload.file || payload.fileName || '';
  const msg = payload.message || '';
  const title = PROGRESS_STAGE_TITLE[stage] || 'Segment/PMS 진행 중';
  const subtitle = buildProgressDetail(stage, msg, file);
  const meta = buildProgressMeta(total, index, file);

  ProgressDialog.show('Segment 매핑/검증', title);
  ProgressDialog.update(percent, subtitle, meta);

  if (stage === 'finish' || stage === 'done') {
    progressHideTimer = setTimeout(() => { ProgressDialog.hide(); progressPrevPct = 0; }, 600);
  } else if (stage === 'error') {
    ProgressDialog.hide();
    progressPrevPct = 0;
  }
}

function normalizeSuggestion(sug) {
  if (!sug) return null;
  const cls = sug.cls || sug.pmsClass || sug.suggestedClass || '';
  const segment = sug.segment || sug.pmsSegmentKey || sug.suggestedSegmentKey || '';
  const rawScore = Number(sug.score ?? sug.Score ?? 0);
  const score = Number.isFinite(rawScore) ? rawScore : 0;
  if (!segment) return null;
  return { cls, segment, score };
}

function pickBestSuggestion(candidates) {
  return candidates
    .filter(Boolean)
    .reduce((best, cur) => {
      const curScore = Number.isFinite(cur.score) ? cur.score : 0;
      if (!best || curScore > best.score) return { ...cur, score: curScore };
      return best;
    }, null);
}

function suggestPms(revitSegmentKey, pmsOpts) {
  const tokens = tokenize(revitSegmentKey);
  if (!tokens.length || !Array.isArray(pmsOpts)) return null;
  let best = null;
  pmsOpts.forEach(opt => {
    const optTokens = tokenize(opt.segment);
    if (!optTokens.length) return;
    const matchCount = countOrderedMatches(tokens, optTokens);
    const lenDiff = Math.abs(optTokens.length - tokens.length);
    if (!best || matchCount > best.score || (matchCount === best.score && lenDiff < best.lenDiff)) {
      best = { cls: opt.cls, segment: opt.segment, score: matchCount, lenDiff };
    }
  });
  if (!best || best.score === 0) return null;
  return { cls: best.cls, segment: best.segment, score: best.score || 0 };
}

function tokenize(text) {
  if (!text) return [];
  return String(text)
    .toUpperCase()
    .replace(/[^\p{L}\p{N}]+/gu, ' ')
    .split(' ')
    .map(t => t.trim())
    .filter(Boolean);
}

function countOrderedMatches(targetTokens, candidateTokens) {
  let idx = 0; let count = 0;
  targetTokens.forEach(t => {
    while (idx < candidateTokens.length) {
      if (t === candidateTokens[idx]) {
        count += 1;
        idx += 1;
        return;
      }
      idx += 1;
    }
  });
  return count;
}

function cardBtn(label, onclick) {
  const b = document.createElement('button');
  b.type = 'button'; b.className = 'btn card-btn'; b.textContent = label; b.onclick = onclick; return b;
}

function td(v) { const t = document.createElement('td'); t.textContent = v == null ? '' : v; return t; }

function normalizeStage(stage) {
  return String(stage || '').toLowerCase();
}

function computeWeightedPercent(stage, total, index, incomingPct, currentRows = 0) {
  const clamp = (n) => Math.max(0, Math.min(100, n));
  const weight = PROGRESS_STAGE_WEIGHT[stage] ?? 0;
  const safeTotal = Math.max(1, total || 1);
  if (stage === 'open' || stage === 'start') progressPrevPct = 0;
  const cursor = Math.max(0, stage.startsWith('excel') || stage === 'autofit' ? currentRows : index);
  const basePct = clamp(weight * 100);
  const ratio = safeTotal > 0 ? Math.min(1, Math.max(0, cursor / safeTotal)) : 0;
  const weightedPct = clamp(basePct + (100 - basePct) * ratio);
  const parsedPct = Number(incomingPct);
  const providedPct = Number.isFinite(parsedPct) ? clamp(parsedPct) : 0;
  const pct = Math.max(progressPrevPct, providedPct, weightedPct);
  progressPrevPct = pct;
  return pct;
}

function buildProgressDetail(stage, message, file) {
  const parts = [];
  const fallback = PROGRESS_STAGE_DETAIL[stage] || '데이터를 처리하고 있습니다.';
  if (message) parts.push(message);
  if (parts.length === 0) parts.push(fallback);
  if (file) parts.push(file);
  return parts.join(' · ');
}

function buildProgressMeta(total, index, file) {
  const bits = [];
  const totalNum = Number(total) || 0;
  const idxNum = Number(index) || 0;
  if (totalNum > 0) bits.push(`${Math.max(idxNum, 0)}/${totalNum}`);
  if (file) bits.push(file);
  return bits.join(' · ');
}
