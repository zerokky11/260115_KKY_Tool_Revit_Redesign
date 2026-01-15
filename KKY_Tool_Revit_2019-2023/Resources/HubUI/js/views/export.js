import { clear, div, tdText, toast, showExcelSavedDialog, chooseExcelMode } from '../core/dom.js';
import { createRvtTable, renderRvtRows, getRvtName } from './rvtTable.js';
import { ProgressDialog } from '../core/progress.js';
import { post, onHost } from '../core/bridge.js';

const state = { files: [], rowsRaw: [], folder: '', unit: 'ft' };
const FT_TO_M = 0.3048;
const FT_TO_MM = 304.8;
const HEADERS = [
  { key: 'ProjectPoint_E(mm)', label: 'E/W', group: 'project' },
  { key: 'ProjectPoint_N(mm)', label: 'N/S', group: 'project' },
  { key: 'ProjectPoint_Z(mm)', label: 'Elev', group: 'project' },
  { key: 'SurveyPoint_E(mm)', label: 'E/W', group: 'survey' },
  { key: 'SurveyPoint_N(mm)', label: 'N/S', group: 'survey' },
  { key: 'SurveyPoint_Z(mm)', label: 'Elev', group: 'survey' }
];

export function renderExport(root) {
    const target = root || document.getElementById('view-root') || document.getElementById('app'); clear(target);
    const top = document.querySelector('#topbar-root .topbar') || document.querySelector('.topbar'); if (top) top.classList.add('hub-topbar');

    const page = div('feature-shell');

    const header = div('feature-header');
    const heading = div('feature-heading');
    heading.innerHTML = `
      <span class="feature-kicker">Export Points with Angle</span>
      <h2 class="feature-title">좌표/북각 추출</h2>
      <p class="feature-sub">RVT 폴더를 선택해 포인트/북각을 미리보기 후 Excel로 저장합니다.</p>`;

    const pick = cardBtn('폴더 선택', () => post('export:browse-folder', {}));
    pick.id = 'btnExPick';
    const addFilesBtn = cardBtn('RVT 파일 추가', () => post('export:add-rvt-files', {}));
    addFilesBtn.id = 'btnExAddFiles';
    const preview = cardBtn('추출 시작', () => {
      const targets = selectedFilePaths();
      if (!targets.length) { toast('선택된 RVT가 없습니다.', 'warn'); return; }
      setWorking(true);
      startProgress('COLLECT', '미리보기 준비 중…', targets.length);
      post('export:preview', { files: targets, unit: state.unit });
    });
    preview.id = 'btnExPreview'; preview.disabled = true; preview.classList.add('btn-primary');
    const save = cardBtn('엑셀 내보내기', () => {
      chooseExcelMode((mode) => {
        const payload = { rows: convertRowsForSave(), unit: state.unit, files: selectedFilePaths(), excelMode: mode || 'fast' };
        setWorking(true);
        startProgress('EXCEL', '엑셀 내보내기 준비 중…', state.rowsRaw.length);
        post('export:save-excel', payload);
      });
    });
    save.id = 'btnExSave'; save.disabled = true; save.classList.add('btn-outline');
    header.append(heading);
    page.append(header);

    const wrap = div('kkyt-stack');

    const left = div('kkyt-left feature-results-panel segmentpms-extract');
    const lbar = div('segmentpms-actions-row');
    const removeBtn = cardBtn('선택 제거', () => {
      const checked = state.files.filter(f => f.checked);
      if (!checked.length) { toast('제거할 RVT를 선택하세요.', 'warn'); return; }
      const removeSet = new Set(checked.map(f => f.path.toLowerCase()));
      state.files = state.files.filter(f => !removeSet.has(f.path.toLowerCase()));
      renderFiles();
      repaintRows();
      syncSaveState();
    });
    removeBtn.id = 'btnExRemove'; removeBtn.disabled = true;
    const clearBtn = cardBtn('등록 목록 비우기', () => {
      state.files = [];
      state.folder = '';
      state.rowsRaw = [];
      renderFiles();
      repaintRows();
      syncSaveState();
    });
    const unitToggle = buildUnitToggle();
    lbar.append(addFilesBtn, pick, removeBtn, clearBtn, preview, save);
    const listWrap = div('segmentpms-rvtlist');
    const { table: tblWrap, tbody: filesBody, master: filesMaster } = createRvtTable();
    listWrap.append(tblWrap);
    const info = div('segmentpms-summary'); info.textContent = '파일 0개';
    left.append(lbar, listWrap, unitToggle, info);

    const right = div('kkyt-right feature-results-panel');
    const tbl = document.createElement('table'); tbl.className = 'kkyt-table';
    const thead = document.createElement('thead');
    paintHead(thead);
    const tbody = document.createElement('tbody'); tbl.append(thead, tbody);

    right.append(tbl);
    wrap.append(left, right);
    page.append(wrap);
    target.append(page);

    // === 파일 선택 응답 ===
    onHost('export:files', ({ files, folder }) => {
        const list = Array.isArray(files) ? files : [];
        const root = folder || commonDir(list);
        state.folder = root || '';
        state.files = toFileItems(list, state.folder);
        renderFiles();
    });

    onHost('export:rvt-files', ({ files }) => {
        const list = Array.isArray(files) ? files : [];
        if (!list.length) return;
        if (!state.folder) state.folder = commonDir(list) || state.folder || '';
        const merged = [...state.files, ...toFileItems(list, state.folder)];
        state.files = dedupFiles(merged);
        renderFiles();
    });

    // === 미리보기 결과 ===
    onHost('export:previewed', ({ rows }) => {
        finishWorking();
        ProgressDialog.hide();
        state.rowsRaw = Array.isArray(rows) ? rows : [];
        repaintRows();
        syncSaveState();
        toast(`미리보기 ${state.rowsRaw.length}행`, 'ok');
    });

    // === 저장 결과 ===
    onHost('export:saved', ({ path }) => {
        const p = path || '';
        finishWorking();
        ProgressDialog.hide();
        if (p) {
            requestAnimationFrame(() => {
              showExcelSavedDialog('엑셀 파일을 내보냈습니다.', p, (fp) => {
                  if (fp) post('excel:open', { path: fp });
              });
            });
        } else {
            toast('엑셀 파일을 내보냈습니다.', 'ok', 2600);
        }
    });

    // === 에러 공통 처리(중요) ===
    onHost('revit:error', ({ message }) => { handleError(message || 'Revit 오류가 발생했습니다.'); });
    onHost('host:error', ({ message }) => { handleError(message || '호스트 오류가 발생했습니다.'); });

    onHost('export:progress', handleProgress);

    function renderFiles() {
        while (filesBody.firstChild) filesBody.removeChild(filesBody.firstChild);
        const allChecked = state.files.length && state.files.every(f => f.checked);
        filesMaster.checked = allChecked;
        filesMaster.disabled = state.files.length === 0;
        filesMaster.onchange = () => { state.files = state.files.map(f => ({ ...f, checked: filesMaster.checked })); renderFiles(); };

        const rows = state.files.map((f, idx) => {
          const path = f.path || f.rel || f.name || '';
          const name = f.name || getRvtName(path);
          return {
            checked: !!f.checked,
            index: idx + 1,
            name,
            path,
            title: path,
            onToggle: (checked) => {
              state.files[idx].checked = checked;
              updateSelectionSummary();
              syncPreviewState();
              syncRemoveState();
            }
          };
        });
        renderRvtRows(filesBody, rows);
        updateSelectionSummary();
        syncPreviewState();
        syncRemoveState();
    }

    function updateSelectionSummary() {
        const picked = state.files.filter(f => f.checked).length;
        if (!state.files.length) {
          info.textContent = '파일 0개';
        } else {
          info.textContent = `파일 ${state.files.length}개 중 ${picked}개 선택`;
        }
    }

    function syncRemoveState() {
      const btn = document.getElementById('btnExRemove');
      if (btn) btn.disabled = state.files.every(f => !f.checked);
      if (clearBtn) clearBtn.disabled = state.files.length === 0;
    }
}

function cardBtn(text, onClick) {
    const b = document.createElement('button');
    b.textContent = text;
    b.className = 'btn card-btn';
    if (typeof onClick === 'function') b.addEventListener('click', onClick);
    return b;
}

function buildUnitToggle() {
    const wrap = document.createElement('div');
    wrap.className = 'unit-toggle';
    wrap.setAttribute('role', 'radiogroup');
    wrap.innerHTML = `
      <label><input type="radio" name="unit" value="ft" checked> Decimal Feet</label>
      <label><input type="radio" name="unit" value="m"> Meters (m)</label>
      <label><input type="radio" name="unit" value="mm"> Millimeters (mm)</label>`;
    wrap.querySelectorAll('input[type="radio"]').forEach(r => {
      r.checked = r.value === state.unit;
      r.onchange = () => { state.unit = r.value; paintHead(); repaintRows(); };
    });
    return wrap;
}

function selectedFilePaths() {
    return state.files.filter(f => f.checked).map(f => f.path);
}

function syncPreviewState() {
    const hasChecked = selectedFilePaths().length > 0;
    const previewBtn = document.getElementById('btnExPreview');
    if (previewBtn) previewBtn.disabled = !hasChecked || isWorking;
}

function commonDir(list) {
    if (!list || !list.length) return '';
    const norm = list.map(p => String(p || '').replace(/\\/g, '/'));
    const parts = norm[0].split('/'); parts.pop();
    let prefix = parts.join('/');
    for (const p of norm.slice(1)) {
      while (prefix && !p.startsWith(prefix)) {
        prefix = prefix.split('/').slice(0, -1).join('/');
      }
    }
    return prefix;
}

function relPath(path, root) {
    const normRoot = String(root || '').replace(/[\\/]+$/, '').replace(/\\/g, '/');
    const normPath = String(path || '').replace(/\\/g, '/');
    if (!normRoot) return normPath;
    if (normPath.startsWith(normRoot)) return normPath.slice(normRoot.length + 1);
    return normPath;
}

function paintHead(target) {
    const unitLabel = state.unit === 'm' ? '(m)' : (state.unit === 'mm' ? '(mm)' : '(ft)');
    const project = HEADERS.filter(h => h.group === 'project');
    const survey = HEADERS.filter(h => h.group === 'survey');
    const head = `
      <tr>
        <th rowspan="2">File</th>
        <th class="group" colspan="${project.length}">Project Point ${unitLabel}</th>
        <th class="group" colspan="${survey.length}">Survey Point ${unitLabel}</th>
        <th rowspan="2">Angle (True North)</th>
      </tr>
      <tr>
        ${project.map(h => `<th>${h.label}</th>`).join('')}
        ${survey.map(h => `<th>${h.label}</th>`).join('')}
      </tr>`;
    const thead = target || document.querySelector('.kkyt-table thead');
    if (thead) thead.innerHTML = head;
}

function formatCoord(v) {
    const n = Number(v);
    if (!Number.isFinite(n)) return v ?? '';
    const scaled = state.unit === 'm' ? n * FT_TO_M : (state.unit === 'mm' ? n * FT_TO_MM : n);
    return scaled.toFixed(4);
}

function formatAngle(v) {
    const n = Number(v);
    return Number.isFinite(n) ? n.toFixed(3) : (v ?? '');
}

function repaintRows() {
    const tbody = document.querySelector('.kkyt-table tbody');
    if (!tbody) return;
    while (tbody.firstChild) tbody.removeChild(tbody.firstChild);
    state.rowsRaw.forEach(r => {
        const tr = document.createElement('tr');
        tr.append(tdText(r.File));
        HEADERS.forEach(h => tr.append(tdText(formatCoord(r[h.key]))));
        tr.append(tdText(formatAngle(r.TrueNorthAngle_deg ?? r['TrueNorthAngle(deg)'])));
        tbody.append(tr);
    });
}

function convertRowsForSave() {
    return (state.rowsRaw || []).map(r => ({ ...r }));
}

function toFileItems(list, root) {
    const base = root || commonDir(list);
    return dedupFiles((list || []).map(path => {
        const name = path?.split(/[\\/]/).pop() || path;
        return {
            path,
            rel: relPath(path, base),
            name,
            checked: true
        };
    }));
}

function dedupFiles(items) {
    const seen = new Set();
    const res = [];
    (items || []).forEach(f => {
        const key = (f?.path || '').toLowerCase();
        if (!key || seen.has(key)) return;
        seen.add(key);
        res.push(f);
    });
    return res;
}

function syncSaveState() {
    const saveBtn = document.getElementById('btnExSave');
    if (saveBtn) saveBtn.disabled = !state.rowsRaw.length || isWorking;
}

const PROGRESS_WEIGHTS = {
  COLLECT: 0.1,
  EXTRACT: 0.7,
  EXCEL: 0.05,
  EXCEL_INIT: 0.02,
  EXCEL_WRITE: 0.11,
  EXCEL_SAVE: 0.02,
  AUTOFIT: 0.0
};
const PROGRESS_ORDER = ['COLLECT', 'EXTRACT', 'EXCEL', 'EXCEL_INIT', 'EXCEL_WRITE', 'EXCEL_SAVE', 'AUTOFIT'];
let lastProgressPct = 0;
let isWorking = false;

function handleProgress(payload) {
    if (!payload) return;
    const phase = normalizePhase(payload.phase);
    const message = payload.message || '';
    const current = Number(payload.current || 0) || 0;
    const total = Number(payload.total || 0) || 0;
    const phaseProgress = clamp01(payload.phaseProgress);
    const percentFromHost = payload.percent;

    const percent = typeof percentFromHost === 'number'
        ? Math.max(lastProgressPct, percentFromHost)
        : computeProgressPercent(phase, current, total, phaseProgress);
    lastProgressPct = Math.max(lastProgressPct, percent);

    const subtitle = buildSubtitle(phase, current, total);
    const detail = buildDetail(message, phase, current, total);

    if (phase !== 'DONE' && phase !== 'ERROR') {
        setWorking(true);
    }
    ProgressDialog.show('좌표/북각 추출', subtitle);
    ProgressDialog.update(percent, subtitle, detail);

    if (phase === 'DONE') {
        setTimeout(() => { ProgressDialog.hide(); resetProgressState(); finishWorking(); }, 360);
    } else if (phase === 'ERROR') {
        resetProgressState();
        finishWorking();
        ProgressDialog.hide();
    }
}

function startProgress(phase, message, total) {
    resetProgressState();
    isWorking = true;
    const normalized = normalizePhase(phase);
    const detail = buildDetail(message || '', normalized, 0, total || 0);
    ProgressDialog.show('좌표/북각 추출', buildSubtitle(normalized, 0, total || 0));
    ProgressDialog.update(0, message || '', detail);
}

function finishWorking() {
    isWorking = false;
    syncPreviewState();
    syncSaveState();
    const pick = document.getElementById('btnExPick');
    if (pick) pick.disabled = false;
}

function setWorking(on) {
    isWorking = !!on;
    const pick = document.getElementById('btnExPick');
    if (pick) pick.disabled = isWorking;
    syncPreviewState();
    syncSaveState();
}

function handleError(message) {
    resetProgressState();
    finishWorking();
    ProgressDialog.hide();
    toast(message || '오류가 발생했습니다.', 'err', 3200);
}

function resetProgressState() {
    lastProgressPct = 0;
}

function buildSubtitle(phase, current, total) {
    const label = phaseLabel(phase);
    const count = total > 0 ? ` (${Math.max(current, 0)}/${total})` : '';
    return `${label}${count}`;
}

function buildDetail(message, phase, current, total) {
    const parts = [];
    if (message) parts.push(message);
    if (total > 0 && phase !== 'DONE' && phase !== 'ERROR') {
        parts.push(`진행률 ${Math.min(100, Math.max(0, computeProgressPercent(phase, current, total, 0))).toFixed(0)}%`);
    }
    return parts.join(' · ');
}

function computeProgressPercent(phase, current, total, phaseProgress) {
    if (phase === 'DONE') return 100;
    const normalized = normalizePhase(phase);
    if (normalized === 'ERROR') return lastProgressPct || 0;
    const completed = PROGRESS_ORDER.reduce((acc, key) => {
        if (key === normalized) return acc;
        return acc + (PROGRESS_WEIGHTS[key] || 0);
    }, 0);
    const weight = PROGRESS_WEIGHTS[normalized] || 0;
    const ratio = total > 0 ? Math.min(1, Math.max(0, current / total)) : 0;
    const staged = Math.max(ratio, clamp01(phaseProgress));
    const pct = (completed + weight * staged) * 100;
    return Math.max(lastProgressPct, Math.min(100, pct));
}

function normalizePhase(phase) {
    return String(phase || '').trim().toUpperCase() || 'EXTRACT';
}

function phaseLabel(phase) {
    switch (normalizePhase(phase)) {
        case 'COLLECT': return '파일 준비';
        case 'EXTRACT': return '포인트 추출';
        case 'EXCEL_INIT': return '엑셀 준비';
        case 'EXCEL_WRITE': return '엑셀 작성';
        case 'EXCEL_SAVE': return '파일 저장';
        case 'AUTOFIT': return 'AutoFit';
        case 'EXCEL': return '엑셀 내보내기';
        case 'DONE': return '완료';
        case 'ERROR': return '오류';
        default: return '진행 중';
    }
}

function clamp01(v) {
    const n = Number(v);
    if (!Number.isFinite(n)) return 0;
    return Math.max(0, Math.min(1, n));
}
