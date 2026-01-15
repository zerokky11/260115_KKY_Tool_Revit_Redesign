// Resources/HubUI/js/views/dup.js
import { clear, div, toast, showExcelSavedDialog, chooseExcelMode } from '../core/dom.js';
import { ProgressDialog } from '../core/progress.js';
import { onHost, post } from '../core/bridge.js';

// Host 이벤트 (fix2 고정)
const RESP_ROWS_EVENTS = ['dup:list', 'dup:rows', 'duplicate:list'];
const EV_DELETE_REQ = 'duplicate:delete';
const EV_RESTORE_REQ = 'duplicate:restore';
const EV_SELECT_REQ  = 'duplicate:select';
const EV_EXPORT_REQ  = 'duplicate:export';

const EV_DELETED_ONE   = 'dup:deleted';
const EV_RESTORED_ONE  = 'dup:restored';
const EV_DELETED_MULTI = 'duplicate:delete';
const EV_RESTORED_MULTI= 'duplicate:restore';
const EV_EXPORTED_A    = 'duplicate:export';
const EV_EXPORTED_B    = 'dup:exported';

export function renderDup(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);

  // 삭제행 시각 보정: 취소선은 없애고, 약간 흐리게만
  if (!document.getElementById('dup-style-override')) {
    const st = document.createElement('style');
    st.id = 'dup-style-override';
    st.textContent = `
      .dup-row.is-deleted .cell { text-decoration: none !important; opacity: .55; }
      .dup-row .row-actions .table-action-btn.restore {
        background: color-mix(in oklab, var(--accent, #4c6fff) 85%, #ffffff 15%);
        color:#fff;
      }
    `;
    document.head.appendChild(st);
  }

  // Topbar (HUB 헤더) 렌더 + sticky용 클래스
  const topbarEl = document.querySelector('#topbar-root .topbar') || document.querySelector('.topbar');
  if (topbarEl) topbarEl.classList.add('hub-topbar');

  // ===== 페이지 뼈대 =====
  const page    = div('dup-page feature-shell');

    const header = div('feature-header dup-toolbar');
    const heading = div('feature-heading');
    heading.innerHTML = `
      <span class="feature-kicker">Duplicate Inspector</span>
      <h2 class="feature-title">중복검토</h2>
      <p class="feature-sub">중복 패밀리/요소를 그룹별로 확인하고 삭제/되돌리기를 관리합니다.</p>`;

  const runBtn    = cardBtn('검토 시작', onRun);
  const exportBtn = cardBtn('엑셀 내보내기', onExport);
  exportBtn.disabled = true;

  const actions = div('feature-actions');
  actions.append(runBtn, exportBtn);
  header.append(heading, actions);
  page.append(header);

  // sticky summary bar
  const summaryBar = div('dup-summarybar sticky hidden');
  page.append(summaryBar);

  const body = div('dup-body');
  page.append(body);
  target.append(page);

  const EXCEL_PHASE_WEIGHT = { EXCEL_INIT: 0.05, EXCEL_WRITE: 0.85, EXCEL_SAVE: 0.08, AUTOFIT: 0.02, DONE: 1, ERROR: 1 };
  const EXCEL_PHASE_ORDER = ['EXCEL_INIT', 'EXCEL_WRITE', 'EXCEL_SAVE', 'AUTOFIT', 'DONE'];

  // ---- state ----
  let rows      = [];
  let groups    = [];
  let deleted   = new Set();
  let expanded  = new Set();
  let waitTimer = null;
  let busy      = false;
  let exporting = false;
  let lastExcelPct = 0;

  renderIntro(body);

  // 공통 오류
  onHost('revit:error', ({ message }) => {
    setLoading(false);
    ProgressDialog.hide();
    lastExcelPct = 0;
    exporting = false;
    exportBtn.disabled = rows.length === 0;
    toast(message || 'Revit 오류가 발생했습니다.', 'err', 3200);
  });
  onHost('host:error',  ({ message }) => {
    setLoading(false);
    ProgressDialog.hide();
    lastExcelPct = 0;
    exporting = false;
    toast(message || '호스트 오류가 발생했습니다.', 'err', 3200);
  });
  onHost('host:warn',   ({ message }) => {
    setLoading(false);
    if (message) toast(message, 'warn', 3600);
  });

  // 모든 이벤트 수신
  onHost(({ ev, payload }) => {
    if (RESP_ROWS_EVENTS.includes(ev)) {
      setLoading(false);
      const list = payload?.rows ?? payload?.data ?? payload ?? [];
      handleRows(list);
      return;
    }

    if (ev === 'dup:result') {
      setLoading(false);
      return;
    }

    if (ev === EV_DELETED_ONE) {
      const id = String(payload?.id ?? '');
      if (id) {
        deleted.add(id);
        updateRowStates();
        refreshSummary();
      }
      return;
    }

    if (ev === EV_RESTORED_ONE) {
      const id = String(payload?.id ?? '');
      if (id) {
        deleted.delete(id);
        updateRowStates();
        refreshSummary();
      }
      return;
    }

    if (ev === EV_DELETED_MULTI) {
      (toIdArray(payload?.ids)).forEach(id => deleted.add(id));
      updateRowStates();
      refreshSummary();
      return;
    }

    if (ev === EV_RESTORED_MULTI) {
      (toIdArray(payload?.ids)).forEach(id => deleted.delete(id));
      updateRowStates();
      refreshSummary();
      return;
    }

    if (ev === EV_SELECT_REQ) {
      return;
    }

    if (ev === 'dup:progress') {
      handleExcelProgress(payload || {});
      return;
    }

    if (ev === EV_EXPORTED_A || ev === EV_EXPORTED_B) {
      lastExcelPct = 0;
      ProgressDialog.hide();
      const path = payload?.path || '';
      if (payload?.ok || path) {
        showExcelSavedDialog('엑셀로 내보냈습니다.', path, (p) => {
          if (p) post('excel:open', { path: p });
        });
      } else {
        toast(payload?.message || '엑셀 내보내기 실패', 'err');
      }
      exporting = false;
      exportBtn.disabled = rows.length === 0;
      return;
    }
  });

  // ===== 액션 =====

  function setLoading(on) {
    busy = on;
    runBtn.disabled = on;
    runBtn.textContent = on ? '검토 중…' : '검토 시작';

    if (!on && waitTimer) {
      clearTimeout(waitTimer);
      waitTimer = null;
    }
  }

  function onRun() {
    setLoading(true);
    exportBtn.disabled = true;
    deleted.clear();

    body.innerHTML = '';
    body.append(buildSkeleton(6));

    waitTimer = setTimeout(() => {
      setLoading(false);
      toast('응답이 없습니다. Add-in 이벤트명을 확인하세요 (예: dup:list).', 'err');
      body.innerHTML = '';
      renderIntro(body);
    }, 10000);

    post('dup:run', {});
  }

  function onExport() {
    if (exporting) return;
    exporting = true;
    exportBtn.disabled = true;
    chooseExcelMode((mode) => post(EV_EXPORT_REQ, { excelMode: mode || 'fast' }));
  }

  function handleExcelProgress(payload) {
    if (!payload) {
      ProgressDialog.hide();
      exporting = false;
      exportBtn.disabled = rows.length === 0;
      lastExcelPct = 0;
      return;
    }
    const phase = normalizeExcelPhase(payload.phase);
    const total = Number(payload.total) || 0;
    const current = Number(payload.current) || 0;
    const percent = computeExcelPercent(phase, current, total, payload.phaseProgress);
    const subtitle = buildExcelSubtitle(phase, current, total);
    const detail = payload.message || '';

    exporting = phase !== 'DONE' && phase !== 'ERROR';
    exportBtn.disabled = exporting || rows.length === 0;

    ProgressDialog.show('중복검토 엑셀 내보내기', subtitle);
    ProgressDialog.update(percent, subtitle, detail);

    if (!exporting) {
      setTimeout(() => { ProgressDialog.hide(); lastExcelPct = 0; }, 280);
    }
  }

  function normalizeExcelPhase(phase) {
    return String(phase || '').trim().toUpperCase() || 'EXCEL_WRITE';
  }

  function computeExcelPercent(phase, current, total, phaseProgress) {
    const norm = normalizeExcelPhase(phase);
    if (norm === 'DONE') { lastExcelPct = 100; return 100; }
    if (norm === 'ERROR') return lastExcelPct;
    const completed = EXCEL_PHASE_ORDER.reduce((acc, key) => {
      if (key === norm) return acc;
      return acc + (EXCEL_PHASE_WEIGHT[key] || 0);
    }, 0);
    const weight = EXCEL_PHASE_WEIGHT[norm] || 0;
    const ratio = total > 0 ? Math.min(1, Math.max(0, current / total)) : 0;
    const staged = Math.max(ratio, clamp01(phaseProgress));
    const pct = Math.min(100, Math.max(lastExcelPct, (completed + weight * staged) * 100));
    lastExcelPct = pct;
    return pct;
  }

  function buildExcelSubtitle(phase, current, total) {
    const labelMap = {
      EXCEL_INIT: '엑셀 준비',
      EXCEL_WRITE: '엑셀 작성',
      EXCEL_SAVE: '파일 저장',
      AUTOFIT: 'AutoFit',
      DONE: '완료',
      ERROR: '오류'
    };
    const label = labelMap[normalizeExcelPhase(phase)] || '엑셀 진행';
    const count = total > 0 ? ` (${Math.max(current, 0)}/${total})` : '';
    return `${label}${count}`;
  }

  function clamp01(v) {
    const n = Number(v);
    if (!Number.isFinite(n)) return 0;
    return Math.max(0, Math.min(1, n));
  }

  // ===== 호스트 응답 처리 =====

  function handleRows(listLike) {
    const list = Array.isArray(listLike) ? listLike : [];
    rows   = list.map(normalizeRow);
    groups = buildGroups(rows);

    exportBtn.disabled = rows.length === 0;
    setLoading(false);

    // 처음엔 앞쪽 10개 그룹만 펼쳐두기
    expanded = new Set(groups.slice(0, 10).map(g => g.key));
    paintGroups();

    if (!rows.length) {
      body.innerHTML = '';
      const empty = div('dup-emptycard');
      empty.innerHTML = `
        <div class="icon">✅</div>
        <div class="title">중복이 없어요</div>
        <div class="desc">모델 상태가 깨끗합니다. 필요 시 다시 검토를 실행하세요.</div>
      `;
      body.append(empty);
    }

    refreshSummary();
  }

  // ===== 렌더링 =====

  function paintGroups() {
    body.innerHTML = '';

    groups.forEach((g, idx) => {
      const card = div('dup-grp');
      card.classList.add(g.rows.length >= 2 ? 'accent-danger' : 'accent-info');

      // 그룹 헤더
      const h    = div('grp-h');
      const left = div('grp-txt');
      const famLabel  = g.family  ? g.family  : (g.category ? `${g.category} Type` : '—');
      const typeLabel = g.type || '—';

      left.innerHTML = `
        <div class="grp-line">
          <span class="chip alt">중복 그룹 ${idx + 1}</span>
          <span class="grp-cat mono">${esc(g.category || '—')}</span>
          <span class="grp-sep">·</span>
          <span class="grp-fam">${esc(famLabel)}</span>
          <span class="grp-sep">·</span>
          <span class="grp-fam">${esc(typeLabel)}</span>
          <span class="chip mono tone">${g.rows.length}개</span>
        </div>
      `;

      const right  = div('grp-actions');
      const toggle = kbtn(expanded.has(g.key) ? '접기' : '펼치기', 'subtle', () => toggleGroup(g.key));
      right.append(toggle);

      h.append(left, right);
      card.append(h);

      // 서브헤더 + 행
      const tbl = div('grp-body');
      const sh  = div('dup-subhead');

      sh.append(
        cell('', 'ck'),
        cell('Element ID', 'th'),
        cell('Category', 'th'),
        cell('Family', 'th'),
        cell('Type', 'th'),
        cell('작업', 'th right')
      );

      tbl.append(sh);

      if (expanded.has(g.key)) {
        g.rows.forEach(r => tbl.append(renderRow(r)));
      }

      card.append(tbl);
      body.append(card);
    });

    updateRowStates();
  }

  function renderRow(r) {
    const row = div('dup-row');
    row.dataset.id = r.id;

    const ckCell = cell(null, 'ck');
    const ck     = document.createElement('input');
    ck.type      = 'checkbox';
    ck.className = 'ckbox';
    ck.onchange  = () => row.classList.toggle('is-selected', ck.checked);
    ckCell.append(ck);
    row.append(ckCell);

    row.append(cell(r.id ?? '-', 'td mono right'));
    row.append(cell(r.category || '—', 'td'));

    const famOut = r.family ? r.family : (r.category ? `${r.category} Type` : '—');
    row.append(cell(famOut, 'td ell'));
    row.append(cell(r.type || '—', 'td ell'));

    const act    = div('row-actions');
    const viewBtn = tableBtn('선택/줌', '', () =>
      post(EV_SELECT_REQ, { id: r.id, zoom: true, mode: 'zoom' })
    );

    const delBtn = tableBtn(r.deleted ? '되돌리기' : '삭제',
                        r.deleted ? 'restore' : 'table-action-btn--danger',
                        () => {
                          const ids = [r.id];
                          if (delBtn.dataset.mode === 'restore') {
                            post(EV_RESTORE_REQ, { id: r.id, ids });
                          } else {
                            post(EV_DELETE_REQ,  { id: r.id, ids });
                          }
                        });

    delBtn.dataset.mode = r.deleted ? 'restore' : 'delete';

    act.append(viewBtn, delBtn);
    row.append(cell(act, 'td right'));

    return row;
  }

  function toggleGroup(key) {
    if (expanded.has(key)) expanded.delete(key);
    else expanded.add(key);
    paintGroups();
  }

  function updateRowStates() {
    [...document.querySelectorAll('.dup-row')].forEach(rowEl => {
      const id = rowEl.dataset.id;
      if (!id) return;

      const isDel = deleted.has(String(id));
      rowEl.classList.toggle('is-deleted', isDel);

      const delBtn = rowEl.querySelector('.row-actions .table-action-btn:last-child');
      if (delBtn) {
        delBtn.textContent = isDel ? '되돌리기' : '삭제';
        delBtn.className   = 'table-action-btn ' + (isDel ? 'restore' : 'table-action-btn--danger');
        delBtn.dataset.mode = isDel ? 'restore' : 'delete';
      }

      const ck = rowEl.querySelector('input.ckbox');
      if (ck) rowEl.classList.toggle('is-selected', !!ck.checked);
    });
  }

  function refreshSummary() {
    const totals = computeSummary(groups);
    summaryBar.innerHTML = '';
    summaryBar.classList.toggle('hidden', totals.totalCount === 0 && !busy);

    [chip(`그룹 ${totals.groupCount}`), chip(`요소 ${totals.totalCount}`)]
      .forEach(c => summaryBar.append(c));
  }

  // ===== 유틸 =====

  function cardBtn(label, handler) {
    const b = document.createElement('button');
    b.className = 'card-action-btn';
    b.type      = 'button';
    b.textContent = label;
    b.onclick   = handler;
    return b;
  }

  function tableBtn(label, tone, handler) {
    const b = document.createElement('button');
    b.className = 'table-action-btn ' + (tone || '');
    b.type = 'button';
    b.textContent = label;
    b.onclick = handler;
    return b;
  }

  function cell(content, cls) {
    const c = document.createElement('div');
    c.className = 'cell ' + (cls || '');
    if (content instanceof Node) c.append(content);
    else if (content != null)   c.textContent = content;
    return c;
  }

  function chip(text, tone) {
    const b = div('chip ' + (tone || ''));
    b.textContent = text;
    return b;
  }

  // Localized compact button (replaces missing legacy kbtn helper)
  function kbtn(label, tone, handler) {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'control-chip chip-btn ' + (tone || '');
    b.textContent = label;
    b.onclick = handler;
    return b;
  }

  function toIdArray(v) {
    if (!v) return [];
    if (Array.isArray(v)) return v.map(String);
    return [String(v)];
  }

  function esc(s) {
    return String(s ?? '').replace(/[&<>"']/g, m => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    }[m]));
  }

  function normalizeRow(r) {
    const id  = safeId(r.elementId ?? r.ElementId ?? r.id ?? r.Id);
    const category = val(r.category ?? r.Category);
    const family   = val(r.family   ?? r.Family);
    const type     = val(r.type     ?? r.Type);

    const connectedIdsRaw =
      r.connectedIds ?? r.ConnectedIds ?? r.links ?? r.Links ?? r.connected ?? [];

    const connectedIds = Array.isArray(connectedIdsRaw)
      ? connectedIdsRaw.map(String)
      : (typeof connectedIdsRaw === 'string' && connectedIdsRaw.length
          ? connectedIdsRaw.split(/[,\s]+/).filter(Boolean)
          : []);

    const deletedFlag = !!(r.deleted ?? r.isDeleted ?? r.Deleted);

    return { id: id || '-', category, family, type, connectedIds, deleted: deletedFlag };
  }

  // (Category / Family / Type / 연결세트) 기준 그룹
  function buildGroups(rs) {
    const map = new Map();

    for (const r of rs) {
      const cluster = [String(r.id), ...r.connectedIds.map(String)]
        .filter(Boolean)
        .map(x => x.trim())
        .sort((a, b) => Number(a) - Number(b))
        .join(',');

      const key = [r.category || '', r.family || '', r.type || '', cluster].join('|');
      let g = map.get(key);
      if (!g) {
        g = { key, category: r.category || '', family: r.family || '', type: r.type || '', rows: [] };
        map.set(key, g);
      }
      g.rows.push(r);
    }

    return [...map.values()];
  }

  function computeSummary(groups) {
    let total = 0;
    groups.forEach(g => { total += g.rows.length; });
    return { groupCount: groups.length, totalCount: total };
  }

  function safeId(v) {
    if (v === 0) return 0;
    if (v == null) return '';
    return String(v);
  }

  function val(v) {
    return v == null || v === '' ? '' : String(v);
  }

  function renderIntro(container) {
    const hero = div('dup-hero');
    hero.innerHTML = `
      <div class="ill">🧭</div>
      <div class="title">중복검토를 시작해 보세요</div>
      <div class="desc">
        모델의 중복 요소를 그룹으로 묶어 보여줍니다.
        각 행에서 <b>삭제/되돌리기</b>, <b>선택/줌</b>을 실행할 수 있어요.
      </div>
      <ul class="tips">
        <li>그룹 헤더 우측의 <b>펼치기</b>로 상세를 열어보세요.</li>
        <li>System Family는 <b>"Category Type"</b> 으로 표시됩니다.</li>
        <li>엑셀 내보내기는 결과가 있을 때만 활성화됩니다.</li>
      </ul>
    `;
    container.append(hero);
  }

  function buildSkeleton(n = 6) {
    const wrap = div('dup-skeleton');
    for (let i = 0; i < n; i++) {
      const line = div('sk-row');
      line.append(
        div('sk-chip'),
        div('sk-id'),
        div('sk-wide'),
        div('sk-wide'),
        div('sk-act')
      );
      wrap.append(line);
    }
    return wrap;
  }
}
