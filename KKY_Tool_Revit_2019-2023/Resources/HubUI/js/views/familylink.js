// Resources/HubUI/js/views/familylink.js
import { clear, div, toast, debounce, showExcelSavedDialog, chooseExcelMode } from '../core/dom.js';
import { ProgressDialog } from '../core/progress.js';
import { onHost, post } from '../core/bridge.js';
import { createRvtTable, renderRvtRows, getRvtName } from './rvtTable.js';

const DEFAULT_SCHEMA = [
  'FileName',
  'HostFamilyName',
  'HostFamilyCategory',
  'NestedFamilyName',
  'NestedTypeName',
  'NestedCategory',
  'TargetParamName',
  'ExpectedGuid',
  'FoundScope',
  'NestedParamGuid',
  'NestedParamDataType',
  'AssocHostParamName',
  'HostParamGuid',
  'HostParamIsShared',
  'Issue',
  'Notes'
];

export function renderFamilyLink(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);

  const topbarEl = document.querySelector('#topbar-root .topbar') || document.querySelector('.topbar');
  if (topbarEl) topbarEl.classList.add('hub-topbar');

  const state = {
    items: [],
    groups: [],
    selectedGroups: new Set(['(All Groups)']),
    selectedParams: new Set(),
    search: '',
    schema: DEFAULT_SCHEMA.slice(),
    rows: [],
    rvtPaths: [],
    rvtChecked: new Set(),
    busy: false
  };

  const page = div('familylink-page feature-shell');

  const header = div('feature-header');
  const heading = div('feature-heading');
  heading.innerHTML = `
    <span class="feature-kicker">Nested Family Association</span>
    <h2 class="feature-title">패밀리 연동 검토(복합/네스티드)</h2>
    <p class="feature-sub">Shared GUID 기준으로 네스티드 패밀리 파라미터 연동 상태를 점검합니다.</p>`;

  const runBtn = cardBtn('검토 시작', onRun);
  const exportBtn = cardBtn('엑셀 내보내기', onExport);
  exportBtn.disabled = true;

  const actions = div('feature-actions');
  actions.append(runBtn, exportBtn);
  header.append(heading, actions);
  page.append(header);

  const body = div('familylink-body');
  const topPanels = div('familylink-top-panels');
  const resultsPanel = div('familylink-results-panel');

  // ----- Shared Parameter 선택 (ParamProp 스타일 재사용) -----
  const paramCard = div('paramprop-card section familylink-card');
  const paramTitle = div('paramprop-title');
  paramTitle.textContent = '공유 파라미터 선택';
  paramCard.append(paramTitle);

  const sourceLine = div('familylink-source');
 


  const searchRow = div('paramprop-row paramprop-search-row');
  const searchBox = document.createElement('input');
  searchBox.type = 'search';
  searchBox.placeholder = '이름 또는 그룹 검색';
  searchBox.className = 'paramprop-search';
  searchBox.addEventListener('input', debounce((e) => {
    state.search = (e.target.value || '').trim();
    renderParamTable();
  }, 120));
  searchRow.append(labelSpan('검색'), searchBox);
  paramCard.append(searchRow);

  const selectGrid = div('paramprop-grid');
  paramCard.append(selectGrid);

  const groupBox = div('paramprop-table-box paramprop-group-box');
  const groupHeader = div('paramprop-subtitle');
  groupHeader.textContent = '그룹 (다중 선택)';
  const groupTable = document.createElement('table');
  groupTable.className = 'paramprop-table paramprop-group-table';
  const groupThead = document.createElement('thead');
  groupThead.innerHTML = '<tr><th>선택</th><th>Group</th></tr>';
  const groupTbody = document.createElement('tbody');
  groupTable.append(groupThead, groupTbody);
  const groupListWrap = div('paramprop-table-wrap paramprop-group-wrap');
  groupListWrap.append(groupTable);
  groupBox.append(groupHeader, groupListWrap);
  selectGrid.append(groupBox);

  const tableBox = div('paramprop-table-box');
  const tableHead = div('paramprop-subtitle');
  const selectedCount = document.createElement('span');
  selectedCount.className = 'familylink-selected-count';
  tableHead.textContent = '파라미터';
  tableHead.append(selectedCount);
  const table = document.createElement('table');
  table.className = 'paramprop-table';
  const thead = document.createElement('thead');
  thead.innerHTML = '<tr><th>선택</th><th>Name</th></tr>';
  const tbody = document.createElement('tbody');
  table.append(thead, tbody);
  const tableWrap = div('paramprop-table-wrap');
  tableWrap.append(table);
  tableBox.append(tableHead, tableWrap);
  selectGrid.append(tableBox);

  // ----- RVT 목록 -----
  const rvtCard = div('paramprop-card section familylink-card');
  const rvtTitle = div('paramprop-title');
  rvtTitle.textContent = '대상 RVT 목록';
  rvtCard.append(rvtTitle);

  const rvtActions = div('familylink-rvt-actions');
  const btnAddRvt = cardBtn('RVT 파일 추가', () => post('familylink:pick-rvts', {}));
  const btnRemoveRvt = cardBtn('선택 제거', () => {
    if (!state.rvtChecked.size) return;
    state.rvtPaths = state.rvtPaths.filter(p => !state.rvtChecked.has(p));
    state.rvtChecked.clear();
    renderRvtList();
    syncRunState();
  });
  const btnClearRvt = cardBtn('등록 목록 비우기', () => {
    state.rvtPaths = [];
    state.rvtChecked.clear();
    renderRvtList();
    syncRunState();
  });
  rvtActions.append(btnAddRvt, btnRemoveRvt, btnClearRvt);
  rvtCard.append(rvtActions);

  const rvtTableWrap = div('familylink-rvt-table');
  const { table: rvtTable, tbody: rvtTbody, master: rvtMaster } = createRvtTable();
  rvtTableWrap.append(rvtTable);
  rvtCard.append(rvtTableWrap);

  // ----- 결과 영역 -----
  const resultHead = div('familylink-results-head');
  const resultTitle = div('familylink-results-title');
  resultTitle.textContent = '검토 결과';
  const resultMeta = div('familylink-results-meta');
  resultMeta.textContent = '0 rows';
  resultHead.append(resultTitle, resultMeta);

  const resultBody = div('familylink-results-body');
  const resultTable = document.createElement('table');
  resultTable.className = 'familylink-table';
  const resultThead = document.createElement('thead');
  const resultTbody = document.createElement('tbody');
  resultTable.append(resultThead, resultTbody);
  resultBody.append(resultTable);

  resultsPanel.append(resultHead, resultBody);

  topPanels.append(paramCard, rvtCard);
  body.append(topPanels, resultsPanel);
  page.append(body);
  target.append(page);

  renderGroups();
  renderParamTable();
  renderRvtList();
  renderResultTable();
  syncRunState();

  onHost('familylink:sharedparams', handleSharedParams);
  onHost('familylink:rvts-picked', handleRvtsPicked);
  onHost('familylink:progress', handleProgress);
  onHost('familylink:result', handleResult);
  onHost('familylink:error', handleError);
  onHost('familylink:exported', handleExported);

  post('familylink:init', {});

  function handleSharedParams(payload) {
    state.items = Array.isArray(payload?.items) ? payload.items : [];
    state.groups = deriveGroups(state.items);
    state.selectedGroups = new Set(['(All Groups)']);
    state.selectedParams = new Set();
    state.search = '';
    searchBox.value = '';

   

    renderGroups();
    renderParamTable();
    syncRunState();
  }

  function handleRvtsPicked(payload) {
    const paths = Array.isArray(payload?.paths) ? payload.paths : [];
    const existing = new Set(state.rvtPaths.map(p => p.toLowerCase()));
    paths.forEach(p => {
      if (!p) return;
      const key = p.toLowerCase();
      if (!existing.has(key)) {
        existing.add(key);
        state.rvtPaths.push(p);
      }
    });
    renderRvtList();
    syncRunState();
  }

  function handleProgress(payload) {
    if (!payload) return;
    const pct = Math.max(0, Math.min(100, Number(payload.percent) || 0));
    const msg = payload.message || '';
    ProgressDialog.show('패밀리 연동 검토', msg || '진행 중...');
    ProgressDialog.update(pct, msg || '진행 중...', '');
    if (pct >= 100) {
      setTimeout(() => ProgressDialog.hide(), 350);
    }
  }

  function handleResult(payload) {
    state.rows = Array.isArray(payload?.rows) ? payload.rows : [];
    state.schema = Array.isArray(payload?.schema) && payload.schema.length ? payload.schema : DEFAULT_SCHEMA.slice();
    renderResultTable();
    exportBtn.disabled = state.rows.length === 0 || state.busy;
    setBusy(false);
    ProgressDialog.hide();
  }

  function handleError(payload) {
    setBusy(false);
    ProgressDialog.hide();
    const message = payload?.message || '작업 중 오류가 발생했습니다.';
    toast(message, 'err', 3200);
  }

  function handleExported(payload) {
    const ok = payload?.ok !== false && payload?.path;
    exportBtn.disabled = state.rows.length === 0 || state.busy;
    if (ok) {
      showExcelSavedDialog('엑셀 내보내기가 완료되었습니다.', payload.path, (p) => post('excel:open', { path: p }));
    } else {
      toast(payload?.message || '엑셀 내보내기 실패', 'err');
    }
  }

  function onRun() {
    if (state.busy) return;
    if (!state.rvtPaths.length) {
      toast('검토할 RVT 파일을 추가하세요.', 'warn');
      return;
    }

    const targets = Array.from(state.selectedParams)
      .map(guid => state.items.find(item => item.guid === guid))
      .filter(Boolean)
      .map(item => ({ name: item.name, guid: item.guid }));

    if (!targets.length) {
      toast('검토할 파라미터를 선택하세요.', 'warn');
      return;
    }

    setBusy(true);
    exportBtn.disabled = true;
    ProgressDialog.show('패밀리 연동 검토', '준비 중...');
    ProgressDialog.update(0, '준비 중...', '');

    post('familylink:run', {
      rvtPaths: state.rvtPaths.slice(),
      targets
    });
  }

  function onExport() {
    if (state.busy || !state.rows.length) return;
    exportBtn.disabled = true;
    chooseExcelMode((mode) => {
      const selected = mode || 'fast';
      post('familylink:export', {
        fastExport: selected === 'fast',
        autoFit: selected === 'normal'
      });
    });
  }

  function setBusy(on) {
    state.busy = on;
    runBtn.disabled = on;
    runBtn.textContent = on ? '검토 중…' : '검토 시작';
    syncRunState();
  }

  function syncRunState() {
    const hasTargets = state.selectedParams.size > 0;
    const hasRvts = state.rvtPaths.length > 0;
    runBtn.disabled = state.busy || !(hasTargets && hasRvts);
    exportBtn.disabled = state.busy || state.rows.length === 0;
    btnRemoveRvt.disabled = state.rvtChecked.size === 0;
  }

  function renderGroups() {
    groupTbody.innerHTML = '';
    const allItem = makeGroupItem('(All Groups)');
    groupTbody.append(allItem);
    state.groups.forEach(g => groupTbody.append(makeGroupItem(g)));
  }

  function makeGroupItem(name) {
    const tr = document.createElement('tr');
    tr.className = state.selectedGroups.has(name) ? 'is-selected' : '';
    tr.dataset.group = name;
    const chk = document.createElement('input');
    chk.type = 'checkbox';
    chk.checked = state.selectedGroups.has(name);
    chk.addEventListener('change', (e) => { e.stopPropagation(); toggleGroup(name, chk.checked); });
    const tdChk = document.createElement('td');
    tdChk.append(chk);
    const nameCell = td(name);
    nameCell.title = name;
    tr.append(tdChk, nameCell);
    tr.addEventListener('click', () => toggleGroup(name, !state.selectedGroups.has(name)));
    return tr;
  }

  function toggleGroup(name, on) {
    if (name === '(All Groups)') {
      state.selectedGroups = on ? new Set(['(All Groups)']) : new Set();
    } else {
      const sg = new Set(state.selectedGroups);
      sg.delete('(All Groups)');
      if (on) sg.add(name); else sg.delete(name);
      if (sg.size === 0) sg.add('(All Groups)');
      state.selectedGroups = sg;
    }
    renderGroups();
    renderParamTable();
  }

  function renderParamTable() {
    tbody.innerHTML = '';
    updateSelectedCount();

    const filtered = filterDefs();
    if (!filtered.length) {
      const tr = document.createElement('tr');
      const tdEmpty = document.createElement('td');
      tdEmpty.colSpan = 2;
      tdEmpty.textContent = state.items.length
        ? '조건에 맞는 항목이 없습니다.'
        : 'Shared Parameter 등록이 필요합니다. Revit에서 Shared Parameter Text를 등록/연결 후 다시 시도하세요.';
      tdEmpty.className = 'paramprop-empty';
      tr.append(tdEmpty);
      tbody.append(tr);
      return;
    }

    filtered.forEach(def => {
      const key = def.guid;
      const tr = document.createElement('tr');
      tr.dataset.key = key;
      tr.dataset.group = def.groupName || '';
      tr.className = state.selectedParams.has(key) ? 'is-selected' : '';
      const tdChk = document.createElement('td');
      const chk = document.createElement('input');
      chk.type = 'checkbox';
      chk.checked = state.selectedParams.has(key);
      chk.addEventListener('change', (e) => {
        e.stopPropagation();
        if (chk.checked) state.selectedParams.add(key); else state.selectedParams.delete(key);
        renderParamTable();
        syncRunState();
      });
      tdChk.append(chk);
      const nameCell = td(def.name);
      nameCell.title = `${def.groupName || ''} • ${def.dataTypeToken || ''}`.trim();
      tr.append(tdChk, nameCell);
      tr.addEventListener('click', () => {
        if (state.selectedParams.has(key)) state.selectedParams.delete(key); else state.selectedParams.add(key);
        renderParamTable();
        syncRunState();
      });
      tbody.append(tr);
    });
  }

  function updateSelectedCount() {
    const count = state.selectedParams.size;
    selectedCount.textContent = count ? ` (선택 ${count}개)` : '';
  }

  function filterDefs() {
    const groups = state.selectedGroups;
    const search = state.search.toLowerCase();
    return state.items.filter(d => {
      const inGroup = groups.has('(All Groups)') || groups.has(d.groupName);
      if (!inGroup) return false;
      if (!search) return true;
      return (d.name || '').toLowerCase().includes(search) || (d.groupName || '').toLowerCase().includes(search);
    });
  }

  function deriveGroups(items) {
    const set = new Set();
    (Array.isArray(items) ? items : []).forEach(d => {
      if (d?.groupName) set.add(d.groupName);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'ko'));
  }

  function renderRvtList() {
    const rows = state.rvtPaths.map((path, idx) => ({
      index: idx + 1,
      name: getRvtName(path, '—'),
      path,
      checked: state.rvtChecked.has(path),
      onToggle: (checked) => {
        if (checked) state.rvtChecked.add(path);
        else state.rvtChecked.delete(path);
        syncRunState();
      }
    }));

    renderRvtRows(rvtTbody, rows, '등록된 RVT가 없습니다.');
    rvtMaster.checked = rows.length > 0 && rows.every(r => r.checked);
    rvtMaster.indeterminate = rows.some(r => r.checked) && !rvtMaster.checked;
    rvtMaster.onchange = () => {
      state.rvtChecked.clear();
      if (rvtMaster.checked) rows.forEach(r => state.rvtChecked.add(r.path));
      renderRvtList();
      syncRunState();
    };
  }

  function renderResultTable() {
    resultThead.innerHTML = '';
    resultTbody.innerHTML = '';

    const headRow = document.createElement('tr');
    state.schema.forEach(h => {
      const th = document.createElement('th');
      th.textContent = h;
      headRow.append(th);
    });
    resultThead.append(headRow);

    if (!state.rows.length) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = state.schema.length;
      td.className = 'familylink-empty';
      td.textContent = '결과가 없습니다. 검토를 실행하세요.';
      tr.append(td);
      resultTbody.append(tr);
    } else {
      state.rows.forEach(row => {
        const tr = document.createElement('tr');
        state.schema.forEach(h => {
          const td = document.createElement('td');
          const val = row?.[h];
          td.textContent = val == null ? '' : String(val);
          tr.append(td);
        });
        resultTbody.append(tr);
      });
    }

    resultMeta.textContent = `${state.rows.length} rows`;
  }
}

function cardBtn(text, onClick) {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'btn card-btn';
  btn.textContent = text;
  btn.onclick = onClick;
  return btn;
}

function labelSpan(text) {
  const span = document.createElement('span');
  span.className = 'paramprop-label';
  span.textContent = text;
  return span;
}

function td(value) {
  const cell = document.createElement('td');
  cell.textContent = value ?? '';
  return cell;
}
