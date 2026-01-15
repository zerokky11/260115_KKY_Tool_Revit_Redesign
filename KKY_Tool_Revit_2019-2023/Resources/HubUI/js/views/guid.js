import { clear, div, toast, showExcelSavedDialog, chooseExcelMode } from '../core/dom.js';
import { createRvtTable, renderRvtRows, getRvtName } from './rvtTable.js';
import { ProgressDialog } from '../core/progress.js';
import { post, onHost } from '../core/bridge.js';

const LS_RVTS = 'kky_guid_rvts';

const HIDDEN_DETAIL_COLS = new Set(['RvtPath']);
const EXCEL_PHASE_WEIGHT = { EXCEL_INIT: 0.05, EXCEL_WRITE: 0.85, EXCEL_SAVE: 0.08, AUTOFIT: 0.02, DONE: 1, ERROR: 1 };

export function renderGuid(root) {
    const target = root || document.getElementById('view-root') || document.getElementById('app');
    clear(target);
    const top = document.querySelector('#topbar-root .topbar') || document.querySelector('.topbar'); if (top) top.classList.add('hub-topbar');

    const initialRvtList = loadRvtList();
    const state = {
        rvtList: initialRvtList,
        rvtChecked: new Set(initialRvtList),
        project: { columns: [], rows: [] },
        familyNav: { columns: [], rows: [] },
        familyDetail: { columns: [], rows: [] },
        activeProjectDoc: '',
        activeFamilyDoc: '',
        activeFamily: '',
        activeTab: 'project',
        busy: false,
        runId: '',
        hasRun: false,
        includeFamily: false,
        includeAnnotation: false,
        familyFilter: 'all',
        sharedParamStatus: null
    };
    let lastExcelPct = 0;

    const page = div('feature-shell guid-page');

    // Header
    const header = div('feature-header guid-header');
    const headerLeft = div('feature-heading');
    headerLeft.innerHTML = `
      <span class="feature-kicker">GUID Audit</span>
      <h2 class="feature-title">공유 파라미터 GUID 검토</h2>
      <p class="feature-sub">프로젝트/패밀리 파라미터 GUID를 공유 파라미터 파일과 비교합니다.</p>`;

    const headerRight = div('guid-header-right');
    const optionRow = div('guid-options');
    const modeToggle = buildModeToggle();
    const annotationToggle = buildAnnotationToggle();
    optionRow.append(modeToggle, annotationToggle);
    const sharedStatus = buildSharedParamStatus();
    headerRight.append(optionRow, sharedStatus);
    header.append(headerLeft, headerRight);
    page.append(header);

    const body = div('guid-body');

    // RVT section
    const rvtSection = div('feature-results-panel guid-panel segmentpms-extract');
    const rvtHeader = document.createElement('div');
    rvtHeader.className = 'feature-results-head guid-rvt-head';
    const rvtTitle = document.createElement('div');
    rvtTitle.className = 'guid-title';
    rvtTitle.innerHTML = '<h3>대상 RVT 목록</h3><p class="feature-note">비우면 현재 활성 문서를 사용합니다.</p>';
    const rvtActions = div('segmentpms-actions-row');
    let btnRemove = null;
    let btnClear = null;
    const btnAdd = cardBtn('RVT 파일 추가', () => post('guid:add-files', { pick: 'files' }));
    const btnAddFolder = cardBtn('폴더 선택', () => post('guid:add-files', { pick: 'folder' }));
    btnRemove = cardBtn('선택 제거', onRemoveSelected);
    btnRemove.disabled = true;
    btnClear = cardBtn('등록 목록 비우기', () => { state.rvtList = []; state.rvtChecked.clear(); persistRvts(); renderRvtList(); syncRvtActionState(); });
    const runBtn = cardBtn('검토 시작', onRun);
    const exportBtn = cardBtn('엑셀 내보내기', onExport);
    runBtn.classList.add('btn-primary');
    exportBtn.classList.add('btn-outline');
    exportBtn.disabled = true;
    rvtActions.append(btnAdd, btnAddFolder, btnRemove, btnClear, runBtn, exportBtn);

    rvtHeader.append(rvtTitle, rvtActions);
    const rvtTableWrap = div('segmentpms-rvtlist guid-rvt-wrap');
    const { table: rvtTable, tbody: rvtBody, master: rvtMaster } = createRvtTable();
    rvtTableWrap.append(rvtTable);
    const rvtSummary = div('segmentpms-summary'); rvtSummary.textContent = '파일 0개';
    rvtSection.append(rvtHeader, rvtTableWrap, rvtSummary);
    body.append(rvtSection);

    // Result tabs
    const tabs = div('feature-results-panel guid-results feature-tabs');
    const tabHead = div('feature-results-head guid-results-head');
    const tabHeadLeft = div('guid-results-tabs');
    const tabHeadRight = div('guid-results-filters');
    const tabBtns = div('pill-tabs');
    const btnTabProject = document.createElement('button'); btnTabProject.type = 'button'; btnTabProject.className = 'pill-tab is-active'; btnTabProject.innerHTML = `<span class="pill-label">RVT 검토결과</span><span class="pill-count">0</span>`;
    const btnTabFamily = document.createElement('button'); btnTabFamily.type = 'button'; btnTabFamily.className = 'pill-tab'; btnTabFamily.innerHTML = `<span class="pill-label">Family 검토결과</span><span class="pill-count">0</span>`;
    tabBtns.append(btnTabProject, btnTabFamily);
    tabHeadLeft.append(tabBtns);
    const filterBar = buildFamilyFilter();
    filterBar.classList.add('is-hidden');
    tabHeadRight.append(filterBar);
    tabHead.append(tabHeadLeft, tabHeadRight);
    tabs.append(tabHead);

    const emptyState = div('guid-empty-state');
    emptyState.textContent = 'RVT 등록 후 검토 시작해주세요';
    tabs.append(emptyState);

    const tabPanels = div('guid-tab-panels');
    const tabPanelProject = div('guid-tab-panel');
    const projWrap = div('guid-detail-wrap');
    const projNavPane = div('guid-detail-nav feature-results-panel guid-scroll-box');
    const projNavList = document.createElement('ul'); projNavList.className = 'guid-nav-list';
    projNavPane.append(projNavList);
    const projDetailPane = div('guid-detail-pane');
    const projTableWrap = div('guid-table-wrap guid-scroll-box');
    const projTable = document.createElement('table'); projTable.className = 'guid-table';
    const projHead = document.createElement('thead');
    const projBody = document.createElement('tbody');
    projTable.append(projHead, projBody);
    projTableWrap.append(projTable);
    projDetailPane.append(projTableWrap);
    projWrap.append(projNavPane, projDetailPane);
    tabPanelProject.append(projWrap);

    const tabPanelFamily = div('guid-tab-panel is-hidden');
    const detailWrap = div('guid-detail-wrap');
    const navPane = div('guid-detail-nav feature-results-panel guid-scroll-box');
    const navList = document.createElement('ul'); navList.className = 'guid-nav-list';
    navPane.append(navList);
    const detailPane = div('guid-detail-pane');
    const familyUserSection = div('guid-section');
    const familyUserHeader = div('guid-section-header');
    const familyUserTitle = document.createElement('div'); familyUserTitle.className = 'guid-section-title'; familyUserTitle.textContent = '사용자 파라미터';
    const familyUserCount = document.createElement('span'); familyUserCount.className = 'guid-section-count'; familyUserCount.textContent = '0';
    familyUserHeader.append(familyUserTitle, familyUserCount);
    const familyEmpty = div('guid-empty'); familyEmpty.textContent = '파라미터 없음';
    const detailTableWrap = div('guid-table-wrap guid-scroll-box');
    const detailTable = document.createElement('table'); detailTable.className = 'guid-table';
    const detailHead = document.createElement('thead');
    const detailBody = document.createElement('tbody');
    detailTable.append(detailHead, detailBody);
    detailTableWrap.append(detailTable);
    familyUserSection.append(familyUserHeader, familyEmpty, detailTableWrap);

    const builtSection = div('guid-section');
    const builtHeader = div('guid-section-header');
    const builtTitle = document.createElement('div'); builtTitle.className = 'guid-section-title'; builtTitle.textContent = 'Built-in 파라미터';
    const builtCount = document.createElement('span'); builtCount.className = 'guid-section-count'; builtCount.textContent = '0';
    const builtToggle = document.createElement('button'); builtToggle.type = 'button'; builtToggle.className = 'pill-tab'; builtToggle.textContent = '펼치기';
    builtHeader.append(builtTitle, builtCount, builtToggle);
    const builtTableWrap = div('guid-table-wrap guid-scroll-box');
    const builtTable = document.createElement('table'); builtTable.className = 'guid-table';
    const builtHead = document.createElement('thead');
    const builtBody = document.createElement('tbody');
    builtTable.append(builtHead, builtBody);
    builtTableWrap.append(builtTable);
    builtSection.append(builtHeader, builtTableWrap);
    detailPane.append(familyUserSection, builtSection);

    builtToggle.onclick = () => {
        builtSection.classList.toggle('is-open');
        builtToggle.textContent = builtSection.classList.contains('is-open') ? '접기' : '펼치기';
        paintFamily();
    };
    detailWrap.append(navPane, detailPane);
    tabPanelFamily.append(detailWrap);

    tabPanels.append(tabPanelProject, tabPanelFamily);
    tabs.append(tabPanels);
    body.append(tabs);

    page.append(body);
    target.append(page);

    renderRvtList();
    syncTabState();
    syncRvtActionState();
    syncModeToggle();
    syncAnnotationToggle();
    syncResultState();
    requestSharedParamStatus();

    // Host events
    onHost('guid:files', ({ paths }) => {
        const list = Array.isArray(paths) ? paths : [];
        let added = 0;
        list.forEach(p => {
            const path = normalizeRvtPath(p);
            if (!path) return;
            const exists = state.rvtList.some(x => samePath(x, path));
            if (!exists) { state.rvtList.push(path); added++; }
            state.rvtChecked.add(path);
        });
        state.rvtList = dedupPaths(state.rvtList);
        if (added) {
            persistRvts();
            renderRvtList();
        } else {
            renderRvtList();
        }
        syncRvtActionState();
    });

    onHost('guid:progress', (payload) => {
        if (payload && payload.phase) {
            handleExcelProgress(payload);
        } else {
            handleRunProgress(payload);
        }
    });

    onHost('guid:done', (payload) => {
        ProgressDialog.hide();
        setBusy(false);
        lastExcelPct = 0;
        const proj = payload?.project || {};
        const famNav = payload?.family || payload?.familyIndex || {};
        state.runId = payload?.runId || '';
        state.hasRun = true;
        state.includeFamily = !!payload?.includeFamily;
        state.includeAnnotation = !!payload?.includeAnnotation;
        state.project = {
            columns: Array.isArray(proj.columns) ? proj.columns : [],
            rows: Array.isArray(proj.rows) ? proj.rows : []
        };
        state.familyNav = {
            columns: Array.isArray(famNav.columns) ? famNav.columns : [],
            rows: Array.isArray(famNav.rows) ? famNav.rows : []
        };
        state.familyDetail = { columns: [], rows: [] };
        state.activeProjectDoc = '';
        state.activeFamilyDoc = '';
        state.activeFamily = '';
        state.familyFilter = 'all';
        syncModeToggle();
        syncAnnotationToggle();
        exportBtn.disabled = !hasRowsForExport();
        updateTabCounts();
        paintProject();
        paintFamily();
        syncTabState();
        syncResultState();
        toast('검토 완료', 'ok');
    });

    onHost('guid:family-detail', (payload) => {
        if (payload?.runId && state.runId && payload.runId !== state.runId) return;
        state.familyDetail = {
            columns: Array.isArray(payload?.columns) ? payload.columns : [],
            rows: Array.isArray(payload?.rows) ? payload.rows : []
        };
        state.activeFamilyDoc = payload?.rvtPath || state.activeFamilyDoc || '';
        state.activeFamily = payload?.familyName || state.activeFamily || '';
        updateTabCounts();
        paintFamily();
        syncTabState();
        syncResultState();
    });

    onHost('guid:warn', ({ message }) => {
        if (message) toast(message, 'warn');
    });

    onHost('guid:exported', ({ path }) => {
        ProgressDialog.hide();
        setBusy(false);
        lastExcelPct = 0;
        if (path) {
            showExcelSavedDialog('엑셀로 내보냈습니다.', path, (p) => post('excel:open', { path: p }));
        } else {
            toast('엑셀 내보내기 완료', 'ok');
        }
    });

    const handleError = ({ message }) => {
        ProgressDialog.hide();
        setBusy(false);
        lastExcelPct = 0;
        if (message) toast(message, 'err');
    };
    onHost('guid:error', handleError);
    onHost('revit:error', handleError);
    onHost('host:error', handleError);
    onHost('sharedparam:status', (payload) => {
        state.sharedParamStatus = payload || {};
        updateSharedParamStatus(sharedStatus, state.sharedParamStatus);
    });

    // UI handlers
    btnTabProject.onclick = () => { state.activeTab = 'project'; syncTabState(); };
    btnTabFamily.onclick = () => {
        if (!state.includeFamily) return;
        state.activeTab = 'family';
        syncTabState();
    };

    function onRun() {
        if (state.busy) return;
        if (!canRunWithSharedParam()) return;
        state.rvtList = dedupPaths(state.rvtList);
        const targets = dedupPaths(state.rvtList.filter(p => state.rvtChecked.has(p)));
        if (state.rvtList.length > 0 && targets.length === 0) {
            toast('선택된 RVT가 없습니다.', 'warn');
            return;
        }

        const includeFamily = !!state.includeFamily;
        const payload = {
            mode: includeFamily ? 2 : 1,
            rvtPaths: state.rvtList.length === 0 ? [] : targets,
            includeFamily,
            includeAnnotation: includeFamily ? !!state.includeAnnotation : false
        };
        persistRvts();

        state.familyNav = { columns: [], rows: [] };
        state.familyDetail = { columns: [], rows: [] };
        state.project = { columns: [], rows: [] };
        state.activeProjectDoc = '';
        state.activeFamilyDoc = '';
        state.activeFamily = '';
        setBusy(true);
        state.runId = '';
        state.activeTab = 'project';
        ProgressDialog.show('GUID Audit', '준비 중…');
        post('guid:run', payload);
    }

    function requestSharedParamStatus() {
        post('sharedparam:status', { source: 'guid' });
    }

    function canRunWithSharedParam() {
        const status = state.sharedParamStatus || {};
        if (!status.status || status.status === 'ok') return true;
        const msg = status.warning || 'Shared Parameter 파일 상태가 올바르지 않습니다.';
        toast(msg, 'err');
        return false;
    }

    function onExport() {
        if (state.busy) return;
        if (!hasRowsForExport()) { toast('저장할 결과가 없습니다.', 'warn'); return; }
        let which = 'project';
        if (state.includeFamily) which = 'all';
        chooseExcelMode((mode) => {
            const excelMode = mode || 'fast';
            lastExcelPct = 0;
            setBusy(true);
            ProgressDialog.show('엑셀 내보내기', '엑셀 파일을 만드는 중…');
            post('guid:export', { which, excelMode });
        });
    }

    function buildModeToggle() {
        const wrap = div('guid-mode');
        const projBadge = document.createElement('span');
        projBadge.className = 'pill-tab guid-badge guid-mode-pill';
        projBadge.innerHTML = `<span class="guid-badge-check">✓</span><span>Project(RVT) Parameter - 기본</span>`;

        const famToggle = document.createElement('button');
        famToggle.type = 'button';
        famToggle.className = 'pill-tab guid-option-toggle';
        famToggle.setAttribute('aria-pressed', 'false');
        const sync = () => {
            const on = !!state.includeFamily;
            famToggle.classList.toggle('is-active', on);
            famToggle.setAttribute('aria-pressed', on ? 'true' : 'false');
            famToggle.innerHTML = `<span>Family 검토결과 추가 검토</span><span class="guid-option-badge">${on ? 'ON' : 'OFF'}</span>`;
        };
        famToggle.onclick = () => { state.includeFamily = !state.includeFamily; syncAnnotationToggle(); syncTabState(); sync(); };
        sync();

        wrap.append(projBadge, famToggle);
        wrap.sync = sync;
        return wrap;
    }

    function buildAnnotationToggle() {
        const wrap = div('guid-annotation');
        const label = document.createElement('label');
        label.className = 'guid-checkbox-row';
        const ck = document.createElement('input');
        ck.type = 'checkbox';
        ck.checked = !!state.includeAnnotation;
        ck.onchange = () => { state.includeAnnotation = !!ck.checked; };
        const span = document.createElement('span');
        span.textContent = 'Annotation 포함';
        label.append(ck, span);
        wrap.append(label);
        wrap.sync = function () {
            ck.checked = !!state.includeAnnotation;
            ck.disabled = !state.includeFamily;
            label.classList.toggle('is-disabled', !state.includeFamily);
            if (!state.includeFamily) ck.checked = false;
        };
        return wrap;
    }

    function buildSharedParamStatus() {
        const wrap = div('sharedparam-status');
        wrap.innerHTML = `
          <div class="sharedparam-status__head">
            <span class="sharedparam-status__title">Shared Parameter 상태</span>
            <span class="sharedparam-status__badge chip">조회 중</span>
          </div>
          <div class="sharedparam-status__body">
            <div class="sharedparam-status__row">
              <span class="sharedparam-status__label">경로</span>
              <span class="sharedparam-status__value" data-sp-path>조회 중</span>
            </div>
            <div class="sharedparam-status__row">
              <span class="sharedparam-status__label">파일 존재</span>
              <span class="sharedparam-status__value" data-sp-exists>—</span>
            </div>
            <div class="sharedparam-status__row">
              <span class="sharedparam-status__label">파일 열기</span>
              <span class="sharedparam-status__value" data-sp-open>—</span>
            </div>
          </div>
          <div class="sharedparam-status__hint" data-sp-hint></div>`;
        return wrap;
    }

    function updateSharedParamStatus(container, payload) {
        if (!container) return;
        const badge = container.querySelector('.sharedparam-status__badge');
        const pathEl = container.querySelector('[data-sp-path]');
        const existsEl = container.querySelector('[data-sp-exists]');
        const openEl = container.querySelector('[data-sp-open]');
        const hintEl = container.querySelector('[data-sp-hint]');

        const status = payload?.status || 'unknown';
        const label = payload?.statusLabel || '알 수 없음';
        const path = payload?.path || '미설정';
        const exists = payload?.existsOnDisk ? '존재' : '없음';
        const canOpen = payload?.canOpen ? 'OK' : '실패';
        const warning = payload?.warning || payload?.errorMessage || '';

        if (pathEl) pathEl.textContent = path;
        if (existsEl) existsEl.textContent = payload?.isSet ? exists : '미설정';
        if (openEl) openEl.textContent = payload?.isSet ? canOpen : '미설정';
        if (badge) {
            badge.textContent = label;
            badge.classList.remove('sharedparam-badge--ok', 'sharedparam-badge--warn', 'sharedparam-badge--error');
            if (status === 'ok') badge.classList.add('sharedparam-badge--ok');
            else if (status === 'unset' || status === 'missing') badge.classList.add('sharedparam-badge--warn');
            else badge.classList.add('sharedparam-badge--error');
        }
        if (hintEl) {
            hintEl.textContent = warning ? `이 상태에서는 검토가 실패할 수 있습니다. ${warning}` : '';
            hintEl.style.display = warning ? 'block' : 'none';
        }
    }

    function buildFamilyFilter() {
        const bar = div('guid-filter-bar');
        const btnAll = document.createElement('button'); btnAll.type = 'button'; btnAll.className = 'pill-tab is-active'; btnAll.textContent = '전체';
        const btnShared = document.createElement('button'); btnShared.type = 'button'; btnShared.className = 'pill-tab'; btnShared.textContent = 'Shared';
        const btnFamily = document.createElement('button'); btnFamily.type = 'button'; btnFamily.className = 'pill-tab'; btnFamily.textContent = 'Family';
        const sync = () => {
            btnAll.classList.toggle('is-active', state.familyFilter === 'all');
            btnShared.classList.toggle('is-active', state.familyFilter === 'shared');
            btnFamily.classList.toggle('is-active', state.familyFilter === 'family');
        };
        btnAll.onclick = () => { state.familyFilter = 'all'; sync(); paintFamily(); };
        btnShared.onclick = () => { state.familyFilter = 'shared'; sync(); paintFamily(); };
        btnFamily.onclick = () => { state.familyFilter = 'family'; sync(); paintFamily(); };
        bar.append(btnAll, btnShared, btnFamily);
        bar.sync = sync;
        return bar;
    }

    function renderRvtList() {
        state.rvtList = dedupPaths(state.rvtList);
        state.rvtChecked = new Set(state.rvtList.filter(p => state.rvtChecked.has(p)));
        const allChecked = state.rvtList.length > 0 && state.rvtList.every(p => state.rvtChecked.has(p));
        rvtMaster.checked = allChecked;
        rvtMaster.indeterminate = state.rvtList.length > 0 && !allChecked && state.rvtChecked.size > 0;
        rvtMaster.disabled = state.rvtList.length === 0;
        rvtMaster.onchange = () => {
            if (rvtMaster.checked) state.rvtChecked = new Set(state.rvtList);
            else state.rvtChecked.clear();
            persistRvts();
            renderRvtList();
        };
        const rows = state.rvtList.map((p, i) => ({
            checked: state.rvtChecked.has(p),
            index: i + 1,
            name: getRvtName(p),
            path: p,
            title: p,
            onToggle: (checked) => {
                if (checked) state.rvtChecked.add(p); else state.rvtChecked.delete(p);
                persistRvts();
                renderRvtList();
            }
        }));
        renderRvtRows(rvtBody, rows);
        rvtSummary.textContent = state.rvtList.length ? `파일 ${state.rvtList.length}개` : '파일 0개';
        syncRvtActionState();
    }

    function paintProject() {
        projNavList.innerHTML = '';
        if (!state.project.rows.length) {
            const empty = document.createElement('li');
            empty.className = 'guid-nav-empty';
            empty.textContent = '결과가 없습니다.';
            projNavList.append(empty);
            projHead.innerHTML = '';
            projBody.innerHTML = '';
            return;
        }
        const idxPath = projectCol('RvtPath');
        const idxName = projectCol('RvtName');
        const docMap = new Map();
        state.project.rows.forEach(row => {
            const path = idxPath >= 0 ? safe(row[idxPath]) : '';
            const name = idxName >= 0 ? safe(row[idxName]) : (path || '(Doc)');
            const key = path || name;
            if (!docMap.has(key)) docMap.set(key, name);
        });
        const docs = Array.from(docMap.entries()).sort((a, b) => a[1].localeCompare(b[1]));
        if (!state.activeProjectDoc && docs.length) state.activeProjectDoc = docs[0][0];
        docs.forEach(([key, name]) => {
            const li = document.createElement('li');
            const btn = document.createElement('button'); btn.type = 'button'; btn.className = 'nav-fam-item'; btn.textContent = name;
            btn.onclick = () => { state.activeProjectDoc = key; paintProject(); };
            if (state.activeProjectDoc === key) btn.classList.add('is-active');
            li.append(btn);
            projNavList.append(li);
        });
        buildHead(projHead, state.project.columns, new Set(['RvtPath']));
        paintVirtualRows(projBody, state.project.columns, filteredProjectRows(), new Set(['RvtPath']));
    }

    function paintFamily() {
        navList.innerHTML = '';
        if (!state.includeFamily) {
            navList.innerHTML = '<li class="guid-nav-empty">Family 검토결과 추가 검토를 선택 후 실행하세요.</li>';
            detailHead.innerHTML = '';
            detailBody.innerHTML = '';
            builtHead.innerHTML = '';
            builtBody.innerHTML = '';
            familyEmpty.style.display = 'none';
            detailTableWrap.style.display = 'none';
            builtTableWrap.style.display = 'none';
            if (filterBar && typeof filterBar.sync === 'function') filterBar.sync();
            return;
        }
        if (!state.familyNav.rows.length) {
            const empty = document.createElement('li');
            empty.className = 'guid-nav-empty';
            empty.textContent = '패밀리 결과가 없습니다.';
            navList.append(empty);
            detailHead.innerHTML = '';
            detailBody.innerHTML = '';
            builtHead.innerHTML = '';
            builtBody.innerHTML = '';
            familyEmpty.style.display = 'none';
            detailTableWrap.style.display = 'none';
            builtTableWrap.style.display = 'none';
            if (filterBar && typeof filterBar.sync === 'function') filterBar.sync();
            return;
        }
        if (filterBar && typeof filterBar.sync === 'function') filterBar.sync();
        const idxPath = indexCol(state.familyNav, 'RvtPath');
        const idxName = indexCol(state.familyNav, 'RvtName');
        const idxFam = indexCol(state.familyNav, 'FamilyName');
        const map = new Map();
        state.familyNav.rows.forEach(row => {
            const path = idxPath >= 0 ? safe(row[idxPath]) : '';
            const rname = idxName >= 0 ? safe(row[idxName]) : (path || '(Doc)');
            const fam = idxFam >= 0 ? safe(row[idxFam]) : '';
            const key = path || rname;
            if (!map.has(key)) map.set(key, { name: rname, families: new Set() });
            if (fam) map.get(key).families.add(fam);
        });
        const entries = Array.from(map.entries()).sort((a, b) => a[1].name.localeCompare(b[1].name));
        if (!state.activeFamilyDoc && entries.length) state.activeFamilyDoc = entries[0][0];
        entries.forEach(([key, info]) => {
            const docItem = document.createElement('li');
            docItem.className = 'guid-nav-doc';
            const docTitle = document.createElement('div');
            docTitle.className = 'nav-doc-title';
            docTitle.textContent = info.name;
            docItem.append(docTitle);

            const famList = document.createElement('ul');
            famList.className = 'guid-nav-fams';

            Array.from(info.families).sort((a, b) => a.localeCompare(b)).forEach(f => {
                const li = document.createElement('li');
                const btn = document.createElement('button'); btn.type = 'button'; btn.className = 'nav-fam-item'; btn.textContent = f;
                btn.title = f;
                btn.onclick = () => { onRequestFamilyDetail(key, f); };
                if (state.activeFamilyDoc === key && state.activeFamily === f) btn.classList.add('is-active');
                li.append(btn);
                famList.append(li);
            });

            docItem.append(famList);
            navList.append(docItem);
        });
        const { userRows, builtRows } = splitFamilyRows();
        familyUserCount.textContent = String(userRows.length || 0);
        builtCount.textContent = String(builtRows.length || 0);
        familyEmpty.style.display = userRows.length ? 'none' : 'block';
        detailTableWrap.style.display = userRows.length ? 'block' : 'none';
        buildHead(detailHead, state.familyDetail.columns, HIDDEN_DETAIL_COLS);
        paintVirtualRows(detailBody, state.familyDetail.columns, userRows, HIDDEN_DETAIL_COLS);

        const builtOpen = builtSection.classList.contains('is-open');
        builtTableWrap.style.display = builtOpen ? 'block' : 'none';
        if (builtOpen) {
            buildHead(builtHead, state.familyDetail.columns, HIDDEN_DETAIL_COLS);
            paintVirtualRows(builtBody, state.familyDetail.columns, builtRows, HIDDEN_DETAIL_COLS);
        } else {
            builtHead.innerHTML = '';
            builtBody.innerHTML = '';
        }
    }

    function onRequestFamilyDetail(rvtPath, familyName) {
        if (!state.runId) {
            toast('먼저 검토를 실행하세요.', 'warn');
            return;
        }
        state.activeFamilyDoc = rvtPath || familyName || '';
        state.activeFamily = familyName || '';
        state.familyDetail = { columns: [], rows: [] };
        paintFamily();
        post('guid:request-family-detail', { runId: state.runId, rvtPath, familyName });
    }

    function filteredProjectRows() {
        if (!state.project.rows.length) return [];
        const idxPath = projectCol('RvtPath');
        const idxName = projectCol('RvtName');
        if (!state.activeProjectDoc) return state.project.rows;
        return state.project.rows.filter(row => {
            const path = idxPath >= 0 ? safe(row[idxPath]) : '';
            const name = idxName >= 0 ? safe(row[idxName]) : '';
            return (path || name) === state.activeProjectDoc;
        });
    }

    function filteredFamilyRows() {
        if (!state.familyDetail.rows.length) return [];
        const idxPath = colIndex('RvtPath');
        const idxFam = colIndex('FamilyName');
        const idxKind = colIndex('ParamKind');
        const idxShared = colIndex('IsShared');
        return state.familyDetail.rows.filter(row => {
            const path = idxPath >= 0 ? safe(row[idxPath]) : '';
            const famName = idxFam >= 0 ? safe(row[idxFam]) : '';
            const kind = idxKind >= 0 ? safe(row[idxKind]) : '';
            const shared = idxShared >= 0 ? safe(row[idxShared]) : '';
            const docMatch = idxPath < 0 || !state.activeFamilyDoc || ((path || '') === state.activeFamilyDoc);
            const famMatch = !state.activeFamily || (famName === state.activeFamily);
            const filterMatch = state.familyFilter === 'all' ||
                (state.familyFilter === 'shared' && (kind === 'Shared' || shared === 'Y')) ||
                (state.familyFilter === 'family' && (kind === 'Family' || shared === 'N'));
            return docMatch && famMatch && filterMatch;
        });
    }

    function splitFamilyRows() {
        const rows = filteredFamilyRows();
        const idxKind = colIndex('ParamKind');
        const userRows = [];
        const builtRows = [];
        rows.forEach(row => {
            const kind = idxKind >= 0 ? safe(row[idxKind]) : '';
            if (kind === 'BuiltIn') {
                builtRows.push(row);
            } else if (kind === 'Shared' || kind === 'Family') {
                userRows.push(row);
            } else {
                userRows.push(row);
            }
        });
        return { userRows, builtRows };
    }

    function colIndex(name) {
        return state.familyDetail.columns.findIndex(c => c === name);
    }

    function projectCol(name) {
        return state.project.columns.findIndex(c => c === name);
    }

    function indexCol(table, name) {
        return (table.columns || []).findIndex(c => c === name);
    }

    function buildHead(thead, columns, hidden) {
        thead.innerHTML = '';
        const tr = document.createElement('tr');
        columns.forEach(c => {
            if (hidden.has(c)) return;
            const th = document.createElement('th');
            th.textContent = c;
            tr.append(th);
        });
        thead.append(tr);
    }

    function paintVirtualRows(tbody, columns, rows, hidden) {
        tbody.innerHTML = '';
        let idx = 0;
        const chunk = () => {
            const frag = document.createDocumentFragment();
            for (let i = 0; i < 200 && idx < rows.length; i++, idx++) {
                const row = rows[idx];
                const tr = document.createElement('tr');
                columns.forEach((c, ci) => {
                    if (hidden.has(c)) return;
                    const td = document.createElement('td');
                    const text = safe(row[ci]);
                    td.textContent = text;
                    td.title = text;
                    tr.append(td);
                });
                frag.append(tr);
            }
            tbody.append(frag);
            if (idx < rows.length) setTimeout(chunk, 0);
        };
        chunk();
    }

    function hasRowsForExport() {
        const hasProject = (state.project.columns || []).length > 0;
        const hasFamily = state.includeFamily && ((state.familyDetail.columns || []).length > 0 || (state.familyNav.columns || []).length > 0);
        return hasProject || hasFamily;
    }

    function syncTabState() {
        btnTabProject.classList.toggle('is-active', state.activeTab === 'project');
        btnTabFamily.classList.toggle('is-active', state.activeTab === 'family');
        btnTabFamily.disabled = !state.includeFamily;
        tabPanelProject.classList.toggle('is-hidden', state.activeTab !== 'project');
        tabPanelFamily.classList.toggle('is-hidden', state.activeTab !== 'family');
        filterBar.classList.toggle('is-hidden', state.activeTab !== 'family');
        if (state.activeTab === 'family' && !state.includeFamily) {
            state.activeTab = 'project';
            tabPanelProject.classList.remove('is-hidden');
            tabPanelFamily.classList.add('is-hidden');
            filterBar.classList.add('is-hidden');
        }
        exportBtn.disabled = !hasRowsForExport();
        updateTabCounts();
    }

    function syncResultState() {
        const showResults = state.hasRun;
        tabPanels.style.display = showResults ? '' : 'none';
        emptyState.style.display = showResults ? 'none' : 'flex';
    }

    function syncAnnotationToggle() {
        if (annotationToggle && typeof annotationToggle.sync === 'function') {
            annotationToggle.sync();
        }
    }

    function syncModeToggle() {
        if (modeToggle && typeof modeToggle.sync === 'function') {
            modeToggle.sync();
        }
    }

    function setBusy(on) {
        state.busy = on;
        runBtn.disabled = on;
        exportBtn.disabled = on || !hasRowsForExport();
    }

    function persistRvts() {
        state.rvtList = dedupPaths(state.rvtList);
        try { localStorage.setItem(LS_RVTS, JSON.stringify(state.rvtList || [])); } catch { }
    }

    function loadRvtList() {
        try {
            const raw = localStorage.getItem(LS_RVTS);
            const arr = JSON.parse(raw || '[]');
            if (Array.isArray(arr)) return dedupPaths(arr);
        } catch { }
        return [];
    }

    function handleRunProgress(payload) {
        const percent = typeof payload?.pct === 'number' ? payload.pct : 0;
        const message = payload?.text || '';
        if (!state.busy && percent <= 0) return;
        if (!state.busy) setBusy(true);
        ProgressDialog.show('GUID Audit', message || '진행 중…');
        ProgressDialog.update(percent, message || '', '');
    }

    function handleExcelProgress(payload) {
        const phase = normalizeExcelPhase(payload?.phase);
        const total = Number(payload?.total) || 0;
        const current = Number(payload?.current) || 0;
        const percent = computeExcelPercent(phase, current, total, payload?.phaseProgress);
        const subtitle = buildExcelSubtitle(phase, current, total);
        const detail = formatExcelDetail(phase, payload?.message);

        const exporting = phase !== 'DONE' && phase !== 'ERROR';
        if (!state.busy && exporting) setBusy(true);

        ProgressDialog.show('엑셀 내보내기', subtitle || '엑셀 내보내기 중…');
        ProgressDialog.update(percent, subtitle, detail);

        if (!exporting) {
            setTimeout(() => { ProgressDialog.hide(); lastExcelPct = 0; setBusy(false); }, 260);
        }
    }

    function normalizeExcelPhase(phase) {
        return String(phase || '').trim().toUpperCase() || 'EXCEL_WRITE';
    }

    function computeExcelPercent(phase, current, total, phaseProgress) {
        const norm = normalizeExcelPhase(phase);
        if (norm === 'DONE') { lastExcelPct = 100; return 100; }
        if (norm === 'ERROR') return lastExcelPct;

        const completed = ['EXCEL_INIT', 'EXCEL_WRITE', 'EXCEL_SAVE', 'AUTOFIT'].reduce((acc, key) => {
            if (key === norm) return acc;
            return acc + (EXCEL_PHASE_WEIGHT[key] || 0);
        }, 0);
        const weight = EXCEL_PHASE_WEIGHT[norm] || 0;
        const ratio = total > 0 ? Math.max(0, Math.min(1, current / total)) : 0;
        const staged = Math.max(ratio, clamp01(phaseProgress));
        const pct = (completed + weight * staged) / (completed + weight) * 100;
        lastExcelPct = Math.max(lastExcelPct, Math.min(100, pct));
        return lastExcelPct;
    }

    function clamp01(v) { const n = Number(v); if (Number.isFinite(n)) return Math.max(0, Math.min(1, n)); return 0; }

    function buildExcelSubtitle(phase, current, total) {
        const norm = normalizeExcelPhase(phase);
        switch (norm) {
            case 'EXCEL_INIT': return '엑셀 워크북 준비 중';
            case 'EXCEL_WRITE': return `엑셀 데이터 작성 중 (${current}/${Math.max(total, current || 1)})`;
            case 'EXCEL_SAVE': return '엑셀 내보내기 중';
            case 'AUTOFIT': return '열 너비 자동 조정 중…';
            case 'DONE': return '엑셀 내보내기 완료';
            case 'ERROR': return '엑셀 내보내기 오류';
            default: return '엑셀 내보내기 중…';
        }
    }

    function formatExcelDetail(phase, message) {
        const norm = normalizeExcelPhase(phase);
        if (norm === 'AUTOFIT') return '열 너비 자동 조정 중…';
        return message || '';
    }

    function onRemoveSelected() {
        if (!state.rvtChecked.size) { toast('제거할 RVT를 선택하세요.', 'warn'); return; }
        state.rvtList = state.rvtList.filter(p => !state.rvtChecked.has(p));
        state.rvtChecked.clear();
        persistRvts();
        renderRvtList();
    }

    function syncRvtActionState() {
        if (btnRemove) btnRemove.disabled = state.rvtChecked.size === 0;
        if (btnClear) btnClear.disabled = state.rvtList.length === 0;
    }

    function updateTabCounts() {
        setTabCount(btnTabProject, state.project.rows.length || 0);
        const detCount = state.includeFamily ? (state.familyNav.rows.length || 0) : 0;
        setTabCount(btnTabFamily, detCount);
    }
}

function normalizeRvtPath(entry) {
    if (!entry) return '';
    if (typeof entry === 'string') return entry.trim();
    if (typeof entry === 'object') {
        if (typeof entry.path === 'string') return entry.path.trim();
        if (typeof entry.fullPath === 'string') return entry.fullPath.trim();
    }
    return '';
}

function dedupPaths(list) {
    const seen = new Set();
    const clean = [];
    (list || []).forEach(item => {
        const path = normalizeRvtPath(item);
        if (!path) return;
        const key = path.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        clean.push(path);
    });
    return clean;
}

function setTabCount(btn, count) {
    if (!btn) return;
    const badge = btn.querySelector('.pill-count');
    if (!badge) return;
    badge.textContent = Number.isFinite(count) ? count : 0;
}

function safe(v) {
    if (v === null || v === undefined) return '';
    return String(v);
}

function samePath(a, b) {
    if (!a || !b) return false;
    return a.toLowerCase() === b.toLowerCase();
}

function cardBtn(text, onClick) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'btn card-btn';
    btn.textContent = text;
    btn.onclick = onClick;
    return btn;
}
