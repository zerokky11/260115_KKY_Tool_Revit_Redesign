// Resources/HubUI/js/views/paramprop.js
import { clear, div, toast, setBusy, showExcelSavedDialog, debounce, chooseExcelMode } from '../core/dom.js';
import { ProgressDialog } from '../core/progress.js';
import { post, onHost } from '../core/bridge.js';

const DEFAULT_GUIDE = '공유 파라미터 연동을 실행하면 결과가 이곳에 표시됩니다.';
const PHASE_WEIGHT = {
    collect: 0.15,
    analyze: 0.25,
    apply: 0.3,
    save: 0.05,
    close: 0.05,
    excel_init: 0.05,
    excel_write: 0.12,
    excel_save: 0.02,
    autofit: 0.01,
    done: 0
};
const PHASE_ORDER = ['collect', 'analyze', 'apply', 'save', 'close', 'excel_init', 'excel_write', 'excel_save', 'autofit', 'done'];
const PHASE_LABEL = {
    collect: '스캔 중',
    analyze: '분석 중',
    apply: '적용 중',
    save: '저장 중',
    close: '마무리 중',
    excel_init: '엑셀 준비',
    excel_write: '엑셀 작성',
    excel_save: '파일 저장',
    autofit: 'AutoFit 적용',
    done: '완료'
};

export function renderParamProp(root) {
    const target = root || document.getElementById('view-root') || document.getElementById('app');
    clear(target);
    const top = document.querySelector('#topbar-root .topbar') || document.querySelector('.topbar'); if (top) top.classList.add('hub-topbar');

    const state = {
        defs: [],
        groups: [],
        targetGroups: [],
        selectedGroups: new Set(['(All Groups)']),
        selectedParams: new Set(),
        search: '',
        targetGroupId: null,
        isInstance: true,
        excludeDummy: true,
        lastReport: DEFAULT_GUIDE,
        lastDetails: [],
    };
    let lastProgressPct = 0;

    const page = div('paramprop-page feature-shell');

    const header = div('feature-header');
    const heading = div('feature-heading');
    heading.innerHTML = `
      <span class="feature-kicker">Shared Parameter Propagator</span>
      <h2 class="feature-title">공유 파라미터 연동</h2>
      <p class="feature-sub">복합 패밀리의 하위 패밀리에 공유 파라미터를 추가하고 연동/검증합니다.</p>`;

    const runBtn = cardBtn('연동 실행', onRun);
    runBtn.id = 'btnParamPropRun';
    const exportBtn = cardBtn('엑셀 내보내기', onExport);
    exportBtn.id = 'btnParamPropExport';
    exportBtn.disabled = true;

    const actions = div('feature-actions');
    actions.append(runBtn, exportBtn);
    header.append(heading, actions);
    page.append(header);

    const layout = div('paramprop-layout');
    page.append(layout);

    // ----- 선택 영역 -----
    const pickerCard = div('paramprop-card section paramprop-picker-card');
    pickerCard.innerHTML = '<div class="paramprop-title">공유 파라미터 선택</div>';

    const searchRow = div('paramprop-row paramprop-search-row');
    const searchBox = document.createElement('input');
    searchBox.type = 'search';
    searchBox.placeholder = '이름 또는 그룹 검색';
    searchBox.className = 'paramprop-search';
    searchBox.addEventListener('input', debounce((e) => {
        state.search = (e.target.value || '').trim();
        renderTable();
    }, 120));
    searchRow.append(labelSpan('검색'), searchBox);
    pickerCard.append(searchRow);

    const selectGrid = div('paramprop-grid');
    pickerCard.append(selectGrid);

    // 그룹 리스트
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

    // 파라미터 테이블
    const tableBox = div('paramprop-table-box');
    const tableHead = div('paramprop-subtitle');
    tableHead.textContent = '파라미터';
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

    // 옵션 영역
    const optionRow = div('paramprop-row paramprop-options');
    const dummyWrap = div('paramprop-opt');
    const dummyChk = document.createElement('input');
    dummyChk.type = 'checkbox';
    dummyChk.checked = true;
    dummyChk.id = 'optDummy';
    dummyChk.addEventListener('change', () => { state.excludeDummy = !!dummyChk.checked; });
    const dummyLbl = document.createElement('label');
    dummyLbl.setAttribute('for', 'optDummy');
    dummyLbl.textContent = "하위 패밀리 'Dummy' 포함 시 제외";
    dummyWrap.append(dummyChk, dummyLbl);

    const instWrap = div('paramprop-opt radios');
    instWrap.innerHTML = '<span class="opt-label">모드</span>';
    const rbInst = radio('param-scope', '인스턴스', true, () => { state.isInstance = true; });
    const rbType = radio('param-scope', '타입', false, () => { state.isInstance = false; });
    instWrap.append(rbInst, rbType);

    const targetGroupWrap = div('paramprop-opt paramprop-opt-group');
    const groupLbl = labelSpan('추가할 파라미터 그룹');
    const groupSel = document.createElement('select');
    groupSel.className = 'paramprop-select';
    groupSel.addEventListener('change', () => { state.targetGroupId = Number(groupSel.value); });
    targetGroupWrap.append(groupLbl, groupSel);

    optionRow.append(dummyWrap, instWrap, targetGroupWrap);
    pickerCard.append(optionRow);

    // ----- 결과 영역 -----
    const resultCard = div('paramprop-card section paramprop-report-card');
    const resultTitle = div('paramprop-title');
    resultTitle.textContent = '결과';
    const reportBox = document.createElement('pre');
    reportBox.className = 'paramprop-report';
    reportBox.textContent = DEFAULT_GUIDE;

    const filterRow = div('paramprop-row paramprop-filters');
    filterRow.innerHTML = '<span class="opt-label">Type 필터</span>';
    const filterAll = chip('전체', 'all');
    const filterScan = chip('ScanFail', 'ScanFail');
    const filterSkip = chip('Skip', 'Skip');
    const filterErr = chip('Error', 'Error');
    const filterChild = chip('ChildError', 'ChildError');
    const filters = [filterAll, filterScan, filterSkip, filterErr, filterChild];
    filterAll.classList.add('is-active');
    filterRow.append(filterAll, filterScan, filterSkip, filterErr, filterChild);

    const detailTable = document.createElement('table');
    detailTable.className = 'paramprop-detail';
    detailTable.innerHTML = '<thead><tr><th>Type</th><th>Family</th><th>Detail</th></tr></thead><tbody></tbody>';
    const detailWrap = div('paramprop-detail-wrap');
    detailWrap.append(detailTable);

    resultCard.append(resultTitle, reportBox, filterRow, detailWrap);

    layout.append(pickerCard, resultCard);
    target.append(page);

    // 이벤트/호스트 응답
    onHost('sharedparam:list', handleList);
    onHost('sharedparam:done', handleDone);
    onHost('sharedparam:exported', handleExported);
    onHost('paramprop:progress', handleProgress);
    onHost('revit:error', ({ message }) => { lastProgressPct = 0; ProgressDialog.hide(); setBusy(false); toast(message || 'Revit 오류가 발생했습니다.', 'err', 3200); });
    onHost('host:error', ({ message }) => { lastProgressPct = 0; ProgressDialog.hide(); setBusy(false); toast(message || '호스트 오류가 발생했습니다.', 'err', 3200); });

    fetchDefinitions();

    // -------- handlers --------
    function fetchDefinitions() {
        setBusy(true, '공유 파라미터 불러오는 중…');
        post('sharedparam:list', {});
    }

    function handleList(payload) {
        setBusy(false);
        const ok = payload?.ok !== false;
        if (!ok) {
            toast(payload?.message || '공유 파라미터 목록을 가져오지 못했습니다.', 'warn');
            state.defs = [];
            state.groups = [];
            renderGroups();
            renderTable();
            return;
        }

        if (Array.isArray(payload?.items)) {
            state.defs = payload.items.map((item) => ({
                groupName: item.groupName || '',
                name: item.name || '',
                paramType: item.dataTypeToken || '',
                visible: true
            }));
        } else {
            state.defs = Array.isArray(payload?.definitions) ? payload.definitions : [];
        }
        state.groups = deriveGroups(state.defs);
        state.targetGroups = Array.isArray(payload?.targetGroups) ? payload.targetGroups : [];
        state.selectedGroups = new Set(['(All Groups)']);
        state.selectedParams.clear();
        state.targetGroupId = state.targetGroups?.[0]?.id ?? null;
        syncTargetGroupSelect();
        renderGroups();
        renderTable();
    }

    function handleDone({ ok, status, message, report, details, rows }) {
        setBusy(false);
        const good = ok !== false && String(status || '').toLowerCase() !== 'failed';
        const text = report || DEFAULT_GUIDE;
        state.lastReport = text;
        state.lastDetails = Array.isArray(rows) ? rows : (Array.isArray(details) ? details : []);
        paintReport(text);
        paintDetails(getActiveFilterKey());
        exportBtn.disabled = state.lastDetails.length === 0;

        const msg = message || (good ? '공유 파라미터 연동을 완료했습니다.' : '공유 파라미터 연동이 실패했습니다.');
        toast(msg, good ? 'ok' : (status === 'cancelled' ? 'info' : 'err'), 2800);
        ProgressDialog.update(100, '완료', '공유 파라미터 추가 연동이 완료되었습니다.');
        setTimeout(() => { lastProgressPct = 0; ProgressDialog.hide(); }, 400);
    }

    function handleExported({ ok, path, message }) {
        setBusy(false);
        exportBtn.disabled = state.lastDetails.length === 0;
        if (ok && path) {
            showExcelSavedDialog('엑셀 내보내기가 완료되었습니다.', path, (p) => post('excel:open', { path: p }));
        } else {
            toast(message || '엑셀 내보내기에 실패했습니다.', 'err');
        }
    }

    function onRun() {
        const selected = Array.from(state.selectedParams);
        if (!selected.length) {
            toast('하나 이상의 파라미터를 선택하세요.', 'warn');
            return;
        }
        lastProgressPct = 0;
        setBusy(true, '공유 파라미터 연동 중…');
        ProgressDialog.show('공유 파라미터 추가 연동', '준비 중...');
        ProgressDialog.update(0, '준비 중...', '');
        exportBtn.disabled = true;
        const payload = {
            paramNames: selected,
            group: state.targetGroupId,
            isInstance: !!state.isInstance,
            excludeDummy: !!state.excludeDummy
        };
        post('sharedparam:run', payload);
    }

    function onExport() {
        if (!state.lastDetails.length) {
            toast('최근 연동 결과가 없습니다.', 'info');
            return;
        }
        setBusy(true, '엑셀 내보내기 중…');
        chooseExcelMode((mode) => post('sharedparam:export-excel', { excelMode: mode || 'fast' }));
    }

    // ----- 렌더링 -----
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
        renderTable();
    }

    function renderTable() {
        tbody.innerHTML = '';
        const filtered = filterDefs();
        if (!filtered.length) {
            const tr = document.createElement('tr');
            const tdEmpty = document.createElement('td');
            tdEmpty.colSpan = 2;
            tdEmpty.textContent = state.defs.length
                ? '조건에 맞는 항목이 없습니다.'
                : 'Shared Parameter 등록이 필요합니다. Revit에서 Shared Parameter Text를 등록/연결 후 다시 시도하세요.';
            tdEmpty.className = 'paramprop-empty';
            tr.append(tdEmpty);
            tbody.append(tr);
            return;
        }

        filtered.forEach(def => {
            const key = def.name;
            const tr = document.createElement('tr');
            tr.dataset.key = key;
            tr.dataset.type = def.paramType || '';
            tr.dataset.visible = def.visible ? 'true' : 'false';
            tr.dataset.group = def.groupName || '';
            tr.className = state.selectedParams.has(key) ? 'is-selected' : '';
            const tdChk = document.createElement('td');
            const chk = document.createElement('input');
            chk.type = 'checkbox';
            chk.checked = state.selectedParams.has(key);
            chk.addEventListener('change', (e) => {
                e.stopPropagation();
                if (chk.checked) state.selectedParams.add(key); else state.selectedParams.delete(key);
                renderTable();
            });
            tdChk.append(chk);
            const nameCell = td(def.name);
            nameCell.title = `${def.groupName || ''} • ${def.paramType || ''} • ${def.visible ? 'Visible' : 'Hidden'}`.trim();
            tr.append(tdChk, nameCell);
            tr.addEventListener('click', () => {
                if (state.selectedParams.has(key)) state.selectedParams.delete(key); else state.selectedParams.add(key);
                renderTable();
            });
            tbody.append(tr);
        });
    }

    function syncTargetGroupSelect() {
        groupSel.innerHTML = '';
        const opts = state.targetGroups || [];
        opts.forEach(o => {
            const opt = document.createElement('option');
            opt.value = o.id;
            opt.textContent = o.name;
            groupSel.append(opt);
        });
        if (opts.length) {
            groupSel.value = opts[0].id;
            state.targetGroupId = opts[0].id;
        }
    }

    function filterDefs() {
        const groups = state.selectedGroups;
        const search = state.search.toLowerCase();
        return state.defs.filter(d => {
            const inGroup = groups.has('(All Groups)') || groups.has(d.groupName);
            if (!inGroup) return false;
            if (!search) return true;
            return (d.name || '').toLowerCase().includes(search) || (d.groupName || '').toLowerCase().includes(search);
        });
    }

    function deriveGroups(defs) {
        const set = new Set();
        (Array.isArray(defs) ? defs : []).forEach(d => {
            if (d?.groupName) set.add(d.groupName);
        });
        return Array.from(set).sort((a, b) => a.localeCompare(b, 'ko')); 
    }

    function paintReport(text) {
        reportBox.textContent = text || DEFAULT_GUIDE;
    }

    function paintDetails(kind) {
        const body = detailTable.querySelector('tbody');
        body.innerHTML = '';
        const rows = (state.lastDetails || []).filter(r => {
            const t = r.type || r.kind || '';
            return kind === 'all' || t === kind;
        });
        if (!rows.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 3;
            td.textContent = state.lastDetails.length ? '필터 결과가 없습니다.' : '결과가 없습니다.';
            tr.append(td);
            body.append(tr);
            return;
        }
        rows.forEach(r => {
            const tr = document.createElement('tr');
            tr.append(td(r.type || r.kind), td(r.family), td(r.detail));
            body.append(tr);
        });
    }

    filters.forEach(btn => {
        btn.addEventListener('click', () => {
            filters.forEach(b => b.classList.remove('is-active'));
            btn.classList.add('is-active');
            const key = btn.dataset.key || 'all';
            paintDetails(key);
        });
    });

    function getActiveFilterKey() {
        return filters.find(f => f.classList.contains('is-active'))?.dataset.key || 'all';
    }

    function labelSpan(text) {
        const s = document.createElement('span');
        s.className = 'opt-label';
        s.textContent = text;
        return s;
    }

    function radio(name, label, checked, onChange) {
        const wrap = document.createElement('label');
        wrap.className = 'paramprop-radio';
        const rb = document.createElement('input');
        rb.type = 'radio';
        rb.name = name;
        rb.checked = checked;
        rb.addEventListener('change', () => { if (rb.checked && typeof onChange === 'function') onChange(); });
        const span = document.createElement('span');
        span.textContent = label;
        wrap.append(rb, span);
        return wrap;
    }

    function chip(label, key) {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'paramprop-chip-btn';
        btn.textContent = label;
        btn.dataset.key = key;
        return btn;
    }

    function td(text) {
        const cell = document.createElement('td');
        cell.textContent = text == null ? '' : text;
        return cell;
    }

    function cardBtn(label, onClick) {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'btn btn-primary';
        btn.textContent = label;
        btn.onclick = onClick;
        return btn;
    }

    function handleProgress(payload) {
        if (!payload) { ProgressDialog.hide(); return; }
        const phase = String(payload.phase || payload.Stage || '').toLowerCase();
        const phaseProgress = clamp01(payload.phaseProgress);
        const current = Number(payload.current || payload.Index || 0) || 0;
        const total = Number(payload.total || payload.Total || 0) || 0;
        const message = payload.message || payload.Message || '';
        const target = payload.target || payload.Target || '';

        const basePct = computeBasePercent(phase);
        const weight = PHASE_WEIGHT[phase] || 0;
        const rawPct = phase === 'done' ? 100 : Math.min(100, Math.max(0, (basePct + weight * phaseProgress) * 100));
        const percent = Math.max(lastProgressPct, rawPct);
        lastProgressPct = percent;

        const subtitle = buildSubtitle(phase, current, total);
        const detail = buildDetail(message, target);

        ProgressDialog.show('공유 파라미터 추가 연동', subtitle);
        ProgressDialog.update(percent, subtitle, detail);

        if (phase === 'done') {
            setTimeout(() => { lastProgressPct = 0; ProgressDialog.hide(); }, 450);
        }
    }

    function clamp01(v) {
        const n = Number(v);
        if (!Number.isFinite(n)) return 0;
        return Math.max(0, Math.min(1, n));
    }

    function computeBasePercent(phase) {
        const p = String(phase || '').toLowerCase();
        let acc = 0;
        for (const key of PHASE_ORDER) {
            if (key === p) break;
            acc += PHASE_WEIGHT[key] || 0;
        }
        return acc;
    }

    function buildSubtitle(phase, current, total) {
        const label = PHASE_LABEL[phase] || '진행 중';
        const count = total > 0 ? ` (${Math.max(current, 0)}/${total})` : '';
        return `${label}${count}`;
    }

    function buildDetail(message, target) {
        const parts = [];
        if (message) parts.push(message);
        if (target) parts.push(target);
        return parts.join(' · ');
    }
}
