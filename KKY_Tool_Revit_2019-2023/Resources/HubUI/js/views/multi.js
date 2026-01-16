import { clear, div, toast, setBusy, showExcelSavedDialog, chooseExcelMode } from '../core/dom.js';
import { ProgressDialog } from '../core/progress.js';
import { post, onHost } from '../core/bridge.js';
import { createRvtTable, renderRvtRows, getRvtName } from './rvtTable.js';

const FEATURE_META = {
  connector: { label: '커넥터 진단', desc: 'Parameter 값 연속성 검토', requiresSharedParams: false },
  guid: { label: 'GUID 검토', desc: '공유 파라미터 GUID 불일치 검토', requiresSharedParams: true },
  points: { label: 'Point 추출', desc: 'Project/Survey Point 좌표 추출', requiresSharedParams: false }
};
const FEATURE_KEYS = Object.keys(FEATURE_META);
const COMMON_OPTIONS_KEY = 'kky.hub.commonOptions';
const GROUP_FILTER_KEY = 'kky.hub.multiGroupFilter';
const GROUPS = [
  { id: 'all', label: '전체' },
  { id: 'bqc', label: '납품 시 BQC 검토' },
  { id: 'periodic', label: '주기적 검토' },
  { id: 'utility', label: '유틸리티' }
];

export function renderMulti(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);
  const top = document.querySelector('#topbar-root .topbar') || document.querySelector('.topbar');
  if (top) top.classList.add('hub-topbar');

  const state = {
    rvtList: [],
    rvtChecked: new Set(),
    busy: false,
    common: createConfigState({
      extraParams: '',
      targetFilter: '',
      excludeEndDummy: false
    }),
    features: {
      connector: createFeatureState({ tol: 1.0, unit: 'inch', param: 'Comments' }),
      guid: createFeatureState({ includeFamily: false, includeAnnotation: false }),
      points: createFeatureState({ unit: 'ft' })
    },
    results: {},
    sharedParamStatus: null,
    ui: {
      modalOpen: false,
      activeFeatureKey: '',
      activeFeatureTitle: '',
      panels: {},
      controls: {},
      lastProgressPct: 0,
      runCompleted: false,
      commonSummaryEl: null,
      sharedParamBanner: null,
      runSummaryTitle: null,
      runSummaryDetail: null,
      runSharedParamHint: null,
      selectedTableBody: null,
      selectedRows: new Map(),
      groupFilter: 'all',
      isRvtListExpanded: false
    }
  };

  FEATURE_KEYS.forEach((k) => {
    state.results[k] = { count: 0, stale: true };
  });

  const page = div('feature-shell multi-page HubShell');
  const hasLocalCommonOptions = loadCommonOptionsFromStorage();
  const header = div('feature-header multi-header');
  header.innerHTML = `
    <div class="feature-heading">
      <span class="feature-kicker">Multi RVT Hub</span>
      <h2 class="feature-title">다중 RVT 검토 허브</h2>
      <p class="feature-sub">파일별로 열고 선택된 기능을 순차 실행합니다.</p>
    </div>`;
  page.append(header);

  const layout = div('multi-layout HubBody');
  const leftCol = div('multi-left HubLeft');
  const rightCol = div('multi-right HubRight');

  const group1 = buildGroupSection('납품 시 BQC 검토', '커넥터 진단 (BQC용)', 'bqc');
  const group2 = buildGroupSection('주기적 검토', 'PMS / GUID / 파라미터 연동', 'periodic');
  const group3 = buildGroupSection('유틸리티', '공유 파라미터 연동 / Point 추출', 'utility');

  const group1Options = buildGroup1Options();
  group1.section.append(group1Options);
  group1.section.append(buildToggleRow('connector', buildConnectorConfig()));
  group2.section.append(buildPmsWorkflowRow());
  group2.section.append(buildToggleRow('guid', buildGuidConfig()));
  group3.section.append(buildToggleRow('points', buildPointsConfig()));

  const rightFilter = buildGroupFilter();
  const leftTop = div('left-sticky HubLeftTop');
  leftTop.append(buildRunBar());
  const leftSelected = div('HubLeftSelected');
  leftSelected.append(buildSelectedFeaturesSection());
  leftCol.append(leftTop, leftSelected, buildRvtSection());
  rightCol.append(rightFilter, group1.wrap, group2.wrap, group3.wrap);
  layout.append(leftCol, rightCol);
  page.append(layout);
  page.append(buildSettingsModal());
  page.append(buildRvtExpandedModal());
  target.append(page);

  renderGroupVisibility();

  onHost('hub:rvt-picked', (payload) => {
    const paths = Array.isArray(payload?.paths) ? payload.paths : [];
    if (!paths.length) return;
    let changed = false;
    paths.forEach((p) => {
      if (!state.rvtList.includes(p)) {
        state.rvtList.push(p);
        state.rvtChecked.add(p);
        changed = true;
      }
    });
    if (changed) {
      markAllStale();
      renderRvtList();
    }
  });

  onHost('hub:multi-progress', (payload) => {
    const basePct = Number(payload?.percent);
    const altPct = Number(payload?.phaseProgress);
    const hasBase = Number.isFinite(basePct);
    const hasAlt = Number.isFinite(altPct);
    const pctValue = hasBase ? basePct : (hasAlt ? altPct * 100 : state.ui.lastProgressPct);
    const pct = Math.max(0, Math.min(100, pctValue));
    state.ui.lastProgressPct = pct;
    const phase = String(payload?.phase || payload?.Phase || '').toLowerCase();
    ProgressDialog.show(payload?.title || '다중 RVT 검토', payload?.message || '');
    ProgressDialog.update(pct, payload?.message || '', payload?.detail || '');
    updateRunProgress(pct, payload?.message || '', payload?.detail || '');
    if (phase === 'done' || pct >= 100) {
      ProgressDialog.hide();
    }
  });

  onHost('hub:multi-done', (payload) => {
    setBusyState(false);
    ProgressDialog.update(100, '완료', '검토가 완료되었습니다.');
    ProgressDialog.hide();
    updateRunProgress(100, '완료', '검토가 완료되었습니다.');
    updateResultSummary(payload?.summary || {});
    state.ui.runCompleted = true;
    updateRunActionLabel();
  });

  onHost('hub:multi-error', (payload) => {
    setBusyState(false);
    ProgressDialog.hide();
    updateRunProgress(0, '오류 발생', payload?.message || '');
    toast(payload?.message || '배치 검토 중 오류가 발생했습니다.', 'err');
    state.ui.runCompleted = false;
    updateRunActionLabel();
  });

  onHost('hub:multi-exported', (payload) => {
    setBusyState(false);
    ProgressDialog.hide();
    state.ui.lastProgressPct = 0;
    const path = payload?.path;
    if (path) {
      requestAnimationFrame(() => {
        showExcelSavedDialog('엑셀 저장 완료', path, (p) => post('excel:open', { path: p }));
      });
    } else {
      toast(payload?.message || '엑셀 저장에 실패했습니다.', 'err');
    }
  });

  onHost('sharedparam:status', (payload) => {
    state.sharedParamStatus = payload || {};
    updateSharedParamBanner();
    updateRunSummary();
  });

  window.addEventListener('keydown', (ev) => {
    if (ev.key === 'Escape' && state.ui.isRvtListExpanded) {
      ev.preventDefault();
      closeExpandedRvtModal();
    }
  });

  function buildGroupSection(title, desc, groupId) {
    const wrap = div('multi-section');
    if (groupId) wrap.dataset.group = groupId;
    const head = div('multi-section-title');
    head.innerHTML = `<h3>${title}</h3><span class="feature-note">${desc}</span>`;
    wrap.append(head);
    return { wrap, section: wrap };
  }

  function buildGroupFilter() {
    const wrap = div('group-filter');
    const stored = getGroupFilter();
    state.ui.groupFilter = stored;
    GROUPS.forEach((group) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'group-filter__btn';
      btn.textContent = group.label;
      btn.dataset.group = group.id;
      btn.classList.toggle('is-active', group.id === stored);
      btn.addEventListener('click', () => {
        state.ui.groupFilter = group.id;
        saveGroupFilter(group.id);
        wrap.querySelectorAll('.group-filter__btn').forEach((el) => {
          el.classList.toggle('is-active', el.dataset.group === group.id);
        });
        renderGroupVisibility();
      });
      wrap.append(btn);
    });
    return wrap;
  }

  onHost('commonoptions:loaded', (payload) => {
    if (hasLocalCommonOptions) return;
    applyCommonOptionsFromStorage(payload);
    persistCommonOptions(state.common.configCommitted, { skipHost: true });
  });

  function buildGroup1Options() {
    const panel = div('group-common-mini');
    const header = div('group-common-mini__header');
    const title = document.createElement('h4');
    title.textContent = '그룹 공통 옵션';
    const settingsBtn = document.createElement('button');
    settingsBtn.type = 'button';
    settingsBtn.className = 'btn btn--secondary';
    settingsBtn.textContent = '공통 옵션 설정';
    settingsBtn.addEventListener('click', () => openSettings('common', '그룹 공통 옵션'));
    header.append(title, settingsBtn);

    const summary = div('group-common-mini__summary');
    state.ui.commonSummaryEl = summary;
    summary.textContent = buildCommonSummary();
    panel.append(header, summary);

    const fields = div('multi-config is-open');
    const extra = makeField('추가 Parameter 값 추출', 'extra', 'PM1, PM2', 'textarea');
    const filter = makeField('검토 대상 필터', 'filter', 'ex) PM1=Value;PM2=Value2', 'text');
    const exclude = makeCheckboxField('End_ + Dummy 패밀리 제외');

    const draft = state.common.configDraft;
    extra.input.value = draft.extraParams;
    filter.input.value = draft.targetFilter;
    exclude.input.checked = draft.excludeEndDummy;

    extra.input.addEventListener('change', () => {
      state.common.configDraft.extraParams = extra.input.value;
      markCommonDirty();
      updateCommonSummary(summary);
    });
    filter.input.addEventListener('change', () => {
      state.common.configDraft.targetFilter = filter.input.value;
      markCommonDirty();
      updateCommonSummary(summary);
    });
    exclude.input.addEventListener('change', () => {
      state.common.configDraft.excludeEndDummy = exclude.input.checked;
      markCommonDirty();
      updateCommonSummary(summary);
    });

    fields.append(extra.field, filter.field, exclude.field);
    fields.append(buildFilterExamples());
    fields.classList.add('settings-panel', 'is-open');
    state.ui.panels.common = fields;
    state.ui.controls.common = { extra, filter, exclude };

    return panel;
  }

  function buildToggleRow(key, config) {
    const meta = FEATURE_META[key] || {};
    const row = div('feature-row');
    row.dataset.key = key;
    const header = div('feature-row__header');
    const toggle = document.createElement('input');
    toggle.type = 'checkbox';
    toggle.className = 'feature-toggle';
    toggle.addEventListener('change', () => {
      const feature = state.features[key];
      feature.enabled = toggle.checked;
      if (!toggle.checked) {
        feature.applied = false;
        feature.dirty = false;
        resetDraftFromCommitted(key);
      } else {
        feature.applied = false;
        feature.dirty = false;
        openSettings(key, meta.label);
      }
      row.classList.toggle('is-active', toggle.checked);
      markStale(key);
      updateRunSummary();
    });

    const metaWrap = div('feature-row__left');
    const metaTitle = document.createElement('strong');
    metaTitle.textContent = meta.label || key;
    const metaDesc = document.createElement('span');
    metaDesc.textContent = meta.desc || '';
    metaWrap.append(toggle, metaTitle, metaDesc);

    header.append(metaWrap);
    row.append(header);
    config.title = meta.label || key;
    config.key = key;
    config.panel.classList.add('settings-panel', 'is-open');
    state.ui.panels[key] = config.panel;
    state.ui.controls[key] = config.controls || {};
    return row;
  }

  function buildPmsWorkflowRow() {
    const row = div('feature-row feature-row--workflow');
    const header = div('feature-row__header');
    const left = div('feature-row__left');
    const icon = document.createElement('span');
    icon.className = 'feature-row__icon';
    icon.textContent = 'PMS';
    const title = document.createElement('strong');
    title.textContent = 'PMS 검토';
    const desc = document.createElement('span');
    desc.textContent = 'Segment ↔ PMS 매핑 및 사이즈 검토 (워크플로우)';
    left.append(icon, title, desc);

    const right = div('feature-row__right');
    const chip = document.createElement('span');
    chip.className = 'chip chip--info';
    chip.textContent = '별도 워크플로우';
    right.append(chip);

    const summary = div('feature-row__summary');
    summary.textContent = '추출 → PMS 등록 → 매핑 준비 → 비교 실행 → 결과 내보내기';
    header.append(left, right);
    row.append(header, summary);
    row.addEventListener('click', () => {
      location.hash = '#segmentpms';
    });
    row.classList.add('is-clickable');
    return row;
  }

  function buildConnectorConfig() {
    const panel = div('multi-config');
    const tol = makeField('허용범위', 'tol', '', 'number');
    tol.input.value = state.features.connector.configDraft.tol;
    tol.input.addEventListener('change', () => {
      state.features.connector.configDraft.tol = parseFloat(tol.input.value || '1') || 1;
      markFeatureDirty('connector');
    });

    const unit = makeSelectField('단위', [
      { value: 'inch', label: 'inch' },
      { value: 'mm', label: 'mm' }
    ]);
    unit.select.value = state.features.connector.configDraft.unit;
    unit.select.addEventListener('change', () => {
      state.features.connector.configDraft.unit = unit.select.value;
      markFeatureDirty('connector');
    });

    const param = makeField('파라미터', 'param', 'Comments', 'text');
    param.input.value = state.features.connector.configDraft.param;
    param.input.addEventListener('change', () => {
      state.features.connector.configDraft.param = param.input.value || 'Comments';
      markFeatureDirty('connector');
    });

    panel.append(tol.field, unit.field, param.field);
    return { panel, controls: { tol, unit, param } };
  }

  function buildGuidConfig() {
    const panel = div('multi-config');
    const includeFamily = makeCheckboxField('패밀리 포함');
    includeFamily.input.checked = state.features.guid.configDraft.includeFamily;
    includeFamily.input.addEventListener('change', () => {
      state.features.guid.configDraft.includeFamily = includeFamily.input.checked;
      markFeatureDirty('guid');
    });
    const includeAnno = makeCheckboxField('Annotation 패밀리 포함');
    includeAnno.input.checked = state.features.guid.configDraft.includeAnnotation;
    includeAnno.input.addEventListener('change', () => {
      state.features.guid.configDraft.includeAnnotation = includeAnno.input.checked;
      markFeatureDirty('guid');
    });
    panel.append(includeFamily.field, includeAnno.field);
    return { panel, controls: { includeFamily, includeAnno } };
  }

  function buildPointsConfig() {
    const panel = div('multi-config');
    const unit = makeSelectField('단위', [
      { value: 'ft', label: 'Decimal Feet' },
      { value: 'm', label: 'Meters (m)' },
      { value: 'mm', label: 'Millimeters (mm)' }
    ]);
    unit.select.value = state.features.points.configDraft.unit;
    unit.select.addEventListener('change', () => {
      state.features.points.configDraft.unit = unit.select.value;
      markFeatureDirty('points');
    });
    panel.append(unit.field);
    return { panel, controls: { unit } };
  }

  function buildRvtSection() {
    const section = div('multi-section rvt-panel HubLeftRvt');
    const head = div('rvt-panel-header');
    const title = document.createElement('div');
    title.className = 'rvt-panel-title';
    const badge = document.createElement('span');
    badge.className = 'chip chip--info';
    title.innerHTML = '<h3>RVT 리스트</h3>';
    title.append(badge);

    const controls = div('multi-rvt-controls');
    const btnAdd = cardBtn('RVT 추가', () => post('hub:pick-rvt', {}), 'btn--primary');
    const btnRemove = cardBtn('선택 제거', () => {
      state.rvtList = state.rvtList.filter((p) => !state.rvtChecked.has(p));
      state.rvtChecked.clear();
      markAllStale();
      renderRvtList();
    }, 'btn--secondary');
    const btnClear = cardBtn('목록 지우기', () => {
      state.rvtList = [];
      state.rvtChecked.clear();
      markAllStale();
      renderRvtList();
    }, 'btn--danger');
    controls.append(btnAdd, btnRemove, btnClear);

    head.append(title, controls);
    section.append(head);

    const body = div('rvt-panel-body');
    const tableWrap = div('rvt-table-wrap');
    const { table, tbody, master } = createRvtTable();
    const summary = div('multi-rvt-summary');
    const footer = div('rvt-list-footer');
    const footerRight = div('rvt-list-footer__right');
    const expandBtn = cardBtn('리스트 크게 보기', () => openExpandedRvtModal(), 'btn--secondary');
    const empty = div('rvt-empty');
    const emptyTitle = document.createElement('strong');
    emptyTitle.textContent = '등록된 RVT가 없습니다.';
    const emptySub = document.createElement('span');
    emptySub.textContent = 'RVT 추가로 파일을 등록하세요.';
    const emptyBtn = cardBtn('RVT 추가', () => post('hub:pick-rvt', {}), 'btn--primary');
    empty.append(emptyTitle, emptySub, emptyBtn);

    tableWrap.append(table);
    footerRight.append(expandBtn);
    footer.append(summary, footerRight);
    body.append(tableWrap, empty, footer);
    section.append(body);

    function syncMaster() {
      const allChecked = state.rvtList.length > 0 && state.rvtList.every((p) => state.rvtChecked.has(p));
      master.checked = allChecked;
    }

    master.addEventListener('change', () => {
      if (master.checked) {
        state.rvtList.forEach((p) => state.rvtChecked.add(p));
      } else {
        state.rvtChecked.clear();
      }
      renderRvtList();
    });

    function renderRvtList() {
      const rows = state.rvtList.map((path, idx) => ({
        index: idx + 1,
        path,
        name: getRvtName(path),
        checked: state.rvtChecked.has(path),
        onToggle: (checked) => {
          if (checked) state.rvtChecked.add(path);
          else state.rvtChecked.delete(path);
          syncMaster();
        }
      }));
      const count = state.rvtList.length;
      tbody.innerHTML = '';
      if (count > 0) {
        renderRvtRows(tbody, rows);
      }
      summary.textContent = `총 파일 수: ${count}`;
      badge.textContent = `${count}개`;
      empty.style.display = count ? 'none' : 'flex';
      tableWrap.style.display = count ? 'block' : 'none';
      expandBtn.disabled = count === 0;
      syncMaster();
      btnRemove.disabled = state.rvtChecked.size === 0;
      btnClear.disabled = state.rvtList.length === 0;
      updateRunSummary();
      if (buildRvtExpandedModal.render) buildRvtExpandedModal.render();
    }

    buildRvtSection.render = renderRvtList;
    renderRvtList();
    return section;
  }

  function buildRvtExpandedModal() {
    const overlay = div('rvt-expand-overlay');
    overlay.classList.add('is-hidden');
    const modal = div('rvt-expand-modal');
    const header = div('rvt-expand-header');
    const titleWrap = div('rvt-expand-title');
    const title = document.createElement('h3');
    title.textContent = 'RVT 리스트 (확대 보기)';
    const badge = document.createElement('span');
    badge.className = 'chip chip--info';
    titleWrap.append(title, badge);
    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.className = 'btn btn--ghost rvt-expand-close';
    closeBtn.textContent = '✕';
    header.append(titleWrap, closeBtn);

    const body = div('rvt-expand-body');
    const tableWrap = div('rvt-expand-table');
    const { table, tbody, master } = createRvtTable();
    table.classList.add('rvt-expand-table__grid');
    tableWrap.append(table);
    body.append(tableWrap);

    const footer = div('rvt-expand-footer');
    const footerBtn = document.createElement('button');
    footerBtn.type = 'button';
    footerBtn.className = 'btn btn--secondary';
    footerBtn.textContent = '닫기';
    footer.append(footerBtn);

    modal.append(header, body, footer);
    overlay.append(modal);

    overlay.addEventListener('click', (ev) => {
      ev.stopPropagation();
    });
    modal.addEventListener('click', (ev) => {
      ev.stopPropagation();
    });
    closeBtn.addEventListener('click', closeExpandedRvtModal);
    footerBtn.addEventListener('click', closeExpandedRvtModal);

    master.addEventListener('change', () => {
      if (master.checked) {
        state.rvtList.forEach((p) => state.rvtChecked.add(p));
      } else {
        state.rvtChecked.clear();
      }
      renderRvtList();
    });

    function renderExpandedList() {
      const rows = state.rvtList.map((path, idx) => ({
        index: idx + 1,
        path,
        name: getRvtName(path),
        checked: state.rvtChecked.has(path),
        onToggle: (checked) => {
          if (checked) state.rvtChecked.add(path);
          else state.rvtChecked.delete(path);
          renderRvtList();
        }
      }));
      const count = state.rvtList.length;
      tbody.innerHTML = '';
      if (count > 0) {
        renderRvtRows(tbody, rows);
      }
      badge.textContent = `${count}개`;
      master.checked = count > 0 && state.rvtList.every((p) => state.rvtChecked.has(p));
    }

    buildRvtExpandedModal.overlay = overlay;
    buildRvtExpandedModal.badge = badge;
    buildRvtExpandedModal.render = renderExpandedList;
    return overlay;
  }

  function openExpandedRvtModal() {
    if (!buildRvtExpandedModal.overlay) return;
    state.ui.isRvtListExpanded = true;
    buildRvtExpandedModal.overlay.classList.remove('is-hidden');
    buildRvtExpandedModal.render();
  }

  function closeExpandedRvtModal() {
    if (!buildRvtExpandedModal.overlay) return;
    state.ui.isRvtListExpanded = false;
    buildRvtExpandedModal.overlay.classList.add('is-hidden');
  }

  function buildSelectedFeaturesSection() {
    const section = div('multi-section selected-panel');
    const head = div('selected-panel__header');
    const title = document.createElement('h3');
    title.textContent = '선택된 기능 목록';
    const count = document.createElement('span');
    count.className = 'chip chip--info';
    head.append(title, count);

    const table = document.createElement('table');
    table.className = 'selected-table';
    table.innerHTML = `
      <colgroup>
        <col style="width:45%">
        <col style="width:120px">
        <col style="width:120px">
        <col style="width:140px">
      </colgroup>
      <thead>
        <tr>
          <th>기능</th>
          <th class="selected-status-col">상태</th>
          <th>설정</th>
          <th>엑셀</th>
        </tr>
      </thead>
      <tbody></tbody>`;
    const tbody = table.querySelector('tbody');
    section.append(head, table);

    state.ui.selectedTableBody = tbody;
    state.ui.selectedCount = count;
    renderSelectedFeatures();
    return section;
  }

  function buildRunBar() {
    const bar = div('run-bar');
    const summary = div('run-summary');
    const summaryTitle = document.createElement('strong');
    summaryTitle.className = 'run-summary__title';
    const summaryDetail = document.createElement('span');
    summaryDetail.className = 'run-summary__detail';
    const sharedHint = div('run-summary__hint');
    sharedHint.style.display = 'none';
    summary.append(summaryTitle, summaryDetail, sharedHint);
    const status = div('run-status');
    const progressText = document.createElement('span');
    const progressDetail = document.createElement('small');
    const progressBar = document.createElement('div');
    progressBar.className = 'run-progress';
    const progressFill = document.createElement('div');
    progressFill.className = 'run-progress-fill';
    progressBar.append(progressFill);
    status.append(progressText, progressDetail, progressBar);

    const startBtn = cardBtn('검토 시작', handleRunAction, 'btn--primary');
    startBtn.classList.add('multi-start-btn');
    bar.append(summary, status, startBtn);

    buildRunBar.startBtn = startBtn;
    buildRunBar.summary = summary;
    buildRunBar.summaryTitle = summaryTitle;
    buildRunBar.summaryDetail = summaryDetail;
    buildRunBar.runSharedParamHint = sharedHint;
    buildRunBar.progressText = progressText;
    buildRunBar.progressDetail = progressDetail;
    buildRunBar.progressFill = progressFill;
    updateRunSummary();
    updateRunProgress(0, '대기 중', '');
    updateRunActionLabel();
    return bar;
  }

  function buildSettingsModal() {
    const overlay = div('modal-overlay');
    const modal = div('modal');
    const header = div('modal__header');
    const title = document.createElement('div');
    title.className = 'modal__title';
    const badge = document.createElement('span');
    badge.className = 'chip chip--warn';
    badge.style.display = 'none';
    header.append(title, badge);

    const body = div('modal__body');
    const form = div('modal__form');
    const help = div('modal__help');
    const sharedBanner = buildSharedParamStatusBanner();
    sharedBanner.style.display = 'none';
    form.append(sharedBanner);
    body.append(form, help);

    const footer = div('modal__footer');
    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.className = 'btn btn--ghost';
    cancelBtn.textContent = '취소';
    const applyBtn = document.createElement('button');
    applyBtn.type = 'button';
    applyBtn.className = 'btn btn--primary';
    applyBtn.textContent = '적용';
    footer.append(cancelBtn, applyBtn);

    modal.append(header, body, footer);
    overlay.append(modal);

    cancelBtn.addEventListener('click', cancelSettings);
    applyBtn.addEventListener('click', applySettings);

    buildSettingsModal.overlay = overlay;
    buildSettingsModal.modal = modal;
    buildSettingsModal.title = title;
    buildSettingsModal.badge = badge;
    buildSettingsModal.form = form;
    buildSettingsModal.help = help;
    buildSettingsModal.sharedBanner = sharedBanner;
    return overlay;
  }

  function makeField(label, name, placeholder, type) {
    const field = div('field');
    const lab = document.createElement('label');
    lab.textContent = label;
    const input = type === 'textarea' ? document.createElement('textarea') : document.createElement('input');
    if (type !== 'textarea') input.type = type;
    input.placeholder = placeholder || '';
    input.name = name;
    field.append(lab, input);
    return { field, input };
  }

  function makeSelectField(label, options) {
    const field = div('field');
    const lab = document.createElement('label');
    lab.textContent = label;
    const select = document.createElement('select');
    options.forEach((opt) => {
      const option = document.createElement('option');
      option.value = opt.value;
      option.textContent = opt.label;
      select.append(option);
    });
    field.append(lab, select);
    return { field, select };
  }

  function makeCheckboxField(label) {
    const field = div('field');
    const wrapper = document.createElement('label');
    const input = document.createElement('input');
    input.type = 'checkbox';
    wrapper.append(input, document.createTextNode(` ${label}`));
    field.append(wrapper);
    return { field, input };
  }

  function cardBtn(label, onClick, variant = 'btn--secondary') {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = `btn ${variant}`;
    btn.textContent = label;
    if (onClick) btn.addEventListener('click', onClick);
    return btn;
  }

  function buildSharedParamStatusBanner() {
    const banner = div('sharedparam-banner');
    const head = div('sharedparam-banner__head');
    const title = document.createElement('strong');
    title.textContent = 'Shared Parameter 상태';
    const badge = document.createElement('span');
    badge.className = 'sharedparam-banner__badge';
    const refresh = document.createElement('button');
    refresh.type = 'button';
    refresh.className = 'btn btn--ghost sharedparam-banner__refresh';
    refresh.textContent = '상태 새로고침';
    refresh.addEventListener('click', () => requestSharedParamStatus('manual'));
    head.append(title, badge, refresh);

    const path = div('sharedparam-banner__path');
    const pathLabel = document.createElement('span');
    pathLabel.textContent = '경로';
    const pathValue = document.createElement('span');
    pathValue.className = 'sharedparam-banner__value';
    path.append(pathLabel, pathValue);

    const note = div('sharedparam-banner__note');

    banner.append(head, path, note);
    state.ui.sharedParamBanner = {
      root: banner,
      badge,
      pathValue,
      note
    };
    return banner;
  }

  function markStale(key) {
    state.results[key].stale = true;
    state.results[key].count = 0;
    syncFeatureRow(key);
    updateFeatureSummary(key);
    post('hub:multi-clear', { key });
  }

  function markAllStale() {
    FEATURE_KEYS.forEach(markStale);
  }

  function syncFeatureRow(key) {
    updateSelectedFeatureRow(key);
  }

  function updateResultSummary(summary) {
    Object.keys(summary || {}).forEach((key) => {
      if (!state.results[key]) return;
      state.results[key].count = summary[key].rows || 0;
      state.results[key].stale = false;
      syncFeatureRow(key);
    });
  }

  function handleRunAction() {
    if (state.ui.runCompleted) {
      resetRunResults();
      return;
    }
    onRun();
  }

  function onRun() {
    state.ui.runCompleted = false;
    updateRunActionLabel();
    const selected = FEATURE_KEYS.filter((k) => state.features[k].enabled);
    if (!selected.length) {
      toast('선택된 기능이 없습니다.', 'warn');
      return;
    }
    if (!canRunWithSharedParams()) {
      return;
    }
    if (!state.rvtList.length) {
      toast('RVT 파일을 추가하세요.', 'warn');
      return;
    }
    setBusyState(true);
    ProgressDialog.show('다중 RVT 검토', '준비 중...');
    ProgressDialog.update(0, '준비 중...', '');
    post('hub:multi-run', buildPayload());
  }

  function buildPayload() {
    return {
      rvtPaths: state.rvtList.slice(),
      commonOptions: state.common.configCommitted,
      features: {
        connector: buildCommittedFeature('connector'),
        guid: buildCommittedFeature('guid'),
        points: buildCommittedFeature('points')
      }
    };
  }

  function onExport(key) {
    setBusyState(true);
    chooseExcelMode((mode) => {
      post('hub:multi-export', { key, excelMode: mode || 'fast' });
    });
  }

  function setBusyState(on) {
    state.busy = on;
    setBusy(on);
    if (buildRunBar.startBtn) buildRunBar.startBtn.disabled = on;
    FEATURE_KEYS.forEach(syncFeatureRow);
    const inputs = page.querySelectorAll('input, select, textarea, button');
    inputs.forEach((el) => {
      if (el.classList.contains('multi-start-btn')) return;
      if (on) {
        el.disabled = true;
      } else {
        el.disabled = false;
      }
    });
    if (!on) renderRvtList();
    if (!on) updateSharedParamRunState();
  }

  function renderRvtList() {
    if (buildRvtSection.render) buildRvtSection.render();
  }

  function openSettings(key, title) {
    const config = key === 'common' ? state.common : state.features[key];
    if (!buildSettingsModal.form) return;
    state.ui.modalOpen = true;
    state.ui.activeFeatureKey = key;
    state.ui.activeFeatureTitle = title || '';
    buildSettingsModal.title.textContent = `${title || ''} 설정`;
    const readiness = key === 'common' ? { label: '설정', className: 'chip--ok' } : getFeatureReadiness(config);
    if (buildSettingsModal.badge) {
      buildSettingsModal.badge.textContent = readiness.label;
      buildSettingsModal.badge.className = `chip ${readiness.className}`;
      buildSettingsModal.badge.style.display = readiness.className === 'chip--warn' ? 'inline-flex' : 'none';
    }
    buildSettingsModal.form.innerHTML = '';
    buildSettingsModal.help.innerHTML = '';
    resetDraftFromCommitted(key);
    syncControlsFromDraft(key);
    renderSharedParamBanner(key);
    const panel = getFeaturePanel(key);
    if (panel) buildSettingsModal.form.append(panel);
    renderHelp(key, title);
    buildSettingsModal.overlay.classList.add('is-open');
  }

  function closeSettings() {
    if (!buildSettingsModal.overlay) return;
    state.ui.modalOpen = false;
    buildSettingsModal.overlay.classList.remove('is-open');
  }

  function applySettings() {
    const key = state.ui.activeFeatureKey;
    if (!key) return;
    if (key === 'common') {
      commitConfig(state.common);
      updateCommonSummary();
      persistCommonOptions(state.common.configCommitted);
      emitCommonOptionsChanged();
      markStale('connector');
    } else {
      commitConfig(state.features[key]);
      markStale(key);
    }
    closeSettings();
  }

  function cancelSettings() {
    const key = state.ui.activeFeatureKey;
    if (!key) return;
    resetDraftFromCommitted(key);
    syncControlsFromDraft(key);
    closeSettings();
  }

  function getFeaturePanel(key) {
    return state.ui.panels[key] || null;
  }

  function updateRunSummary() {
    if (!buildRunBar.summary) return;
    const enabledCount = FEATURE_KEYS.filter((k) => state.features[k].enabled).length;
    const rvtCount = state.rvtList.length;
    if (buildRunBar.summaryTitle) {
      buildRunBar.summaryTitle.textContent = `선택 기능: ${enabledCount}개`;
    }
    if (buildRunBar.summaryDetail) {
      buildRunBar.summaryDetail.textContent = `RVT: ${rvtCount}개`;
    }
    renderSelectedFeatures();
    updateSharedParamRunState();
  }

  function updateRunProgress(percent, message, detail) {
    if (!buildRunBar.progressText) return;
    buildRunBar.progressText.textContent = message || '대기 중';
    buildRunBar.progressDetail.textContent = detail || '';
    if (buildRunBar.progressFill) {
      const pct = Math.max(0, Math.min(100, Number(percent) || 0));
      buildRunBar.progressFill.style.width = `${pct}%`;
    }
  }

  function updateRunActionLabel() {
    if (!buildRunBar.startBtn) return;
    buildRunBar.startBtn.textContent = state.ui.runCompleted ? '검토 결과 초기화' : '검토 시작';
  }

  function resetRunResults() {
    state.ui.runCompleted = false;
    state.ui.lastProgressPct = 0;
    updateRunProgress(0, '대기 중', '');
    FEATURE_KEYS.forEach((key) => {
      if (state.results[key]) {
        state.results[key].count = 0;
        state.results[key].stale = true;
      }
    });
    syncFeatureRow('connector');
    syncFeatureRow('guid');
    syncFeatureRow('points');
    updateRunActionLabel();
    post('hub:multi-clear', {});
  }

  function getFeatureReadiness(feature) {
    if (!feature?.enabled) {
      return { label: 'OFF', className: 'chip--off' };
    }
    if (!feature.applied || feature.dirty) {
      return { label: '설정 필요', className: 'chip--warn' };
    }
    return { label: '검토 준비됨', className: 'chip--ok' };
  }

  function updateDrawerBadge(key) {
    if (!state.ui.modalOpen || state.ui.activeFeatureKey !== key || !buildSettingsModal.badge) return;
    const readiness = getFeatureReadiness(state.features[key]);
    buildSettingsModal.badge.textContent = readiness.label;
    buildSettingsModal.badge.className = `chip ${readiness.className}`;
    buildSettingsModal.badge.style.display = readiness.className === 'chip--warn' ? 'inline-flex' : 'none';
  }

  function updateFeatureSummary(key) {
    updateSelectedFeatureRow(key);
    updateDrawerBadge(key);
  }

  function buildCommonSummary() {
    const committed = state.common.configCommitted;
    const extraCount = committed.extraParams ? committed.extraParams.split(',').filter((v) => v.trim()).length : 0;
    const filterText = committed.targetFilter ? committed.targetFilter : '필터 없음';
    const excludeText = committed.excludeEndDummy ? 'Dummy 제외' : 'Dummy 포함';
    return `extra=${extraCount} / ${filterText} / ${excludeText}`;
  }

  function updateCommonSummary(el) {
    const target = el || state.ui.commonSummaryEl;
    if (target) {
      target.textContent = buildCommonSummary();
    }
    updateFeatureSummary('connector');
  }

  function renderHelp(key, title) {
    const help = buildSettingsModal.help;
    if (!help) return;
    const helpTitle = document.createElement('strong');
    helpTitle.textContent = title || '설정 안내';
    const list = document.createElement('ul');
    list.className = 'help-list';
    getHelpItems(key).forEach((text) => {
      const item = document.createElement('li');
      item.textContent = text;
      list.append(item);
    });
    help.append(helpTitle, list);
  }

  function getHelpItems(key) {
    if (key === 'common') {
      return [
        '추가 Parameter 값은 콤마로 구분해 입력합니다.',
        '검토 대상 필터는 “PM1=Value” 형식으로 입력합니다.',
        'Dummy/End_ 패밀리 제외 여부를 설정합니다.'
      ];
    }
    if (key === 'connector') {
      return [
        '허용범위는 연결 판단 기준 거리입니다.',
        '단위는 inch/mm 중 선택 가능합니다.',
        '파라미터는 Comments 등 대상 값을 지정합니다.'
      ];
    }
    if (key === 'guid') {
      return [
        '패밀리/Annotation 포함 여부를 선택합니다.',
        '공유 파라미터 GUID 일치 여부를 검토합니다.'
      ];
    }
    if (key === 'points') {
      return [
        '좌표 추출 단위를 선택합니다.',
        'Decimal Feet / Meter / Millimeter를 지원합니다.'
      ];
    }
    return [];
  }

  function createFeatureState(config) {
    return {
      enabled: false,
      applied: false,
      dirty: false,
      configCommitted: deepCopy(config),
      configDraft: deepCopy(config)
    };
  }

  function createConfigState(config) {
    return {
      applied: false,
      dirty: false,
      configCommitted: deepCopy(config),
      configDraft: deepCopy(config)
    };
  }

  function deepCopy(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  function renderSelectedFeatures() {
    if (!state.ui.selectedTableBody) return;
    const enabledKeys = FEATURE_KEYS.filter((key) => state.features[key].enabled);
    state.ui.selectedRows.clear();
    state.ui.selectedTableBody.innerHTML = '';

    if (enabledKeys.length === 0) {
      const row = document.createElement('tr');
      const cell = document.createElement('td');
      cell.colSpan = 4;
      cell.className = 'selected-empty';
      cell.textContent = '선택된 기능이 없습니다.';
      row.append(cell);
      state.ui.selectedTableBody.append(row);
    }

    enabledKeys.forEach((key) => {
      const row = document.createElement('tr');
      row.dataset.key = key;
      const nameCell = document.createElement('td');
      const nameWrap = div('selected-name');
      const nameMain = document.createElement('strong');
      nameMain.textContent = FEATURE_META[key]?.label || key;
      const nameSub = document.createElement('span');
      nameSub.textContent = FEATURE_META[key]?.desc || '';
      nameWrap.append(nameMain, nameSub);
      nameCell.append(nameWrap);

      const statusCell = document.createElement('td');
      statusCell.className = 'selected-status-col';
      const statusChip = document.createElement('span');
      statusChip.className = 'chip status-chip';
      statusCell.append(statusChip);

      const settingsCell = document.createElement('td');
      const settingsBtn = document.createElement('button');
      settingsBtn.type = 'button';
      settingsBtn.className = 'btn btn--secondary';
      settingsBtn.textContent = '설정';
      settingsBtn.addEventListener('click', () => openSettings(key, FEATURE_META[key]?.label));
      settingsCell.append(settingsBtn);

      const exportCell = document.createElement('td');
      const exportBtn = document.createElement('button');
      exportBtn.type = 'button';
      exportBtn.className = 'btn btn--secondary';
      exportBtn.textContent = '엑셀 내보내기';
      exportBtn.addEventListener('click', () => onExport(key));
      exportCell.append(exportBtn);

      row.append(nameCell, statusCell, settingsCell, exportCell);
      state.ui.selectedTableBody.append(row);

      state.ui.selectedRows.set(key, { row, statusChip, exportBtn });
      updateSelectedFeatureRow(key);
    });

    if (state.ui.selectedCount) {
      state.ui.selectedCount.textContent = `${enabledKeys.length}개`;
    }
  }

  function updateSelectedFeatureRow(key) {
    const entry = state.ui.selectedRows.get(key);
    if (!entry) return;
    const status = getSelectedFeatureStatus(key);
    entry.statusChip.textContent = status.label;
    entry.statusChip.className = `chip status-chip ${status.className}`;

    const res = state.results[key];
    const hasResult = !!res && !res.stale && res.count > 0;
    entry.exportBtn.disabled = state.busy || !hasResult;
    entry.exportBtn.classList.toggle('btn--primary', hasResult);
    entry.exportBtn.classList.toggle('btn--secondary', !hasResult);
    entry.exportBtn.title = entry.exportBtn.disabled ? '검토 후 내보내기 가능' : '';
  }

  function getSelectedFeatureStatus(key) {
    const feature = state.features[key];
    if (!feature) return { label: '검토 전', className: 'status-chip--idle' };
    if (requiresSharedParams(key) && state.sharedParamStatus?.status && state.sharedParamStatus.status !== 'ok') {
      return { label: 'Shared Param 확인 필요', className: 'status-chip--warn' };
    }
    if (state.busy) {
      return { label: '진행 중', className: 'status-chip--running' };
    }
    const res = state.results[key];
    if (res && !res.stale && res.count > 0) {
      return { label: '완료', className: 'status-chip--done' };
    }
    if (!feature.applied || feature.dirty) {
      return { label: '검토 전', className: 'status-chip--idle' };
    }
    return { label: '검토 준비됨', className: 'status-chip--ready' };
  }

  function requiresSharedParams(key) {
    return !!FEATURE_META[key]?.requiresSharedParams;
  }

  function getGroupFilter() {
    try {
      return localStorage.getItem(GROUP_FILTER_KEY) || 'all';
    } catch {
      return 'all';
    }
  }

  function saveGroupFilter(value) {
    try {
      localStorage.setItem(GROUP_FILTER_KEY, value);
    } catch {
    }
  }

  function renderGroupVisibility() {
    const filter = state.ui.groupFilter || 'all';
    const sections = rightCol.querySelectorAll('.multi-section');
    sections.forEach((section) => {
      const group = section.dataset.group || '';
      const show = filter === 'all' || group === filter;
      section.classList.toggle('is-hidden', !show);
    });
  }

  function requestSharedParamStatus(context) {
    post('sharedparam:status', { source: 'multi', context });
  }

  function updateSharedParamBanner() {
    const banner = buildSettingsModal.sharedBanner || state.ui.sharedParamBanner?.root;
    if (!banner || banner.style.display === 'none') return;
    const status = state.sharedParamStatus || {};
    const label = status.statusLabel || '조회 중';
    const badge = state.ui.sharedParamBanner?.badge;
    const pathValue = state.ui.sharedParamBanner?.pathValue;
    const note = state.ui.sharedParamBanner?.note;
    if (badge) {
      badge.textContent = label;
      badge.classList.remove('is-ok', 'is-warn', 'is-error');
      if (status.status === 'ok') badge.classList.add('is-ok');
      else if (status.status === 'unset' || status.status === 'missing') badge.classList.add('is-warn');
      else badge.classList.add('is-error');
    }
    if (pathValue) {
      const pathText = status.path || '미설정';
      pathValue.textContent = pathText;
      pathValue.title = pathText;
    }
    if (note) {
      const warning = status.warning || status.errorMessage || '';
      note.textContent = warning ? warning : '';
      note.style.display = warning ? 'block' : 'none';
    }
  }

  function renderSharedParamBanner(key) {
    const banner = buildSettingsModal.sharedBanner;
    if (!banner) return;
    if (!requiresSharedParams(key)) {
      banner.style.display = 'none';
      return;
    }
    banner.style.display = 'block';
    buildSettingsModal.form.append(banner);
    requestSharedParamStatus('settings');
    updateSharedParamBanner();
  }

  function updateSharedParamRunState() {
    const needsShared = FEATURE_KEYS.some((key) => state.features[key].enabled && requiresSharedParams(key));
    const ok = state.sharedParamStatus?.status === 'ok';
    if (buildRunBar.runSharedParamHint) {
      if (needsShared && !ok) {
        const warning = state.sharedParamStatus?.warning || 'Shared Parameter 미등록으로 실행이 제한됩니다.';
        buildRunBar.runSharedParamHint.textContent = warning;
        buildRunBar.runSharedParamHint.style.display = 'block';
      } else {
        buildRunBar.runSharedParamHint.style.display = 'none';
      }
    }
    if (buildRunBar.startBtn && !state.busy) {
      buildRunBar.startBtn.disabled = needsShared && !ok;
    }
  }

  function canRunWithSharedParams() {
    const needsShared = FEATURE_KEYS.some((key) => state.features[key].enabled && requiresSharedParams(key));
    if (!needsShared) return true;
    const status = state.sharedParamStatus || {};
    if (status.status === 'ok') return true;
    if (!status.status) {
      requestSharedParamStatus('run');
      toast('Shared Parameter 상태를 확인 중입니다.', 'warn');
      return false;
    }
    requestSharedParamStatus('run');
    const msg = status.warning || status.errorMessage || 'Shared Parameter 상태를 확인하세요.';
    toast(msg, 'warn');
    return false;
  }

  function normalizeCommonOptions(raw) {
    return {
      extraParams: typeof raw?.extraParamsText === 'string' ? raw.extraParamsText : '',
      targetFilter: typeof raw?.targetFilterText === 'string' ? raw.targetFilterText : '',
      excludeEndDummy: !!raw?.excludeEndDummy
    };
  }

  function loadCommonOptionsFromStorage() {
    let stored = null;
    try {
      const raw = localStorage.getItem(COMMON_OPTIONS_KEY);
      if (raw) stored = JSON.parse(raw);
    } catch {
      stored = null;
    }
    if (stored) {
      applyCommonOptionsFromStorage(stored);
      return true;
    }
    post('commonoptions:get', { source: 'multi' });
    return false;
  }

  function applyCommonOptionsFromStorage(stored) {
    const normalized = normalizeCommonOptions(stored);
    state.common.configCommitted = deepCopy(normalized);
    state.common.configDraft = deepCopy(normalized);
    state.common.applied = true;
    state.common.dirty = false;
    syncControlsFromDraft('common');
    updateCommonSummary();
  }

  function persistCommonOptions(committed, options = {}) {
    const payload = {
      extraParamsText: committed.extraParams || '',
      targetFilterText: committed.targetFilter || '',
      excludeEndDummy: !!committed.excludeEndDummy
    };
    try {
      localStorage.setItem(COMMON_OPTIONS_KEY, JSON.stringify(payload));
    } catch {
    }
    if (!options.skipHost) {
      post('commonoptions:save', payload);
    }
  }

  function emitCommonOptionsChanged() {
    const detail = { ...state.common.configCommitted };
    window.dispatchEvent(new CustomEvent('commonOptions:changed', { detail }));
  }

  function buildCommittedFeature(key) {
    const feature = state.features[key];
    return {
      enabled: feature.enabled,
      ...feature.configCommitted
    };
  }

  function commitConfig(target) {
    target.configCommitted = deepCopy(target.configDraft);
    target.applied = true;
    target.dirty = false;
    if (state.ui.activeFeatureKey !== 'common') {
      updateFeatureSummary(state.ui.activeFeatureKey);
    }
  }

  function resetDraftFromCommitted(key) {
    if (key === 'common') {
      state.common.configDraft = deepCopy(state.common.configCommitted);
      state.common.dirty = false;
      return;
    }
    const feature = state.features[key];
    if (!feature) return;
    feature.configDraft = deepCopy(feature.configCommitted);
    feature.dirty = false;
  }

  function markFeatureDirty(key) {
    const feature = state.features[key];
    if (!feature) return;
    feature.dirty = true;
    feature.applied = false;
    markStale(key);
  }

  function markCommonDirty() {
    state.common.dirty = true;
    state.common.applied = false;
    markStale('connector');
  }

  function syncControlsFromDraft(key) {
    const controls = state.ui.controls[key];
    if (!controls) return;
    if (key === 'connector') {
      const draft = state.features.connector.configDraft;
      controls.tol.input.value = draft.tol;
      controls.unit.select.value = draft.unit;
      controls.param.input.value = draft.param;
    } else if (key === 'guid') {
      const draft = state.features.guid.configDraft;
      controls.includeFamily.input.checked = draft.includeFamily;
      controls.includeAnno.input.checked = draft.includeAnnotation;
    } else if (key === 'points') {
      const draft = state.features.points.configDraft;
      controls.unit.select.value = draft.unit;
    } else if (key === 'common') {
      const draft = state.common.configDraft;
      controls.extra.input.value = draft.extraParams;
      controls.filter.input.value = draft.targetFilter;
      controls.exclude.input.checked = draft.excludeEndDummy;
    }
  }

  function buildFilterExamples() {
    const wrap = div('filter-examples');
    const title = document.createElement('strong');
    title.textContent = '필터 예시';
    const note = document.createElement('p');
    note.textContent = '좌측 Param 토큰은 공백 없는 이름을 권장합니다. 구분자는 콤마(,) 또는 세미콜론(;)을 사용할 수 있습니다.';
    note.className = 'filter-examples__note';
    const list = document.createElement('ul');
    list.className = 'filter-examples__list';

    const examples = [
      "and(PM1='A',PM2='B')",
      "or(SYSTEM='DCW',SYSTEM='DHW')",
      "not(Family='End_Dummy')",
      "and(PM1='A',not(PM2='X'))",
      "PM1='A';PM2='B'"
    ];

    examples.forEach((text) => {
      const item = document.createElement('li');
      const code = document.createElement('code');
      code.textContent = text;
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'btn btn--ghost';
      btn.textContent = '복사';
      btn.addEventListener('click', () => copyToClipboard(text));
      item.append(code, btn);
      list.append(item);
    });

    wrap.append(title, note, list);
    return wrap;
  }

  function copyToClipboard(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(() => toast('복사되었습니다.', 'ok')).catch(() => toast('복사에 실패했습니다.', 'err'));
      return;
    }
    const temp = document.createElement('textarea');
    temp.value = text;
    temp.style.position = 'fixed';
    temp.style.opacity = '0';
    document.body.append(temp);
    temp.focus();
    temp.select();
    try {
      document.execCommand('copy');
      toast('복사되었습니다.', 'ok');
    } catch (e) {
      toast('복사에 실패했습니다.', 'err');
    } finally {
      temp.remove();
    }
  }
}
