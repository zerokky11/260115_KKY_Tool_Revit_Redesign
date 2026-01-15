import { clear, div, toast, setBusy, showExcelSavedDialog, chooseExcelMode } from '../core/dom.js';
import { ProgressDialog } from '../core/progress.js';
import { post, onHost } from '../core/bridge.js';
import { createRvtTable, renderRvtRows, getRvtName } from './rvtTable.js';

const FEATURE_KEYS = ['connector', 'guid', 'points'];

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
    ui: {
      modalOpen: false,
      activeFeatureKey: '',
      activeFeatureTitle: '',
      panels: {},
      controls: {},
      lastProgressPct: 0,
      runCompleted: false
    }
  };

  FEATURE_KEYS.forEach((k) => {
    state.results[k] = { count: 0, stale: true };
  });

  const page = div('feature-shell multi-page');
  const header = div('feature-header multi-header');
  header.innerHTML = `
    <div class="feature-heading">
      <span class="feature-kicker">Multi RVT Hub</span>
      <h2 class="feature-title">다중 RVT 검토 허브</h2>
      <p class="feature-sub">파일별로 열고 선택된 기능을 순차 실행합니다.</p>
    </div>`;
  page.append(header);

  const layout = div('multi-layout');
  const leftCol = div('multi-left');
  const rightCol = div('multi-right');

  const group1 = buildGroupSection('납품 시 BQC 검토', '커넥터 진단 (BQC용)');
  const group2 = buildGroupSection('주기적 검토', 'PMS / GUID / 파라미터 연동');
  const group3 = buildGroupSection('유틸리티', '공유 파라미터 연동 / Point 추출');

  const group1Options = buildGroup1Options();
  group1.section.append(group1Options);
  group1.section.append(buildToggleRow('connector', '커넥터 진단', 'Parameter 값 연속성 검토', buildConnectorConfig()));
  group2.section.append(buildPmsWorkflowRow());
  group2.section.append(buildToggleRow('guid', 'GUID 검토', '공유 파라미터 GUID 불일치 검토', buildGuidConfig()));
  group3.section.append(buildToggleRow('points', 'Point 추출', 'Project/Survey Point 좌표 추출', buildPointsConfig()));

  leftCol.append(group1.wrap, group2.wrap, group3.wrap);
  rightCol.append(buildRunBar(), buildRvtSection());
  layout.append(leftCol, rightCol);
  page.append(layout);
  page.append(buildSettingsModal());
  target.append(page);

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

  function buildGroupSection(title, desc) {
    const wrap = div('multi-section');
    const head = div('multi-section-title');
    head.innerHTML = `<h3>${title}</h3><span class="feature-note">${desc}</span>`;
    wrap.append(head);
    return { wrap, section: wrap };
  }

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

  function buildToggleRow(key, title, desc, config) {
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
        openSettings(key, title);
      }
      row.classList.toggle('is-active', toggle.checked);
      markStale(key);
      updateRunSummary();
    });

    const meta = div('feature-row__left');
    const metaTitle = document.createElement('strong');
    metaTitle.textContent = title;
    const metaDesc = document.createElement('span');
    metaDesc.textContent = desc;
    meta.append(toggle, metaTitle, metaDesc);

    const statusWrap = div('feature-row__right');
    const statusChip = document.createElement('span');
    statusChip.className = 'chip chip--off feature-chip';
    statusChip.addEventListener('click', () => {
      if (statusChip.classList.contains('chip--warn')) {
        openSettings(key, title);
      }
    });
    const resultChip = document.createElement('span');
    resultChip.className = 'chip chip--result';
    resultChip.style.display = 'none';
    statusWrap.append(statusChip, resultChip);

    const settingsBtn = document.createElement('button');
    settingsBtn.type = 'button';
    settingsBtn.className = 'btn btn--secondary settings-btn';
    settingsBtn.textContent = '설정';
    settingsBtn.addEventListener('click', () => openSettings(key, title));

    const exportBtn = document.createElement('button');
    exportBtn.type = 'button';
    exportBtn.className = 'btn btn--secondary export-btn';
    exportBtn.textContent = '엑셀 내보내기';
    exportBtn.disabled = true;
    exportBtn.addEventListener('click', () => onExport(key));
    statusWrap.append(statusChip, resultChip, settingsBtn, exportBtn);

    header.append(meta, statusWrap);
    const summary = div('feature-row__summary');
    summary.textContent = buildFeatureSummary(key);
    row.append(header, summary);
    config.exportBtn = exportBtn;
    config.statusChip = statusChip;
    config.resultChip = resultChip;
    config.summary = summary;
    config.title = title;
    config.key = key;
    config.panel.classList.add('settings-panel', 'is-open');
    state.ui.panels[key] = config.panel;
    state.ui.controls[key] = config.controls || {};
    syncFeatureRow(key);
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
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'btn btn--secondary';
    btn.textContent = 'PMS 워크플로우 열기';
    btn.addEventListener('click', () => {
      location.hash = '#segmentpms';
    });
    right.append(chip, btn);

    const summary = div('feature-row__summary');
    summary.textContent = '추출 → PMS 등록 → 매핑 준비 → 비교 실행 → 결과 내보내기';
    header.append(left, right);
    row.append(header, summary);
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
    const section = div('multi-section rvt-panel');
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
    }, 'btn--secondary');
    controls.append(btnAdd, btnRemove, btnClear);

    head.append(title, controls);
    section.append(head);

    const body = div('rvt-panel-body');
    const { table, tbody, master } = createRvtTable();
    const summary = div('multi-rvt-summary');
    const empty = div('rvt-empty');
    const emptyTitle = document.createElement('strong');
    emptyTitle.textContent = '등록된 RVT가 없습니다.';
    const emptySub = document.createElement('span');
    emptySub.textContent = 'RVT 추가로 파일을 등록하세요.';
    const emptyBtn = cardBtn('RVT 추가', () => post('hub:pick-rvt', {}), 'btn--primary');
    empty.append(emptyTitle, emptySub, emptyBtn);

    body.append(table, empty, summary);
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
      renderRvtRows(tbody, rows, '등록된 RVT가 없습니다.');
      const count = state.rvtList.length;
      summary.textContent = `총 파일 수: ${count}`;
      badge.textContent = `${count}개`;
      empty.style.display = count ? 'none' : 'flex';
      syncMaster();
      btnRemove.disabled = state.rvtChecked.size === 0;
      btnClear.disabled = state.rvtList.length === 0;
      updateRunSummary();
    }

    buildRvtSection.render = renderRvtList;
    renderRvtList();
    return section;
  }

  function buildRunBar() {
    const bar = div('run-bar');
    const summary = div('run-summary');
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
    const row = page.querySelector(`.feature-row[data-key="${key}"]`);
    if (!row) return;
    const exportBtn = row.querySelector('button.export-btn');
    const statusChip = row.querySelector('.feature-chip');
    const resultChip = row.querySelector('.chip--result');
    const res = state.results[key];
    exportBtn.disabled = state.busy || res.stale || res.count === 0;
    exportBtn.title = exportBtn.disabled ? '결과가 없습니다.' : '';

    const feature = state.features[key];
    const readiness = getFeatureReadiness(feature);
    if (statusChip) {
      statusChip.textContent = readiness.label;
      statusChip.className = `chip ${readiness.className} feature-chip`;
      statusChip.classList.toggle('is-clickable', readiness.className === 'chip--warn');
    }
    if (resultChip) {
      if (!res.stale && res.count > 0) {
        resultChip.textContent = `결과 ${res.count}`;
        resultChip.style.display = 'inline-flex';
      } else {
        resultChip.style.display = 'none';
      }
    }
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
    buildRunBar.summary.innerHTML = `<strong>선택 기능: ${enabledCount}개</strong><span>RVT: ${rvtCount}개</span>`;
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
    const row = page.querySelector(`.feature-row[data-key="${key}"]`);
    if (!row) return;
    const summary = row.querySelector('.feature-row__summary');
    if (!summary) return;
    summary.textContent = buildFeatureSummary(key);
    updateDrawerBadge(key);
  }

  function buildFeatureSummary(key) {
    const feature = state.features[key];
    if (key === 'connector') {
      const committed = feature.configCommitted;
      const commonCommitted = state.common.configCommitted;
      const extraCount = commonCommitted.extraParams ? commonCommitted.extraParams.split(',').filter((v) => v.trim()).length : 0;
      const filterText = commonCommitted.targetFilter ? commonCommitted.targetFilter : '필터 없음';
      const excludeText = commonCommitted.excludeEndDummy ? 'Dummy 제외' : 'Dummy 포함';
      return `tol=${committed.tol} ${committed.unit} / param=${committed.param} / extra=${extraCount} / ${filterText} / ${excludeText}`;
    }
    if (key === 'guid') {
      const committed = feature.configCommitted;
      const famText = committed.includeFamily ? 'Family=ON' : 'Family=OFF';
      const annoText = committed.includeAnnotation ? 'Annotation=ON' : 'Annotation=OFF';
      return `${famText} / ${annoText}`;
    }
    if (key === 'points') {
      const committed = feature.configCommitted;
      return `Unit=${committed.unit}`;
    }
    return '';
  }

  function buildCommonSummary() {
    const committed = state.common.configCommitted;
    const extraCount = committed.extraParams ? committed.extraParams.split(',').filter((v) => v.trim()).length : 0;
    const filterText = committed.targetFilter ? committed.targetFilter : '필터 없음';
    const excludeText = committed.excludeEndDummy ? 'Dummy 제외' : 'Dummy 포함';
    return `extra=${extraCount} / ${filterText} / ${excludeText}`;
  }

  function updateCommonSummary(el) {
    if (el) {
      el.textContent = buildCommonSummary();
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
