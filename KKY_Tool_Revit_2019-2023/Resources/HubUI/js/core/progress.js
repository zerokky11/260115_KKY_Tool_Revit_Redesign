// Resources/HubUI/js/core/progress.js
// 공용 중앙 진행 다이얼로그 (SegmentPMS / ParamPropagate 등에서 재사용)

const MIN_UPDATE_MS = 140;

let root = null;
let titleEl = null;
let detailEl = null;
let metaEl = null;
let pctEl = null;
let barFillEl = null;
let actionsEl = null;
let cancelBtn = null;
let skipBtn = null;
let statusEl = null;
let lastUpdate = 0;
let pendingTimer = null;
let pendingData = null;
let isVisible = false;
let actionHandlers = { onCancel: null, onSkip: null };

function ensure() {
    if (root && root.isConnected) return;
    root = document.createElement('div');
    root.className = 'segmentpms-progress is-hidden';

    const card = document.createElement('div');
    card.className = 'segmentpms-progress-card';

    titleEl = document.createElement('div'); titleEl.className = 'segmentpms-progress-title';
    detailEl = document.createElement('div'); detailEl.className = 'segmentpms-progress-detail';
    metaEl = document.createElement('div'); metaEl.className = 'segmentpms-progress-meta';

    const bar = document.createElement('div'); bar.className = 'segmentpms-progress-bar';
    barFillEl = document.createElement('div'); barFillEl.className = 'segmentpms-progress-fill';
    bar.append(barFillEl);

    pctEl = document.createElement('div'); pctEl.className = 'segmentpms-progress-pct';

    actionsEl = document.createElement('div');
    actionsEl.className = 'segmentpms-progress-actions';
    statusEl = document.createElement('div');
    statusEl.className = 'segmentpms-progress-status';
    cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.className = 'btn btn--danger';
    cancelBtn.textContent = '취소';
    skipBtn = document.createElement('button');
    skipBtn.type = 'button';
    skipBtn.className = 'btn btn--secondary';
    skipBtn.textContent = '다음 파일로';
    cancelBtn.addEventListener('click', () => {
        if (actionHandlers.onCancel) actionHandlers.onCancel();
    });
    skipBtn.addEventListener('click', () => {
        if (actionHandlers.onSkip) actionHandlers.onSkip();
    });
    actionsEl.append(statusEl, skipBtn, cancelBtn);

    card.append(titleEl, detailEl, metaEl, bar, pctEl, actionsEl);
    root.append(card);
    document.body.append(root);
}

function applyUpdate({ percent, subtitle, detail }) {
    ensure();
    const pct = Math.max(0, Math.min(100, Number(percent) || 0));
    if (barFillEl) barFillEl.style.width = `${pct}%`;
    if (pctEl) {
        const label = Number.isInteger(pct) ? `${pct.toFixed(0)}%` : `${pct.toFixed(1)}%`;
        pctEl.textContent = label;
    }
    if (detailEl && subtitle != null) detailEl.textContent = subtitle;
    if (metaEl && detail != null) metaEl.textContent = detail;
}

function throttledUpdate(data) {
    const now = performance.now ? performance.now() : Date.now();
    const elapsed = now - lastUpdate;
    if (elapsed >= MIN_UPDATE_MS) {
        lastUpdate = now;
        applyUpdate(data);
        pendingData = null;
        if (pendingTimer) { clearTimeout(pendingTimer); pendingTimer = null; }
        return;
    }
    pendingData = data;
    if (!pendingTimer) {
        pendingTimer = setTimeout(() => {
            pendingTimer = null;
            if (pendingData) throttledUpdate(pendingData);
        }, MIN_UPDATE_MS - elapsed);
    }
}

export const ProgressDialog = {
    show(title, subtitle) {
        ensure();
        root.classList.remove('is-hidden');
        if (!isVisible) {
            isVisible = true;
            document.body.classList.add('is-busy');
        }
        if (titleEl) titleEl.textContent = title || '작업 진행 중';
        if (detailEl) detailEl.textContent = subtitle || '';
        if (metaEl) metaEl.textContent = '';
        if (statusEl) statusEl.textContent = '';
    },
    update(percent, subtitle, detail) {
        throttledUpdate({ percent, subtitle: subtitle ?? '', detail: detail ?? '' });
    },
    hide() {
        if (root) root.classList.add('is-hidden');
        if (isVisible) {
            isVisible = false;
            document.body.classList.remove('is-busy');
        }
        pendingData = null;
        if (pendingTimer) { clearTimeout(pendingTimer); pendingTimer = null; }
        if (statusEl) statusEl.textContent = '';
        if (cancelBtn) cancelBtn.disabled = false;
        if (skipBtn) skipBtn.disabled = false;
        actionHandlers = { onCancel: null, onSkip: null };
    }
};

ProgressDialog.setActions = function setActions({ onCancel, onSkip, cancelLabel, skipLabel } = {}) {
    ensure();
    actionHandlers = { onCancel: onCancel || null, onSkip: onSkip || null };
    if (cancelBtn) {
        cancelBtn.textContent = cancelLabel || '취소';
        cancelBtn.style.display = actionHandlers.onCancel ? 'inline-flex' : 'none';
    }
    if (skipBtn) {
        skipBtn.textContent = skipLabel || '다음 파일로';
        skipBtn.style.display = actionHandlers.onSkip ? 'inline-flex' : 'none';
    }
};

ProgressDialog.setActionState = function setActionState({ statusText, cancelDisabled, skipDisabled } = {}) {
    if (statusEl && typeof statusText === 'string') statusEl.textContent = statusText;
    if (cancelBtn && typeof cancelDisabled === 'boolean') cancelBtn.disabled = cancelDisabled;
    if (skipBtn && typeof skipDisabled === 'boolean') skipBtn.disabled = skipDisabled;
};

export default ProgressDialog;
