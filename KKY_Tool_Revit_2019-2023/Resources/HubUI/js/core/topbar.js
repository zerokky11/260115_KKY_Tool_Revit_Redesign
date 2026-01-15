// Resources/HubUI/js/core/topbar.js
import { div, toast } from './dom.js';
import { toggleTheme } from './theme.js';
import { setConn, ping, post } from './bridge.js';

export const APP_VERSION = 'v0.9.5';

let _docNameEl = null;
let _docSelectEl = null;
let _docList = [];
let _activeDoc = { name: '', path: '' };
let _topbarEl = null;
let _backBtn = null;
let _backHandler = null;
let _navBackBtn = null;
let _navWrap = null;
let _progressWrap = null;
let _progressFill = null;
let _progressText = null;
let _progressPct = null;

export function renderTopbar(root, withBack = false, onBack = null, canGoBack = false, onNavBack = null) {
    const host = document.getElementById('topbar-root') || root;
    if (!host) return;
    if (_topbarEl) {
        if (!_topbarEl.parentElement) host.append(_topbarEl);
    } else {
        _topbarEl = div('topbar');
        const left = div('topbar-left');
        const center = div('topbar-center');
        center.innerHTML = '<p class="topbar-tagline">Revit 워크플로우를 하나의 허브에서 관리하세요.</p>';

        const right = div('topbar-right');
        _topbarEl.append(left, center, right);
        host.append(_topbarEl);
        _navWrap = div('topbar-nav');
        left.append(_navWrap);
        buildBrand(left);
        renderTopbarChips();

        _progressWrap = div('topbar-progress hidden');
        const progRow = div('topbar-progress-row');
        _progressText = div('topbar-progress-text');
        _progressPct = div('topbar-progress-pct');
        progRow.append(_progressText, _progressPct);
        const progBar = div('topbar-progress-bar');
        _progressFill = div('topbar-progress-fill'); _progressFill.style.width = '0%';
        progBar.append(_progressFill);
        _progressWrap.append(progRow, progBar);
        _topbarEl.append(_progressWrap);
    }

    configureBackButton(withBack, onBack, canGoBack, onNavBack);
    setConn(true);
}

function configureBackButton(withBack, onBack, canGoBack, onNavBack) {
    const left = _topbarEl?.querySelector('.topbar-left');
    if (!left) return;
    if (!_navWrap) {
        _navWrap = div('topbar-nav');
        left.prepend(_navWrap);
    }
    if (!_backBtn) {
        _backBtn = document.createElement('button');
        _backBtn.className = 'btn btn-ghost';
        _backBtn.type = 'button';
        const icon = document.createElement('img');
        icon.className = 'back-btn-icon';
        icon.src = 'assets/icons/HubHome_24.png';
        icon.alt = '';
        const label = document.createElement('span');
        label.className = 'back-btn-label';
        label.textContent = '허브 홈으로';
        _backBtn.append(icon, label);
        _navWrap.append(_backBtn);
    } else {
        const label = _backBtn.querySelector('.back-btn-label');
        if (label) label.textContent = '허브 홈으로';
    }

    if (!_navBackBtn) {
        _navBackBtn = document.createElement('button');
        _navBackBtn.className = 'btn btn-ghost';
        _navBackBtn.type = 'button';
        const icon = document.createElement('img');
        icon.className = 'back-btn-icon';
        icon.src = 'assets/icons/HubHome_24.png';
        icon.alt = '';
        const label = document.createElement('span');
        label.className = 'back-btn-label';
        label.textContent = '뒤로가기';
        _navBackBtn.append(icon, label);
        _navWrap.append(_navBackBtn);
    }

    _backHandler = onBack;
    const smartGoHome = () => {
        try { window.dispatchEvent(new CustomEvent('kkyt:go-home')); } catch (_) { /* noop */ }
        const before = location.href;
        try { history.back(); } catch (_) { /* ignore */ }
        setTimeout(() => {
            if (location.href === before) {
                const url = new URL(location.href);
                const parts = url.pathname.split('/');
                parts[parts.length - 1] = 'index.html';
                url.pathname = parts.join('/');
                location.href = url.toString();
            }
        }, 80);
    };
    _backBtn.onclick = () => {
        if (typeof _backHandler === 'function') {
            try { _backHandler(); } catch (_) { smartGoHome(); }
        } else {
            smartGoHome();
        }
    };
    _backBtn.classList.toggle('hidden', !withBack);

    if (_navBackBtn) {
        _navBackBtn.disabled = !canGoBack;
        _navBackBtn.classList.toggle('is-disabled', !canGoBack);
        _navBackBtn.onclick = () => {
            if (!canGoBack) return;
            if (typeof onNavBack === 'function') onNavBack();
        };
    }
}

function buildBrand(host) {
    const wrap = div('topbar-brand');
    const logo = document.createElement('span');
    logo.className = 'topbar-logo';
    const logoImg = document.createElement('img');
    logoImg.src = 'assets/icons/KKY_Tool_48.png';
    logoImg.alt = 'KKY Tool';
    logo.append(logoImg);

    const text = document.createElement('div');
    text.className = 'topbar-brand-text';
    text.innerHTML = '<strong>KKY Tool Hub</strong><span>Revit 작업 보조 통합 도구</span>';

    const ver = document.createElement('span');
    ver.className = 'topbar-version';
    ver.textContent = APP_VERSION;

    wrap.append(logo, text, ver);
    host.append(wrap);
}

export function renderTopbarChips() {
    const actions = document.querySelector('.topbar-right');
    if (!actions) return;
    actions.innerHTML = '';

    const docCtrl = createDocControl();
    actions.append(docCtrl);

    const chipRow = document.createElement('div');
    chipRow.className = 'chip-row';

    const conn = createControlButton({
        id: 'connChip',
        label: '연결됨',
        icon: 'plug',
        classes: 'chip-toggle chip-connection',
        statusDot: true
    });
    conn.addEventListener('click', ping);
    chipRow.append(conn);

    const pin = createControlButton({
        id: 'pinChip',
        label: '항상 위',
        icon: 'pin',
        classes: 'chip-toggle pin-chip'
    });
    pin.setAttribute('aria-pressed', 'false');
    pin.classList.add('is-off');
    pin.onclick = () => { try { post('ui:toggle-topmost'); } catch (e) { console.error(e); } };
    chipRow.append(pin);

    const themeBtn = createControlButton({
        label: '테마',
        icon: 'theme',
        classes: 'chip-btn theme-chip'
    });
    themeBtn.setAttribute('aria-pressed', 'false');
    const applyThemeState = () => {
        const cur = document.documentElement.dataset.theme || 'dark';
        themeBtn.classList.toggle('is-dark', cur === 'dark');
        themeBtn.classList.toggle('is-active', cur === 'dark');
        themeBtn.setAttribute('aria-pressed', cur === 'dark' ? 'true' : 'false');
    };
    themeBtn.onclick = () => { toggleTheme(); applyThemeState(); };
    applyThemeState();
    chipRow.append(themeBtn);

    const help = createControlButton({
        label: '설정',
        icon: 'gear',
        classes: 'chip-btn settings-chip'
    });
    help.setAttribute('aria-expanded', 'false');
    help.addEventListener('click', () => toggleHelpPanel(help));
    chipRow.append(help);

    actions.append(chipRow);

    applyActiveDocumentState();
}

export function updateTopMost(on) {
    const pin = document.querySelector('#pinChip');
    if (!pin) return;
    const active = !!on;
    pin.classList.toggle('is-active', active);
    pin.classList.toggle('is-off', !active);
    pin.setAttribute('aria-pressed', active ? 'true' : 'false');
    const label = pin.querySelector('.chip-text');
    if (label) label.textContent = '항상 위';
}

export function setTopbarProgress(state) {
    if (!_progressWrap || !_progressFill || !_progressText || !_progressPct) return;
    if (!state) {
        _progressWrap.classList.add('hidden');
        _progressFill.style.width = '0%';
        _progressText.textContent = '';
        _progressPct.textContent = '';
        return;
    }
    const total = Math.max(0, Number(state.total) || 0);
    const idx = Math.max(0, Number(state.index) || 0);
    const pct = Math.max(0, Math.min(100, Number(state.percent) || 0));
    const file = state.file || '';
    const msg = state.message || '';
    _progressWrap.classList.remove('hidden');
    _progressFill.style.width = `${pct}%`;
    _progressText.textContent = `(${idx}/${total || '?'}) ${file} - ${msg}`;
    _progressPct.textContent = `${Math.round(pct)}%`;
}

function applyActiveDocumentState() {
    if (_docNameEl) {
        const hasDoc = !!_activeDoc.name;
        _docNameEl.textContent = hasDoc ? _activeDoc.name : '활성 문서 없음';
        _docNameEl.title = _activeDoc.path || _activeDoc.name || '';
        _docNameEl.classList.toggle('is-empty', !hasDoc);
    }
    updateConnectionChipLabel();
    rebuildDocSelect();
}

function updateConnectionChipLabel() {
    const label = document.querySelector('#connChip .chip-text');
    if (!label) return;
    label.textContent = _activeDoc.name ? `연결됨 · ${_activeDoc.name}` : '연결됨';
}

function normalizeDocs(payload) {
    let list = [];
    if (Array.isArray(payload?.docs)) list = payload.docs;
    if (Array.isArray(payload)) list = payload;
    return list
        .map(doc => ({ name: doc?.name || '', path: doc?.path || '' }))
        .filter(doc => !!doc.path);
}

function rebuildDocSelect() {
    if (!_docSelectEl) return;
    const activePath = (_activeDoc.path || '').toLowerCase();

    _docSelectEl.innerHTML = '';
    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = _docList.length ? '다른 문서로 전환' : '열린 Revit 문서가 없습니다';
    placeholder.disabled = !_docList.length;
    placeholder.selected = true;
    _docSelectEl.append(placeholder);

    let hasSelected = false;

    for (const doc of _docList) {
        const opt = document.createElement('option');
        opt.value = doc?.path || '';
        opt.textContent = doc?.name || doc?.path || '(이름 없는 문서)';
        opt.title = doc?.path || opt.textContent;
        if (opt.value && opt.value.toLowerCase() === activePath) {
            opt.selected = true;
            hasSelected = true;
        }
        _docSelectEl.append(opt);
    }

    if (hasSelected) placeholder.selected = false;
    _docSelectEl.disabled = !_docList.length;
}

function createDocControl() {
    const wrap = document.createElement('div');
    wrap.className = 'doc-chip';

    const glyph = document.createElement('span');
    glyph.className = 'chip-glyph';
    glyph.innerHTML = iconSvg('doc');

    const meta = document.createElement('div');
    meta.className = 'doc-meta';
    const label = document.createElement('span');
    label.className = 'doc-label';
    label.textContent = '연결 문서';
    const name = document.createElement('span');
    name.className = 'doc-name';
    name.textContent = '활성 문서 없음';
    meta.append(label, name);

    const select = document.createElement('select');
    select.className = 'doc-select';
    select.addEventListener('change', () => {
        const path = select.value;
        if (path && path !== _activeDoc.path) {
            try { toast?.('문서 전환은 Revit에서 프로젝트를 선택해 주세요.'); } catch (_) { console.warn('문서 전환은 Revit에서 프로젝트를 선택해 주세요.'); }
        }
        if (_activeDoc.path) {
            select.value = _activeDoc.path;
        } else {
            select.value = '';
        }
    });

    _docNameEl = name;
    _docSelectEl = select;
    rebuildDocSelect();

    wrap.append(glyph, meta, select);
    return wrap;
}

export function setActiveDocument(doc = {}) {
    _activeDoc = {
        name: doc?.name || '',
        path: doc?.path || ''
    };
    applyActiveDocumentState();
}

export function setDocList(payload) {
    const list = normalizeDocs(payload);
    _docList = list;
    rebuildDocSelect();
}

function toggleHelpPanel(trigger) {
    const existing = document.querySelector('.settings-backdrop');
    if (existing) {
        if (existing._escListener) {
            document.removeEventListener('keydown', existing._escListener);
        }
        existing.remove();
        if (trigger) trigger.setAttribute('aria-expanded', 'false');
        return;
    }

    const backdrop = document.createElement('div');
    backdrop.className = 'settings-backdrop';
    const panel = document.createElement('section');
    panel.className = 'settings-panel';

    const header = document.createElement('header');
    header.innerHTML = '<span>도움말 — KKY Tool Hub</span>';

    const body = document.createElement('div');
    body.className = 'body';
    body.innerHTML = `
        <div><strong>단축키</strong></div>
        <ul>
            <li><code>/</code> 검색 포커스</li>
            <li>카드 선택 후 <code>Enter</code> 실행, <code>F</code> 즐겨찾기</li>
            <li><code>Ctrl</code>+<code>Shift</code>+<code>L</code> 테마 전환</li>
            <li><code>Ctrl</code>+<code>Shift</code>+<code>T</code> 항상 위</li>
            <li><code>F1</code> 또는 <code>?</code> 도움말</li>
        </ul>`;

    const actions = document.createElement('div');
    actions.className = 'actions';
    const closeBtn = document.createElement('button');
    closeBtn.className = 'btn';
    closeBtn.type = 'button';
    closeBtn.textContent = '닫기';

    const closePanel = () => {
        backdrop.remove();
        if (trigger) trigger.setAttribute('aria-expanded', 'false');
        document.removeEventListener('keydown', escListener);
    };

    const escListener = (e) => { if (e.key === 'Escape') closePanel(); };

    document.addEventListener('keydown', escListener);
    backdrop._escListener = escListener;
    closeBtn.onclick = closePanel;
    actions.append(closeBtn);

    panel.append(header, body, actions);
    backdrop.append(panel);
    document.body.append(backdrop);
    trigger?.setAttribute('aria-expanded', 'true');

    backdrop.addEventListener('click', (e) => { if (e.target === backdrop) closePanel(); });
}

function createControlButton({ id, label, icon, classes = '', statusDot = false }) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = `control-chip ${classes}`.trim();
    if (id) btn.id = id;
    if (statusDot) {
        btn.classList.add('has-status-dot');
        const status = document.createElement('span');
        status.className = 'status-dot';
        status.setAttribute('aria-hidden', 'true');
        btn.append(status);
    }
    const glyph = document.createElement('span');
    glyph.className = 'chip-glyph';
    glyph.setAttribute('aria-hidden', 'true');
    glyph.innerHTML = iconSvg(icon);
    const text = document.createElement('span');
    text.className = 'chip-text';
    text.textContent = label;
    btn.append(glyph, text);
    return btn;
}

function iconSvg(name) {
    switch (name) {
        case 'plug':
            return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M8 3v5m8-5v5" stroke-linecap="round"/><path d="M6 8h12v5a6 6 0 1 1-12 0Z" stroke-linejoin="round"/><path d="M12 18v3" stroke-linecap="round"/></svg>';
        case 'pin':
            return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M8 4h8l-1 5 3 3-6 6-6-6 3-3z" fill="none"/><path d="M12 18v4" stroke-linecap="round"/></svg>';
        case 'theme':
            return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path class="half-dark" d="M12 3a9 9 0 0 1 0 18V3Z" fill="currentColor"/><path class="half-light" d="M12 21a9 9 0 0 1 0-18v18Z" fill="currentColor"/><circle cx="12" cy="12" r="8.5" fill="none"/></svg>';
        case 'gear':
            return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M20 13.5v-3l-2.1-.6a6.1 6.1 0 0 0-.6-1.4l1.2-1.8-2.1-2.1-1.8 1.2a6.1 6.1 0 0 0-1.4-.6L13.5 2h-3l-.6 2.1a6.1 6.1 0 0 0-1.4.6L6.7 3.5 4.6 5.6l1.2 1.8c-.26.44-.47.91-.6 1.4L3 10.5v3l2.1.6c.13.49.34.96.6 1.4l-1.2 1.8 2.1 2.1 1.8-1.2c.44.26.91.47 1.4.6l.6 2.1h3l.6-2.1c.49-.13.96-.34 1.4-.6l1.8 1.2 2.1-2.1-1.2-1.8c.26-.44.47-.91.6-1.4Z" fill="none"/><circle cx="12" cy="12" r="3.2" fill="none"/></svg>';
        case 'help':
            return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><circle cx="12" cy="12" r="9"/><path d="M10.89 9.05a1.11 1.11 0 0 1 2.22 0c0 1.11-1.67 1.11-1.67 2.78" stroke-linecap="round"/><circle cx="12" cy="15.5" r="0.5" fill="currentColor" stroke="none"/></svg>';
        case 'doc':
            return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M6 3h7l5 5v13H6z" fill="none"/><path d="M13 3v6h5" fill="none"/><path d="M9 13h6m-6 3h6" stroke-linecap="round"/></svg>';
        default:
            return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><circle cx="12" cy="12" r="8"/></svg>';
    }
}
