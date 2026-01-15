import { initTheme } from './core/theme.js';
import { onHost, post } from './core/bridge.js';
import { updateTopMost, setActiveDocument, setDocList, renderTopbar } from './core/topbar.js';
import { initLogConsole, toggleLogConsole, log } from './core/dom.js';
import { renderHome } from './views/home.js';
import { renderActiveMenu } from './views/activeMenu.js';
import { renderDup } from './views/dup.js';
import { renderConn } from './views/conn.js';
import { renderExport } from './views/export.js';
import { renderParamProp } from './views/paramprop.js';
import { renderSegmentPms } from './views/segmentpms.js';
import { renderGuid } from './views/guid.js';
import { renderFamilyLink } from './views/familylink.js';
import { renderMulti } from './views/multi.js';

initTheme();

// 직전 TopMost 값 기억(중복 수신 무시)
let _lastTop = null;
let _viewRoot = null;
let _topbarRoot = null;
let _lastHash = null;
let _historyStack = [];
let _suppressHistory = false;

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();

function boot() {
    // 부팅 스켈레톤 제거
    const bootEl = document.getElementById('boot'); if (bootEl) bootEl.remove();
    const app = document.getElementById('app'); if (app) app.hidden = false;
    _viewRoot = document.getElementById('view-root') || app;
    _topbarRoot = document.getElementById('topbar-root') || app;
    renderTopbar(_topbarRoot, false, null);

    initLogConsole();

    window.addEventListener('keydown', (ev) => {
        const tag = ev.target && ev.target.tagName;
        if (tag === 'INPUT' || tag === 'TEXTAREA') return;

        if (ev.key === 'F12') {
            ev.preventDefault();
            toggleLogConsole();
            log('Log console toggled via F12');
        }
    });

    // 초기 TopMost 상태 동기화(이제 여기서만 전송)
    try { post('ui:query-topmost'); } catch { }

    route();
    window.addEventListener('hashchange', route);

    // 제네릭 onHost: { ev, payload }
    onHost((msg) => {
        try {
            if (!msg || !msg.ev) return;

            // (디버그용) 수신 로그
            try { console.log('[host] ←', msg.ev, msg.payload); } catch { }

            switch (msg.ev) {
                case 'host:topmost': {
                    const on = (msg && typeof msg.payload === 'object') ? !!msg.payload.on : !!msg.payload;
                    // ★ 디듀프: 이전 값과 같으면 무시
                    if (_lastTop === on) return;
                    _lastTop = on;
                    updateTopMost(on);
                    break;
                }
                case 'host:doc-changed': {
                    setActiveDocument(msg.payload || {});
                    break;
                }
                case 'host:doc-list': {
                    setDocList(msg.payload);
                    break;
                }
                default:
                    break;
            }
        } catch (e) {
            console.error('[main] onHost dispatch error:', e);
        }
    });
}

function route() {
    const hash = (location.hash || '').replace('#', '');
    if (!_suppressHistory && _lastHash !== null && _lastHash !== hash) {
        _historyStack.push(_lastHash);
    }
    if (_suppressHistory) _suppressHistory = false;
    if (hash === '') _historyStack = [];
    _lastHash = hash;

    const onBack = () => {
        _historyStack = [];
        location.hash = '';
    };
    const onNavBack = () => {
        if (_historyStack.length === 0) return;
        const prev = _historyStack.pop();
        _suppressHistory = true;
        location.hash = prev ? `#${prev}` : '';
    };
    const withBack = hash !== '';
    renderTopbar(_topbarRoot, withBack, hash === '' ? null : onBack, _historyStack.length > 0, onNavBack);
    if (_viewRoot) _viewRoot.innerHTML = '';
    const targetRoot = _viewRoot || document.getElementById('app');
    switch (hash) {
        case 'dup': return renderDup(targetRoot);
        case 'conn': return renderConn(targetRoot);
        case 'export': return renderExport(targetRoot);
        case 'paramprop': return renderParamProp(targetRoot);
        case 'segmentpms': return renderSegmentPms(targetRoot);
        case 'guid': return renderGuid(targetRoot);
        case 'familylink': return renderFamilyLink(targetRoot);
        case 'multi': return renderMulti(targetRoot);
        case 'active-menu': return renderActiveMenu(targetRoot);
        default: return renderHome(targetRoot);
    }
}
