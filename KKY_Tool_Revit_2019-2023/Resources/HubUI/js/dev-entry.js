// Resources/HubUI/js/dev-entry.js
// 브라우저에서 index.html을 직접 열 때(#app이 숨겨져 있으면) 0.3초 후 대체 렌더.
// Revit(WebView2) 환경에서는 main.js가 이미 렌더하므로 아무것도 하지 않음.

import { renderTopbar } from './core/topbar.js';

function renderDevHome(root) {
    // 간단한 홈 카드 목업(스타일 확인용)
    const wrap = document.createElement('div');

    const hero = document.createElement('div');
    hero.className = 'hero';
    hero.innerHTML = `
    <h2>KKY Tools — 통합 도구 허브</h2>
    <p>fix2 규격으로 연결된 세 기능을 한 곳에서 실행합니다. (개발모드 미리보기)</p>`;
    wrap.appendChild(hero);

    const grid = document.createElement('div');
    grid.className = 'card-grid';
    grid.style.display = 'grid';
    grid.style.gridTemplateColumns = 'repeat(3, minmax(260px, 1fr))';
    grid.style.gap = '16px';

    function card(title, desc, tag) {
        const c = document.createElement('div');
        c.className = 'card';
        c.style.padding = '16px';
        c.style.border = '1px solid var(--border)';
        c.style.borderRadius = '14px';
        c.style.background = 'var(--panel)';
        c.style.boxShadow = 'var(--shadow)';
        c.innerHTML = `
      <div style="font-weight:900;margin-bottom:4px">${title}</div>
      <div style="color:var(--muted);min-height:40px">${desc}</div>
      <div style="display:flex;justify-content:space-between;align-items:center;margin-top:12px">
        <span class="kkyt-chip"># ${tag}</span>
        <button class="kkyt-btn">열기</button>
      </div>`;
        return c;
    }

    grid.append(
        card('중복검토', '그룹별 결과/삭제·되돌리기/선택·줌, 엑셀 내보내기까지 fix2 스키마로 제공합니다.', '검토'),
        card('커넥터 진단', '허용범위·단위·파라미터 입력 후 거리(inch) 일괄의 진단표를 생성/저장합니다.', '진단'),
        card('좌표/북각 추출', '폴더 스캔→미리보기→엑셀 내보내기까지 한 흐름으로 처리합니다.', '좌표')
    );

    wrap.appendChild(grid);
    root.appendChild(wrap);
}

function start() {
    const app = document.getElementById('app');
    const boot = document.getElementById('boot');
    if (!app) return;

    // 이미 렌더되었으면 무시
    if (!app.hidden) return;

    // 대체 렌더 시작
    app.hidden = false;
    if (boot) boot.remove();

    // 상단바 + 간단 홈
    renderTopbar(document.body, false, null);
    renderDevHome(app);
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => setTimeout(start, 300), { once: true });
} else {
    setTimeout(start, 300);
}
