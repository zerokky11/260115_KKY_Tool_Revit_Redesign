// Resources/HubUI/js/views/home.js
import { clear, div } from '../core/dom.js';

export function renderHome(root) {
    const target = root || document.getElementById('view-root') || document.getElementById('app');
    clear(target);

    const view = div('home-choice');
    const hero = div('home-choice-hero');
    hero.innerHTML = `
        <p class="home-choice-kicker">KKY Tool Hub</p>
        <h2>검토 방식을 선택하세요</h2>
        <p>활성 문서 기반 검토 또는 다중 RVT 배치 검토를 시작할 수 있습니다.</p>`;

    const grid = div('home-choice-grid');
    grid.append(
        buildCard(
            '활성 문서 검토',
            '현재 열려있는 Revit 문서를 대상으로 빠르게 검토를 수행합니다.',
            'active-menu',
            [
                '중복 객체 검토: 현재 열린 문서에서 중복 요소/패밀리 점검',
                '패밀리 공유파라미터 추가 및 연동: 공유 파라미터 추가/연동 수행'
            ]
        ),
        buildCard(
            '다중 RVT 검토',
            '여러 RVT 파일을 등록하고 배치 검토 및 엑셀 추출을 실행합니다.',
            'multi',
            [
                '커넥터 진단(BQC용): 파라미터 값 연속성/연결 관계 점검',
                'GUID 검토: 공유 파라미터 GUID 불일치 확인',
                '패밀리 공유파라미터 연동 검토: 복합 패밀리 연동 상태 점검',
                'Point 추출: Project/Survey 포인트 좌표 추출'
            ]
        )
    );

    view.append(hero, grid);
    target.append(view);

    function buildCard(title, desc, hash, items) {
        const card = document.createElement('button');
        card.type = 'button';
        card.className = 'home-choice-card';
        const listHtml = Array.isArray(items) && items.length
          ? `<ul class="home-choice-list">${items.map((item) => `<li>${item}</li>`).join('')}</ul>`
          : '';
        card.innerHTML = `
            <div class="home-choice-card__body">
              <div>
                <h3>${title}</h3>
                <p>${desc}</p>
                ${listHtml}
              </div>
              <span class="home-choice-card__icon">→</span>
            </div>
            <span class="home-choice-cta btn btn--primary">바로가기</span>`;
        card.addEventListener('click', () => { location.hash = `#${hash}`; });
        return card;
    }
}
