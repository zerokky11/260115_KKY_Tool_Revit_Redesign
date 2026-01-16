import { clear, div } from '../core/dom.js';

export function renderActiveMenu(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);

  const page = div('active-menu feature-shell');
  const header = div('feature-header');
  const heading = div('feature-heading');
  heading.innerHTML = `
    <span class="feature-kicker">Active Document</span>
    <h2 class="feature-title">활성 문서 검토</h2>
    <p class="feature-sub">현재 열려있는 Revit 문서에서 사용할 기능을 선택하세요.</p>`;
  header.append(heading);
  page.append(header);

  const grid = div('active-menu-grid');
  grid.append(
    buildCard(
      '중복검토',
      '중복 패밀리/요소를 그룹별로 확인하고 삭제/되돌리기를 관리합니다.',
      'dup'
    ),
    buildCard(
      '패밀리 공유파라미터 추가 및 연동(활성 문서)',
      '활성 문서에서 공유 파라미터 추가 및 연동 상태를 점검합니다.',
      'paramprop'
    )
  );

  page.append(grid);
  target.append(page);

  function buildCard(title, desc, hash) {
    const card = document.createElement('button');
    card.type = 'button';
    card.className = 'active-menu-card';
    card.innerHTML = `
      <div class="active-menu-card__body">
        <div>
          <h3>${title}</h3>
          <p>${desc}</p>
        </div>
        <span class="active-menu-card__icon">→</span>
      </div>
      <span class="active-menu-cta btn btn--primary">열기</span>`;
    card.addEventListener('click', () => { location.hash = `#${hash}`; });
    return card;
  }
}
