export function getRvtName(path, fallback = '') {
  if (!path) return fallback || '';
  const parts = String(path).split(/[/\\]/);
  return parts[parts.length - 1] || fallback || '';
}

export function createRvtTable() {
  const table = document.createElement('table');
  table.className = 'segmentpms-table rvt-register-table';
  table.style.tableLayout = 'fixed';

  const colgroup = document.createElement('colgroup');
  colgroup.innerHTML = '<col style="width:40px"><col style="width:50px"><col style="width:180px"><col>';
  table.append(colgroup);

  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');
  const thCheck = document.createElement('th');
  thCheck.style.textAlign = 'center';
  const master = document.createElement('input');
  master.type = 'checkbox';
  thCheck.append(master);

  const thIndex = document.createElement('th');
  thIndex.textContent = '#';
  thIndex.style.textAlign = 'center';
  const thName = document.createElement('th');
  thName.textContent = '파일명';
  const thPath = document.createElement('th');
  thPath.textContent = '파일 경로';

  headRow.append(thCheck, thIndex, thName, thPath);
  thead.append(headRow);
  const tbody = document.createElement('tbody');
  table.append(thead, tbody);

  return { table, tbody, master };
}

export function renderRvtRows(tbody, rows, emptyMessage = '등록된 RVT가 없습니다.') {
  tbody.innerHTML = '';
  if (!rows.length) {
    const emptyRow = document.createElement('tr');
    emptyRow.className = 'empty-row';
    const emptyCell = document.createElement('td');
    emptyCell.className = 'empty-cell';
    emptyCell.colSpan = 4;
    emptyCell.textContent = emptyMessage;
    emptyRow.append(emptyCell);
    tbody.append(emptyRow);
    return;
  }

  rows.forEach((row, idx) => {
    const tr = document.createElement('tr');
    const ckCell = document.createElement('td');
    ckCell.style.textAlign = 'center';
    const ck = document.createElement('input');
    ck.type = 'checkbox';
    ck.checked = !!row.checked;
    if (typeof row.onToggle === 'function') {
      ck.onchange = () => row.onToggle(ck.checked, idx);
    }
    ckCell.append(ck);

    const idxCell = document.createElement('td');
    idxCell.style.textAlign = 'center';
    idxCell.textContent = String(row.index);

    const nameCell = document.createElement('td');
    nameCell.className = 'segmentpms-path-cell';
    nameCell.textContent = row.name || '—';
    nameCell.title = row.name || '';

    const pathCell = document.createElement('td');
    pathCell.className = 'segmentpms-path-cell';
    pathCell.textContent = row.path || '—';
    pathCell.title = row.title || row.path || '';

    tr.append(ckCell, idxCell, nameCell, pathCell);
    tbody.append(tr);
  });
}
