export const $  = (s, root=document) => root.querySelector(s);
export const $$ = (s, root=document) => Array.from(root.querySelectorAll(s));
export const clear = el => { while (el.firstChild) el.removeChild(el.firstChild); };
export const div = cls => { const d=document.createElement('div'); d.className=cls||''; return d; };
export const tdText = v => { const t=document.createElement('td'); t.textContent = v==null ? '' : String(v); return t; };
export const debounce = (fn, delay=200) => { let t; return (...a)=>{ clearTimeout(t); t=setTimeout(()=>fn(...a), delay); }; };

let logConsoleBody = null;
let logConsoleRoot = null;

export function initLogConsole() {
  logConsoleRoot = document.querySelector('[data-log-console]');
  logConsoleBody = document.querySelector('[data-log-body]');

  const closeBtn = logConsoleRoot?.querySelector('.log-console-close');
  if (closeBtn) closeBtn.addEventListener('click', hideLogConsole);

  hideLogConsole();
}

export function showLogConsole() {
  if (!logConsoleRoot) return;
  logConsoleRoot.classList.remove('is-hidden');
}

export function hideLogConsole() {
  if (!logConsoleRoot) return;
  logConsoleRoot.classList.add('is-hidden');
}

export function toggleLogConsole() {
  if (!logConsoleRoot) return;
  if (logConsoleRoot.classList.contains('is-hidden')) showLogConsole();
  else hideLogConsole();
}

export function log(message, payload) {
  try {
    if (payload !== undefined) console.log(message, payload);
    else console.log(message);
  } catch { }

  if (!logConsoleBody) return;

  const now = new Date();
  const hh = String(now.getHours()).padStart(2, '0');
  const mm = String(now.getMinutes()).padStart(2, '0');
  const ss = String(now.getSeconds()).padStart(2, '0');

  const line = document.createElement('div');
  line.className = 'log-line';

  let text = `[${hh}:${mm}:${ss}] ${message}`;
  if (payload !== undefined) {
    try {
      const json = JSON.stringify(payload);
      text += ` ${json}`;
    } catch { }
  }

  line.textContent = text;
  logConsoleBody.appendChild(line);
  logConsoleBody.scrollTop = logConsoleBody.scrollHeight;
}

let busyEl=null;
export function setBusy(on, text='작업 중…'){
  if (on){
    if (busyEl) return;
    busyEl=document.createElement('div'); busyEl.className='busy';
    const sp=document.createElement('div'); sp.className='spinner'; sp.textContent=text;
    busyEl.append(sp); document.body.append(busyEl);
  } else { if (busyEl){ busyEl.remove(); busyEl=null; } }
}

export function toast(msg, kind='info', ms=2600){
  let wrap = $('.toast-wrap');
  if (!wrap){ wrap=document.createElement('div'); wrap.className='toast-wrap'; document.body.append(wrap); }
  const t = document.createElement('div');
  t.className = 'toast' + (kind==='ok'?' ok':kind==='err'?' err':'');
  t.textContent = msg; wrap.append(t);
  setTimeout(()=>{ t.remove(); if(!wrap.children.length) wrap.remove(); }, ms);
}

export function showExcelSavedDialog(message, filePath, onOpen){
  toast(message || '엑셀로 내보냈습니다.', 'ok');

  const existing = document.querySelector('.excel-dialog-backdrop');
  if (existing) existing.remove();

  const backdrop = document.createElement('div');
  backdrop.className = 'excel-dialog-backdrop';

  const dialog = document.createElement('div');
  dialog.className = 'excel-dialog';

  const title = document.createElement('div');
  title.className = 'excel-dialog-title';
  title.textContent = message || '엑셀 파일을 내보냈습니다.';

  const desc = document.createElement('div');
  desc.className = 'excel-dialog-desc';
  desc.textContent = '엑셀 파일을 바로 여시겠습니까?';

  const path = document.createElement('div');
  path.className = 'excel-dialog-path';
  path.textContent = filePath || '';

  const actions = document.createElement('div');
  actions.className = 'excel-dialog-actions';

  const btnOpen = document.createElement('button');
  btnOpen.type = 'button';
  btnOpen.className = 'btn btn-primary';
  btnOpen.textContent = '예, 엑셀 열기';

  const btnClose = document.createElement('button');
  btnClose.type = 'button';
  btnClose.className = 'btn';
  btnClose.textContent = '닫기';

  const close = () => backdrop.remove();

  btnOpen.addEventListener('click', () => {
    close();
    if (typeof onOpen === 'function') onOpen(filePath);
  });
  btnClose.addEventListener('click', close);
  backdrop.addEventListener('click', (e) => { if (e.target === backdrop) close(); });

  actions.append(btnOpen, btnClose);
  dialog.append(title, desc, path, actions);
  backdrop.append(dialog);
  document.body.append(backdrop);

  btnOpen.focus();
}

window.addEventListener('error', e => toast(`에러: ${e.message}`,'err',4200));
window.addEventListener('unhandledrejection', e => toast(`에러: ${e.reason}`,'err',4200));

// 엑셀 내보내기 모드 선택 (fast/normal)
export function chooseExcelMode(onSelect){
  const existing = document.querySelector('.excelmode-backdrop');
  if (existing) existing.remove();
  const backdrop = document.createElement('div');
  backdrop.className = 'excelmode-backdrop';

  const dialog = document.createElement('div');
  dialog.className = 'excelmode-dialog';

  const title = document.createElement('div');
  title.className = 'excelmode-title';
  title.textContent = '엑셀 내보내기 옵션을 선택하세요';

  const desc = document.createElement('div');
  desc.className = 'excelmode-desc';
  desc.textContent = '빠른 추출은 열 너비 자동 맞춤을 건너뛰고, 일반 추출은 저장 후 AutoFit을 수행합니다.';

  const actions = document.createElement('div');
  actions.className = 'excelmode-actions';

  const close = (mode) => {
    backdrop.remove();
    if (typeof onSelect === 'function') onSelect(mode || 'fast');
  };

  const btnFast = document.createElement('button');
  btnFast.type = 'button';
  btnFast.className = 'btn btn-primary';
  btnFast.textContent = '빠른 추출(기본)';
  btnFast.addEventListener('click', () => { close('fast'); });

  const btnNormal = document.createElement('button');
  btnNormal.type = 'button';
  btnNormal.className = 'btn';
  btnNormal.textContent = '일반 추출(열 너비 AutoFit)';
  btnNormal.addEventListener('click', () => { close('normal'); });

  actions.append(btnFast, btnNormal);
  dialog.append(title, desc, actions);
  backdrop.append(dialog);
  document.body.append(backdrop);

  backdrop.addEventListener('click', (e) => { if (e.target === backdrop) close('fast'); });
}
