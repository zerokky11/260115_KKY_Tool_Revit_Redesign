const THEME_KEY = 'kky_theme';
export function applyTheme(t) { document.documentElement.dataset.theme = t; localStorage.setItem(THEME_KEY, t); }
export function toggleTheme() { const cur = localStorage.getItem(THEME_KEY) || 'light'; applyTheme(cur === 'dark' ? 'light' : 'dark'); }
export function initTheme() { applyTheme(localStorage.getItem(THEME_KEY) || 'light'); }
