export const FAV_KEY = 'kky_favs';
export const LAST_KEY = 'kky_last_run';
export const CARD_ORDER_KEY = 'kky_card_order';
const ALL_VIEWS = ['dup', 'conn', 'export', 'paramprop'];

export function getFavs() { try { const a = JSON.parse(localStorage.getItem(FAV_KEY) || '[]'); return a.filter(x => ALL_VIEWS.includes(x)); } catch { return []; } }
export function setFavs(list) {
    const uniq = list.filter((x, i, a) => a.indexOf(x) === i).filter(x => ALL_VIEWS.includes(x));
    localStorage.setItem(FAV_KEY, JSON.stringify(uniq));
}
export function toggleFav(view) {
    const arr = getFavs(); const i = arr.indexOf(view);
    if (i >= 0) arr.splice(i, 1); else arr.push(view);
    setFavs(arr); return arr;
}
export const saveLast = view => localStorage.setItem(LAST_KEY, JSON.stringify({ view, t: Date.now() }));
export const getLast = () => { try { return JSON.parse(localStorage.getItem(LAST_KEY) || 'null'); } catch { return null; } };

export function getCardOrder(defaultOrder) {
    try {
        const saved = JSON.parse(localStorage.getItem(CARD_ORDER_KEY) || '[]');
        return defaultOrder.filter(v => saved.includes(v)).concat(defaultOrder.filter(v => !saved.includes(v)));
    } catch { return defaultOrder; }
}
export const saveCardOrder = ids => localStorage.setItem(CARD_ORDER_KEY, JSON.stringify(ids));
