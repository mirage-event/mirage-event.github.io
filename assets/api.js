// Client for the Apps Script backend.
// GET for reads; POST with a text/plain body for writes (no CORS preflight — Apps Script can't answer one).
const API = {
  configured() { const u = window.MIRAGE_API || ''; return /^https?:\/\//.test(u) || u.startsWith('/'); },
  async get(params) {
    const r = await fetch(window.MIRAGE_API + '?' + new URLSearchParams(params), { redirect: 'follow' });
    return r.json();
  },
  async post(body) {
    const r = await fetch(window.MIRAGE_API, { method: 'POST', body: JSON.stringify(body), redirect: 'follow' });
    return r.json();
  },
};
const store = {
  get(k){ try { return localStorage.getItem(k); } catch(e){ return null; } },
  set(k, v){ try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch(e){} },
};
const AR_MONTHS = ['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'];
const EN_MONTHS = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'];
function fmtAr(iso){ const p = String(iso).split('-'); return p.length === 3 ? `${+p[2]} ${AR_MONTHS[+p[1]-1]} ${p[0]}` : iso; }
function fmtEn(iso){ const p = String(iso).split('-'); return p.length === 3 ? `${+p[2]} ${EN_MONTHS[+p[1]-1]} ${p[0]}` : iso; }
function esc(s){ return String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
const safeUrl = u => /^https?:\/\//i.test(u || '') ? u : '';
const siteBase = () => location.origin + location.pathname.replace(/[^/]*$/, '');
// Sessions are bound to badge type: VIP → morning, guest → evening, press → both
const SESS_AR = { am: 'الفترة الصباحية', pm: 'الفترة المسائية' };
const TYPE_SESS = { vip: ['am'], public: ['pm'], press: ['am', 'pm'] };
const TYPE_LOOK = {
  vip:    { ar: 'كبار الضيوف', en: 'VIP',   color: '#e6c36a' },
  public: { ar: 'ضيف',         en: 'GUEST', color: '#1cc8ee' },
  press:  { ar: 'إعلام',        en: 'PRESS', color: '#ff9a6b' },
};
function sessTime(ev, s){ return s === 'am' ? ev.time : ev.time_pm; }
function typeSessions(ev, type){ return (TYPE_SESS[type] || ['pm']).map(s => ({ s, label: SESS_AR[s], time: sessTime(ev, s) || '' })); }
function typeTimesText(ev, type){ return typeSessions(ev, type).map(x => x.label + (x.time ? ' — ' + x.time : '')).join(' · '); }
