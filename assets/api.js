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
