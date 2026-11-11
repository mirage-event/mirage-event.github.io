// Reports, list export (Excel) and CRM export (CSV) — all computed in the admin's browser
// from the same `list` data the admin page already loads. Nothing here writes to the Sheet.
const RPT = (() => {
  const TYPES = ['vip', 'press', 'public'];
  const TYPE_AR = { vip: 'VIP', press: 'إعلام', public: 'ضيوف' };
  const ST_AR = { pending: 'قيد المراجعة', approved: 'معتمد', rejected: 'مرفوض' };
  const ATT_AR = { yes: 'أؤكد حضوري', no: 'لن أحضر', maybe: 'غير متأكد' };
  const SESS_LIST = s => (TYPE_SESS[s] || ['pm']);

  const inSess = (r, s) => s === 'am' ? !!r.checkin_at : !!r.checkin2_at;
  const attended = r => !!(r.checkin_at || r.checkin2_at);
  const expects = (r, s) => r.status === 'approved' && SESS_LIST(r.type).includes(s);

  // ---------- filtering (export builder) ----------
  // opts: { types:[...], status:'all|approved|attended|noshow|pending|rejected', session:'all|am|pm' }
  function filter(regs, opts) {
    const types = opts.types && opts.types.length ? opts.types : TYPES;
    const s = opts.session === 'am' || opts.session === 'pm' ? opts.session : null;
    return regs.filter(r => {
      if (!types.includes(r.type)) return false;
      if (s && !SESS_LIST(r.type).includes(s)) return false;
      switch (opts.status) {
        case 'approved': return r.status === 'approved';
        case 'pending':  return r.status === 'pending';
        case 'rejected': return r.status === 'rejected';
        case 'attended': return s ? inSess(r, s) : attended(r);
        case 'noshow':   return r.status === 'approved' && (s ? !inSess(r, s) : !attended(r));
        default: return true;
      }
    }).sort((a, b) => String(a.badge).localeCompare(String(b.badge)));
  }

  // ---------- summary numbers ----------
  function stats(regs) {
    const row = () => ({ total: 0, pending: 0, rejected: 0, approved: 0, declined: 0, am: 0, pm: 0, attended: 0, noshow: 0, expAm: 0, expPm: 0 });
    const by = { vip: row(), press: row(), public: row(), all: row() };
    regs.forEach(r => {
      [by[r.type] || by.public, by.all].forEach(x => {
        x.total++; x[r.status]++;
        if (r.status === 'approved' && r.attend === 'no') x.declined++;
        if (r.checkin_at) x.am++; if (r.checkin2_at) x.pm++;
        if (attended(r)) x.attended++;
        if (r.status === 'approved' && !attended(r)) x.noshow++;
        if (expects(r, 'am')) x.expAm++; if (expects(r, 'pm')) x.expPm++;
      });
    });
    Object.values(by).forEach(x => { x.rate = x.approved ? Math.round(100 * x.attended / x.approved) : 0; });
    return by;
  }

  // arrivals per 15 minutes for one session (or both)
  function arrivals(regs, sess) {
    const bins = {};
    regs.forEach(r => (sess === 'am' ? [r.checkin_at] : sess === 'pm' ? [r.checkin2_at] : [r.checkin_at, r.checkin2_at]).forEach(t => {
      if (!t) return;
      const [h, m] = t.slice(11, 16).split(':').map(Number); if (isNaN(h)) return;
      const k = h * 60 + Math.floor(m / 15) * 15; bins[k] = (bins[k] || 0) + 1;
    }));
    const keys = Object.keys(bins).map(Number).sort((a, b) => a - b);
    if (!keys.length) return [];
    const out = []; for (let k = keys[0]; k <= keys[keys.length - 1]; k += 15) out.push({ min: k, n: bins[k] || 0 });
    return out;
  }
  const hhmm = min => String(Math.floor(min / 60)).padStart(2, '0') + ':' + String(min % 60).padStart(2, '0');

  // ---------- Excel ----------
  const COLS = [
    ['رقم البطاقة', r => r.badge], ['النوع', r => TYPE_AR[r.type]], ['الحالة', r => ST_AR[r.status]],
    ['الاسم', r => r.name], ['الهاتف', r => r.phone], ['الإيميل', r => r.email], ['المسمى الوظيفي', r => r.title],
    ['الجهة', r => r.org], ['الجنس', r => r.gender], ['تأكيد الحضور', r => ATT_AR[r.attend]],
    ['الفترات المسموحة', r => SESS_LIST(r.type).map(s => SESS_AR[s]).join(' + ')],
    ['وقت التسجيل', r => r.created], ['القرار بواسطة', r => r.decided_by],
    ['دخول الصباحية', r => r.checkin_at], ['بواسطة', r => r.checkin_by],
    ['دخول المسائية', r => r.checkin2_at], ['بواسطة ', r => r.checkin2_by],
  ];
  function sheetOf(rows) {
    const ws = XLSX.utils.aoa_to_sheet([COLS.map(c => c[0]), ...rows.map(r => COLS.map(c => c[1](r) ?? ''))]);
    ws['!cols'] = COLS.map(([h]) => ({ wch: ['الاسم', 'الجهة', 'المسمى الوظيفي', 'الإيميل'].includes(h) ? 26 : 15 }));
    ws['!autofilter'] = { ref: ws['!ref'] };
    return ws;
  }
  function summaryAoa(ev, regs) {
    const st = stats(regs);
    const line = (k, label) => { const x = st[k]; return [label, x.total, x.pending, x.rejected, x.approved, x.declined,
      k === 'public' ? '—' : x.am, k === 'vip' ? '—' : x.pm, x.attended, x.noshow, x.rate + '%']; };
    return [
      [ev.name_ar || ''], [[fmtAr(ev.date), ev.venue_ar].filter(Boolean).join(' · ')],
      ['تاريخ التقرير', new Date().toLocaleString('ar-LY')], [],
      ['النوع', 'المسجلين', 'قيد المراجعة', 'مرفوض', 'معتمد', 'معتمد واعتذر', 'حضروا الصباحية', 'حضروا المسائية', 'حضروا (أي فترة)', 'لم يحضروا', 'نسبة الحضور'],
      line('vip', 'VIP'), line('press', 'إعلام'), line('public', 'ضيوف'), line('all', 'الإجمالي'), [],
      ['نسبة الحضور = من حضر في أي فترة ÷ المعتمدين. «لم يحضروا» = معتمدون لم يُمسحوا في أي فترة.'],
    ];
  }
  function save(wb, name) { wb.Workbook = { Views: [{ RTL: true }] }; XLSX.writeFile(wb, name); }
  function excelList(ev, rows, label) {
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, sheetOf(rows), 'القائمة');
    save(wb, `${ev.code}-${label}.xlsx`);
  }
  function excelFull(ev, regs) {
    const wb = XLSX.utils.book_new();
    const sum = XLSX.utils.aoa_to_sheet(summaryAoa(ev, regs)); sum['!cols'] = [{ wch: 18 }, ...Array(10).fill({ wch: 14 })];
    XLSX.utils.book_append_sheet(wb, sum, 'الملخص');
    XLSX.utils.book_append_sheet(wb, sheetOf(filter(regs, { status: 'all' })), 'كل المسجلين');
    TYPES.forEach(t => XLSX.utils.book_append_sheet(wb, sheetOf(filter(regs, { types: [t], status: 'all' })), TYPE_AR[t]));
    XLSX.utils.book_append_sheet(wb, sheetOf(filter(regs, { status: 'attended' })), 'الحضور');
    XLSX.utils.book_append_sheet(wb, sheetOf(filter(regs, { status: 'noshow' })), 'لم يحضروا');
    save(wb, `${ev.code}-تقرير-كامل.xlsx`);
  }

  // ---------- CRM CSV (matches the CRM's «استيراد ليدز» file: Meta Lead Ads column names) ----------
  // Same rule the CRM applies: Libyan 09[1-5]xxxxxxx, or an international number written with + / 00.
  function crmPhone(raw) {
    let s = String(raw || '').replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[\s\-().]/g, '');
    if (s.startsWith('00')) s = '+' + s.slice(2);
    let m = s.match(/^(?:\+218|218)?0?(9[1-5]\d{7})$/); if (m) return '+218' + m[1];
    m = s.match(/^\+(\d{8,15})$/); if (m && !m[1].startsWith('218')) return '+' + m[1];
    return null;
  }
  function crmRows(ev, regs, types) {
    // rejected requests never go to sales
    const rows = regs.filter(r => types.includes(r.type) && r.status !== 'rejected');
    const camp = `حدث: ${ev.name_ar} — ${ev.date}`;
    return rows.map(r => {
      const att = r.checkin_at || r.checkin2_at ? 'حضر' : r.status === 'approved' ? 'لم يحضر' : 'لم يُعتمد';
      return { ok: !!crmPhone(r.phone), cells: [
        'rsvp-' + r.badge,
        r.created ? r.created.replace(' ', 'T') + ':00+02:00' : '',
        r.name, crmPhone(r.phone) || r.phone, r.email || '',
        camp, `${TYPE_AR[r.type]} · ${att}${r.org ? ' · ' + r.org : ''}`, 'MIRAGE RSVP', 'Event',
      ] };
    });
  }
  function crmCsv(ev, regs, types, label) {
    const rows = crmRows(ev, regs, types);
    const head = ['id', 'created_time', 'full_name', 'phone_number', 'email', 'campaign_name', 'ad_name', 'form_name', 'platform'];
    const q = v => '"' + String(v ?? '').replace(/"/g, '""') + '"';
    const csv = '﻿' + [head, ...rows.map(r => r.cells)].map(l => l.map(q).join(',')).join('\r\n');
    const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
    a.download = `${ev.code}-CRM-${label}.csv`; a.click();
    return { total: rows.length, bad: rows.filter(r => !r.ok).length };
  }

  return { TYPES, TYPE_AR, filter, stats, arrivals, hhmm, excelList, excelFull, crmRows, crmCsv, crmPhone };
})();
