// Reads a CamEd Faculty Portal CV (PDF export) into structured data, then summarises it
// against the Faculty Classification Policy codes for a given academic year.
const DATE = /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{1,2}, (\d{4})$/;
const DATE_IN = /(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{1,2}, (\d{4})/;
const SECTIONS = {
  'ACADEMIC DEGREE': 'date', 'PROFESSIONAL DEGREE': 'date', 'TRAINING COURSE/WORKSHOP /SEMINAR': 'date',
  'CONFERENCE /WEBINAR /SYMPOSIUM': 'date', 'ACADEMIC POSITION': 'date', 'OTHER POSITION': 'date',
  'RESEARCH PUBLICATION': 'year', 'SUPERVISION OF STUDENT PROJECTS /RESEARCH /CLUB': 'date', 'CURRICULUM DEVELOPMENT': 'date',
  'BUSINESS PROFESSIONAL ASSOCIATION': 'date', 'CONSULTANCY PROJECT': 'date', 'NETWORKING EVENT': 'date'
};
const OTHER = ['QUALIFICATION', 'TRAINING AND DEVELOPMENT', 'WORK EXPERIENCE', 'ACADEMIC ENGAGEMENT ACTIVITY', 'PROFESSIONAL ENGAGEMENT ACTIVITY'];
const HEAD_WORDS = new Set('institution country degree field of study start date end training topic name conference/webinar/symposium role position title research year publication category journal/book type journal author student project/research/club program/course business professional association funding organization consultancy event'.split(' '));
const key = t => t.replace(/\s+/g, '').toUpperCase();
const SKEYS = Object.keys(SECTIONS).reduce((o, k) => { o[key(k)] = k; return o; }, {});
const OKEYS = new Set(OTHER.map(key));

async function readLines(pdfjs, buf) {
  const doc = await pdfjs.getDocument({ data: new Uint8Array(buf) }).promise;
  const lines = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p), vp = page.getViewport({ scale: 1 }), tc = await page.getTextContent();
    const items = tc.items.filter(i => i.str && i.str.trim()).map(i => ({ s: i.str.trim(), x: i.transform[4], y: p * 10000 + (vp.height - i.transform[5]), w: i.width || 0 }));
    items.sort((a, b) => a.y - b.y || a.x - b.x);
    for (const it of items) {
      let L = lines[lines.length - 1];
      if (!L || Math.abs(L.y - it.y) > 2.5) { L = { y: it.y, items: [] }; lines.push(L); }
      L.items.push(it);
    }
  }
  lines.forEach(L => {
    L.items.sort((a, b) => a.x - b.x);
    const m = [];
    L.items.forEach(it => { const last = m[m.length - 1]; const gap = last ? it.x - (last.x + last.w) : 99;
      if (last && gap < 3.5) { last.s += (gap > 0.6 ? ' ' : '') + it.s; last.w = it.x + it.w - last.x; } else m.push({ ...it }); });
    L.items = m; L.text = m.map(i => i.s).join(' ').replace(/\s+/g, ' ').trim();
  });
  return lines;
}

function isHeaderLine(L) { const words = L.text.toLowerCase().split(/\s+/).filter(Boolean); return words.length > 0 && words.every(w => HEAD_WORDS.has(w)); }

function parseTable(lines, mode) {
  if (!lines.length) return [];
  const head = lines.find(isHeaderLine);
  if (!head) return [];
  const starts = head.items.map(i => i.x);
  const colOf = x => { let c = 0; starts.forEach((s, i) => { if (x >= s - 10) c = i; }); return c; };
  const data = lines.filter(L => !isHeaderLine(L) && !/^No records found/i.test(L.text));
  const isAnchor = L => L.items.some(i => mode === 'year' ? /^(19|20)\d{2}\b/.test(i.s) && colOf(i.x) === 1 : DATE.test(i.s) || DATE_IN.test(i.s));
  const anchors = data.filter(isAnchor);
  if (!anchors.length) return [];
  const rows = anchors.map(() => starts.map(() => []));
  data.forEach(L => {
    let best = 0, bd = Infinity;
    if (mode === 'year') { anchors.forEach((A, i) => { if (A.y <= L.y + 0.5) best = i; }); }
    else anchors.forEach((A, i) => { const d = Math.abs(A.y - L.y); if (d < bd) { bd = d; best = i; } });
    L.items.forEach(it => rows[best][colOf(it.x)].push(it.s));
  });
  return rows.map(r => r.map(c => c.join(' ').replace(/\s+/g, ' ').trim()));
}

export async function parseCV(pdfjs, buf) {
  const lines = await readLines(pdfjs, buf);
  const header = [], secs = {}; let cur = null, inHeader = true;
  lines.forEach(L => {
    const k = key(L.text);
    if (SKEYS[k]) { cur = SKEYS[k]; secs[cur] = []; inHeader = false; return; }
    if (OKEYS.has(k)) { cur = null; inHeader = false; return; }
    if (inHeader) header.push(L); else if (cur) secs[cur].push(L);
  });
  const labels = [], vals = [];
  header.forEach(L => L.items.forEach(it => { if (/:\s*$/.test(it.s)) labels.push({ k: it.s.replace(/\s*:\s*$/, '').trim(), y: L.y, v: [] }); else vals.push({ s: it.s, y: L.y }); }));
  vals.forEach(v => { let best = null, bd = Infinity; labels.forEach(l => { const d = Math.abs(l.y - v.y); if (d < bd) { bd = d; best = l; } }); if (best) best.v.push(v); });
  const info = {}; labels.forEach(l => { info[l.k] = l.v.sort((a, b) => a.y - b.y).map(v => v.s).join(' ').replace(/\s+/g, ' ').trim(); });
  const tables = {}; Object.keys(SECTIONS).forEach(k => { tables[k] = parseTable(secs[k] || [], SECTIONS[k]); });
  return { info, tables };
}

const yearOf = s => { const m = String(s || '').match(DATE_IN); return m ? +m[2] : (/present/i.test(s || '') ? 9999 : null); };
const norm = s => String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();

export function summarizeCV(cv, opts) {
  const Y = opts.year, from = Y - 5, names = (opts.names || []).map(n => ({ n, parts: norm(n).split(' ').filter(p => p.length > 1) }));
  const inWin = (start, end) => { const a = yearOf(start), b = yearOf(end) ?? a; return a != null && a <= Y && (b ?? a) >= from; };
  const T = cv.tables, out = { info: cv.info, counts: {}, pubs: [], notes: [], items: {} };
  const cnt = (k, v) => { if (v) out.counts[k] = (out.counts[k] || 0) + v; };
  const det = (k, rows, fn) => { if (rows.length) (out.items[k] ||= []).push(...rows.map(fn)); };
  const when = (a, b) => (a && b && a !== b ? a + ' – ' + b : a || '');
  const dateRows = (k, si, ei) => (T[k] || []).filter(r => inWin(r[si], r[ei]));
  const tr = dateRows('TRAINING COURSE/WORKSHOP /SEMINAR', 3, 4); cnt('P9', tr.length); det('P9', tr, r => ({ t: r[2], sub: [r[0], r[1], when(r[3], r[4])].filter(Boolean).join(' · ') }));
  const conf = dateRows('CONFERENCE /WEBINAR /SYMPOSIUM', 3, 4);
  const speak = conf.filter(r => /speaker|presenter|trainer|panel|keynote|resource|facilitator|moderator/i.test(r[2]));
  cnt('A16', speak.length); cnt('P10', conf.length - speak.length);
  const cf = r => ({ t: r[0], sub: [r[2], r[1], when(r[3], r[4])].filter(Boolean).join(' · ') });
  det('A16', speak, cf); det('P10', conf.filter(r => !speak.includes(r)), cf);
  const nw = dateRows('NETWORKING EVENT', 2, 3); cnt('P10', nw.length); det('P10', nw, r => ({ t: r[1], sub: ['Networking', r[0], when(r[2], r[3])].filter(Boolean).join(' · ') }));
  const sv = dateRows('SUPERVISION OF STUDENT PROJECTS /RESEARCH /CLUB', 2, 3).filter(r => !/club/i.test(r[1])); cnt('A15', sv.length); det('A15', sv, r => ({ t: r[1], sub: [r[0], when(r[2], r[3])].filter(Boolean).join(' · ') }));
  const as = dateRows('BUSINESS PROFESSIONAL ASSOCIATION', 2, 3); cnt('P6', as.length); det('P6', as, r => ({ t: r[1], sub: [r[0], when(r[2], r[3])].filter(Boolean).join(' · ') }));
  const cs = dateRows('CONSULTANCY PROJECT', 3, 4); cnt('P7', cs.length); det('P7', cs, r => ({ t: r[2], sub: [r[1], r[0], when(r[3], r[4])].filter(Boolean).join(' · ') }));
  const acts = []; let aid = 0;
  const act = (src, code, t, sub) => acts.push({ id: 'x' + (aid++), src, code, t: t || '(untitled)', sub });
  tr.forEach(r => act('Training and development', 'P9', r[2], [r[0], r[1], when(r[3], r[4])].filter(Boolean).join(' · ')));
  conf.forEach(r => act('Conferences, webinars and symposiums', speak.includes(r) ? 'A16' : 'P10', r[0], [r[2], r[1], when(r[3], r[4])].filter(Boolean).join(' · ')));
  nw.forEach(r => act('Networking events', 'P10', r[1], [r[0], when(r[2], r[3])].filter(Boolean).join(' · ')));
  dateRows('SUPERVISION OF STUDENT PROJECTS /RESEARCH /CLUB', 2, 3).forEach(r => act('Supervision of students', /club/i.test(r[1]) ? 'skip' : 'A15', r[1], [r[0], when(r[2], r[3])].filter(Boolean).join(' · ')));
  as.forEach(r => act('Professional associations', 'P6', r[1], [r[0], when(r[2], r[3])].filter(Boolean).join(' · ')));
  cs.forEach(r => act('Consultancy projects', 'P7', r[2], [r[1], r[0], when(r[3], r[4])].filter(Boolean).join(' · ')));
  ['ACADEMIC POSITION', 'OTHER POSITION'].forEach(k => dateRows(k, 2, 3).forEach(r => {
    const lead = /manager|director|head|chair|president|dean|ceo|partner|lead|founder|owner|vice/i.test(r[0]), camed = /camed/i.test(r[1]);
    act('Work experience', /owner|founder/i.test(r[0]) ? 'P4' : !camed ? (lead ? 'P5' : 'P1') : (lead && k === 'OTHER POSITION' ? 'skip' : 'skip'), r[0], [r[1], when(r[2], r[3])].filter(Boolean).join(' · '));
  }));
  dateRows('CURRICULUM DEVELOPMENT', 2, 3).forEach(r => act('Curriculum development', 'skip', r[1], [r[0], when(r[2], r[3])].filter(Boolean).join(' · ')));
  out.acts = acts;
  const certs = [...new Set((T['PROFESSIONAL DEGREE'] || []).map(r => r[2]).filter(d => d && !/diploma/i.test(d)))];
  out.certs = certs.join('; ');
  (T['RESEARCH PUBLICATION'] || []).forEach(r => {
    const y = +(String(r[1]).match(/(19|20)\d{2}/) || [])[0]; if (!y || y < from || y > Y) return;
    const cat = r[2], type = r[4], journal = r[3];
    const p = /basic|discovery/i.test(cat) ? 'B' : /teaching|learning/i.test(cat) ? 'D' : 'C';
    const ty = /^peer/i.test(type) ? 'F' : /additional/i.test(type) ? 'G' : 'H';
    const authors = String(r[5] || '').split(/,|;| and /).map(a => a.trim()).filter(Boolean);
    const camed = names.filter(N => authors.some(a => { const ap = norm(a).split(' '); return N.parts.length && N.parts.filter(x => ap.includes(x)).length >= Math.min(2, N.parts.length); })).map(N => N.n);
    const jt = [journal, r[0]].join(' ');
    const code = /jafess/i.test(jt) ? 'A2' : ty === 'F' ? (/scopus|abdc/i.test(type) ? 'A1' : 'A2') : ty === 'G' ? (/proceeding|conference|lecture notes/i.test(journal) ? 'A4' : /\bbook\b|chapter|volume|press/i.test(journal) ? 'A8' : 'A3') : null;
    if (code) { cnt(code, 1); (out.items[code] ||= []).push({ t: r[0], sub: [y, journal, type].filter(Boolean).join(' · ') }); }
    out.pubs.push({ y, t: r[0], j: journal, au: r[5] || '', p, ty, c: Math.round(100 / Math.max(1, camed.length)) / 100, code, camed });
  });
  const cls = (String(cv.info['Faculty Qualification'] || '').match(/\((SA|PA|SP|IP|A)\)/) || [])[1] || '';
  const role = /supporting/i.test(cv.info['Faculty Sufficiency'] || '') ? 'S' : /participating/i.test(cv.info['Faculty Sufficiency'] || '') ? 'P' : '';
  const dg = String(cv.info['Terminal/Highest Degree'] || '').replace(/\s*degree\s*$/i, '').trim();
  out.fields = { name: cv.info['Faculty Name'] || '', title: cv.info['Academic Title'] || '', courses: cv.info['Teaching Course'] || '', deg: /doctor|phd/i.test(dg) ? 'PhD' : dg, discipline: cv.info['Discipline'] || '', cls, role };
  out.trainingCount = tr.length;
  return out;
}
