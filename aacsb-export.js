// Fills the AACSB tables workbook template (Table 3-1, 3-2, 8-1) and adds Tables 3-1A, 3-2b and 8-1A from portal data.
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const COLS = ['E', 'F', 'G', 'H', 'I', 'J', 'K'];
const CLS_COL = { SA: 'G', PA: 'H', SP: 'I', IP: 'J', A: 'K' };

function rowOf(xml, r) { const m = xml.match(new RegExp(`<row r="${r}"[^>]*>[\\s\\S]*?</row>`)); return m ? m[0] : ''; }
function cell(col, r, sa, v) {
  const ref = `${col}${r}`;
  if (v == null || v === '') return `<c r="${ref}"${sa}/>`;
  if (typeof v === 'number') return `<c r="${ref}"${sa}><v>${isFinite(v) ? v : 0}</v></c>`;
  if (typeof v === 'object') return `<c r="${ref}"${sa}><f>${esc(v.f)}</f><v>${isFinite(v.v) ? v.v : 0}</v></c>`;
  return `<c r="${ref}"${sa} t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
}
function mk(tpl, r, vals = {}, o = {}) {
  let row = tpl.replace(/^<row r="\d+"/, `<row r="${r}"`);
  if (o.ht) row = row.replace(/ ht="[^"]*"/, '').replace(/ customHeight="1"/, '').replace(/^<row /, `<row ht="${o.ht}" customHeight="1" `);
  return row.replace(/<c r="([A-Z]+)\d+"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, (m, col, attrs, inner) => {
    const s = (attrs.match(/ s="(\d+)"/) || [])[1];
    const st = o.styles && o.styles[col] != null ? o.styles[col] : s;
    const sa = st != null ? ` s="${st}"` : '';
    if (col in vals) return cell(col, r, sa, vals[col]);
    if (o.keep && inner) return `<c r="${col}${r}"${attrs}>${inner}</c>`;
    return `<c r="${col}${r}"${sa}/>`;
  });
}
function shiftRow(row, d) {
  return row.replace(/^<row r="(\d+)"/, (m, n) => `<row r="${+n + d}"`).replace(/<c r="([A-Z]+)(\d+)"/g, (m, c, n) => `<c r="${c}${+n + d}"`);
}
const pctTxt = v => Math.round(v * 100) + '%';
const metTxt = (label, v, t) => `${label} >= ${t}% ${v * 100 >= t - 1e-9 ? 'met' : 'not met'} (${pctTxt(v)})`;
const div = (a, b) => (b ? a / b : 0);

function buildTable31(xml, m) {
  const head = xml.slice(0, xml.indexOf('<sheetData>') + 11).replace(/<dimension ref="[^"]*"\/>/, '');
  const tail = xml.slice(xml.indexOf('</sheetData>'));
  const T = n => rowOf(xml, n);
  const out = [];
  for (let r = 1; r <= 7; r++) out.push(r === 4 ? mk(T(4), 4, { A: `Faculty Sufficiency and Qualifications Summary for January-December ${m.year}` }, { keep: true }) : T(r));
  const merges = ['A1:L1', 'A3:L3', 'A4:L4', 'H5:J5', 'E6:F6', 'G6:K6', 'L6:L7', 'A6:D6'];
  let r = 8; const totals = []; const grand = { E: 0, F: 0, G: 0, H: 0, I: 0, J: 0, K: 0 }; let count = 0;
  m.disciplines.forEach(d => {
    const h1 = r, h2 = r + 1;
    out.push(mk(d.offered ? T(8) : T(36), h1, { A: `Discipline ${d.letter}: ${d.name}`, C: d.offered ? 'Degrees or majors are offered in this discipline. ☑ YES' : 'Degrees or majors are offered in this discipline.' }));
    out.push(mk(d.offered ? T(9) : T(37), h2, { C: d.offered ? 'Degrees or majors are not offered in this discipline.' : 'Degrees or majors are not offered in this discipline. ☑ YES' }));
    merges.push(`A${h1}:A${h2}`, `C${h1}:L${h1}`, `C${h2}:L${h2}`);
    r += 2; const first = r; const sum = { E: 0, F: 0, G: 0, H: 0, I: 0, J: 0, K: 0 };
    d.faculty.forEach(f => {
      const v = { A: f.name, B: f.spec || '', C: f.deg, D: f.yr != null ? String(f.yr) : '', E: f.P || null, F: f.S || null, L: f.basis || null };
      v[CLS_COL[f.cls]] = f.pct;
      const styles = {}; COLS.forEach(c => { styles[c] = v[c] != null ? 28 : 29; });
      ['E', 'F'].forEach(c => { sum[c] += v[c] || 0; }); sum[CLS_COL[f.cls]] += f.pct;
      const lines = String(f.basis || '').split('\n').reduce((s, l) => s + Math.max(1, Math.ceil(l.length / 55)), 0);
      out.push(mk(T(10), r, v, { styles, ht: Math.max(26.4, lines * 13.2) }));
      r++;
    });
    const last = r - 1, t = r;
    const tv = { A: 'Total ' + d.name, B: d.faculty.length };
    COLS.forEach(c => { tv[c] = { f: `SUM(${c}${first}:${c}${last})`, v: sum[c] }; grand[c] += sum[c]; });
    out.push(mk(T(24), t, tv));
    const ps = sum.E + sum.F, qs = sum.G + sum.H + sum.I + sum.J + sum.K;
    const pv = {};
    ['E', 'F'].forEach(c => { pv[c] = { f: `IFERROR(${c}${t}/SUM($E$${t}:$F$${t}),0)`, v: div(sum[c], ps) }; });
    ['G', 'H', 'I', 'J', 'K'].forEach(c => { pv[c] = { f: `IFERROR(${c}${t}/SUM($G$${t}:$K$${t}),0)`, v: div(sum[c], qs) }; });
    out.push(mk(T(25), t + 1, pv));
    out.push(mk(T(26), t + 2, { A: d.name + ' Ratio', E: metTxt('P', div(sum.E, ps), 60), G: metTxt('SA', div(sum.G, qs), 40) + '\n' + metTxt('SA+PA+SP+IP', div(qs - sum.K, qs), 90) }, { ht: 26.4 }));
    merges.push(`E${t + 2}:F${t + 2}`, `G${t + 2}:L${t + 2}`);
    totals.push(t); count += d.faculty.length; r = t + 3;
  });
  const g = r, gv = {};
  COLS.forEach(c => { gv[c] = { f: totals.map(t => c + t).join('+') || '0', v: grand[c] }; });
  out.push(mk(T(84), g, gv));
  const ps = grand.E + grand.F, qs = grand.G + grand.H + grand.I + grand.J + grand.K;
  const pv = { A: 'Grand Total', B: count };
  ['E', 'F'].forEach(c => { pv[c] = { f: `IFERROR(${c}${g}/SUM($E$${g}:$F$${g}),0)`, v: div(grand[c], ps) }; });
  ['G', 'H', 'I', 'J', 'K'].forEach(c => { pv[c] = { f: `IFERROR(${c}${g}/SUM($G$${g}:$K$${g}),0)`, v: div(grand[c], qs) }; });
  out.push(mk(T(85), g + 1, pv));
  out.push(mk(T(86), g + 2, { A: 'Overall Ratio', E: metTxt('P', div(grand.E, ps), 75), G: metTxt('SA', div(grand.G, qs), 40) + '\n' + metTxt('SA+PA+SP+IP', div(qs - grand.K, qs), 90) }, { ht: 26.4 }));
  out.push(mk(T(87), g + 3, {}, { keep: true }));
  merges.push(`E${g + 2}:F${g + 2}`, `G${g + 2}:L${g + 2}`, `G${g + 3}:L${g + 3}`);
  const end = g + 3;
  return head.replace('<sheetViews>', `<dimension ref="A1:R${end}"/><sheetViews>`) + out.join('') +
    tail.replace(/<mergeCells count="\d+">[\s\S]*?<\/mergeCells>/, `<mergeCells count="${merges.length}">${merges.map(x => `<mergeCell ref="${x}"/>`).join('')}</mergeCells>`);
}

function buildTable32(xml, m) {
  xml = xml.replace(rowOf(xml, 4), mk(rowOf(xml, 4), 4, { A: `Deployment of Faculty by Qualification Status in Support of Degree Programs in AY ${m.year}` }, { keep: true }));
  m.programs.forEach((p, i) => {
    const r = 8 + i, tpl = rowOf(xml, r);
    if (!tpl) return;
    xml = xml.replace(tpl, mk(tpl, r, { B: p.SA, C: p.PA, D: p.SP, E: p.IP, F: p.A, G: p.total }, { keep: true }));
  });
  return xml;
}

function buildTable81(xml, m) {
  xml = xml.replace(rowOf(xml, 4), mk(rowOf(xml, 4), 4, { A: ` Intellectual Contributions Data Date Range: Jan ${m.year - 5} - Dec ${m.year}` }, { keep: true }));
  const start = xml.indexOf('<sheetData>') + 11, stop = xml.indexOf('</sheetData>');
  const rows = xml.slice(start, stop).match(/<row [\s\S]*?<\/row>|<row [^>]*\/>/g) || [];
  const rn = row => +row.match(/^<row r="(\d+)"/)[1];
  const tplD = rows.find(x => rn(x) === 10), tplT = rows.find(x => rn(x) === 17);
  const n = m.ic.length, totalR = 10 + n, delta = totalR - 17;
  const out = [];
  rows.filter(x => rn(x) < 10).forEach(x => out.push(x));
  const tot = { B: 0, C: 0, D: 0, F: 0, G: 0, H: 0, pf: 0, pfIC: 0, fte: 0, fteIC: 0 };
  m.ic.forEach((d, i) => {
    const E = d.B + d.C + d.D, I = d.F + d.G + d.H;
    out.push(mk(tplD, 10 + i, { A: d.name, B: d.B, C: d.C, D: d.D, E, F: d.F, G: d.G, H: d.H, I, J: div(d.pfIC, d.pf), K: div(d.fteIC, d.fte) }));
    Object.keys(tot).forEach(k => { tot[k] += d[k] || 0; });
  });
  const tv = { A: 'Total' };
  ['B', 'C', 'D', 'F', 'G', 'H'].forEach(c => { tv[c] = { f: `SUM(${c}10:${c}${totalR - 1})`, v: tot[c] }; });
  tv.E = { f: `SUM(E10:E${totalR - 1})`, v: tot.B + tot.C + tot.D };
  tv.I = { f: `SUM(I10:I${totalR - 1})`, v: tot.F + tot.G + tot.H };
  tv.J = div(tot.pfIC, tot.pf); tv.K = div(tot.fteIC, tot.fte);
  out.push(mk(tplT, totalR, tv));
  rows.filter(x => rn(x) > 17).forEach(x => out.push(shiftRow(x, delta)));
  let res = xml.slice(0, start) + out.join('') + xml.slice(stop);
  res = res.replace(/<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/g, (mm, a, r1, b, r2) => +r1 > 17 ? `<mergeCell ref="${a}${+r1 + delta}:${b}${+r2 + delta}"/>` : mm);
  res = res.replace(/<brk id="(\d+)" man="1"\/><\/rowBreaks>/, (mm, id) => `<brk id="${+id + delta}" man="1"/></rowBreaks>`);
  res = res.replace(/<dimension ref="A1:([A-Z]+)(\d+)"\/>/, (mm, c, r) => `<dimension ref="A1:${c}${+r + delta}"/>`);
  return res;
}



// ---- Generated sheets (Table 3-1A, 3-2b, 8-1A) ----
const colL = i => { let s = ''; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
function addStyles(st) {
  const push = (plural, xml) => {
    const re = new RegExp('<' + plural + ' count="(\\d+)"([^>]*)>'); const m = st.match(re); const n = +m[1];
    st = st.replace(re, () => '<' + plural + ' count="' + (n + 1) + '"' + m[2] + '>'); const k = st.indexOf('</' + plural + '>'); st = st.slice(0, k) + xml + st.slice(k); return n;
  };
  if (!/<numFmts /.test(st)) st = st.replace(/(<styleSheet[^>]*>)/, '$1<numFmts count="0"></numFmts>');
  push('numFmts', '<numFmt numFmtId="180" formatCode="0.0%"/>');
  push('numFmts', '<numFmt numFmtId="181" formatCode="0.00"/>');
  const F = (sz, b) => push('fonts', '<font>' + (b ? '<b/>' : '') + '<sz val="' + sz + '"/><color rgb="FF000000"/><name val="Arial"/><family val="2"/></font>');
  const f9 = F(9), f9b = F(9, 1), f12b = F(12, 1);
  const fill = rgb => push('fills', '<fill><patternFill patternType="solid"><fgColor rgb="' + rgb + '"/><bgColor indexed="64"/></patternFill></fill>');
  const gHead = fill('FFD9D9D9'), gDisc = fill('FFEDEDED'), gTot = fill('FFF2F2F2');
  const bd = push('borders', '<border><left style="thin"><color rgb="FF999999"/></left><right style="thin"><color rgb="FF999999"/></right><top style="thin"><color rgb="FF999999"/></top><bottom style="thin"><color rgb="FF999999"/></bottom><diagonal/></border>');
  const X = (font, fl, b, nf, h, v) => push('cellXfs', '<xf numFmtId="' + nf + '" fontId="' + font + '" fillId="' + fl + '" borderId="' + b + '" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="' + h + '" vertical="' + (v || 'center') + '" wrapText="1"/></xf>');
  const S = {
    title: X(f12b, 0, 0, 0, 'left'), sub: X(f9, 0, 0, 0, 'left'),
    hdr: X(f9b, gHead, bd, 0, 'center'), disc: X(f9b, gDisc, bd, 0, 'left'),
    txt: X(f9, 0, bd, 0, 'left', 'top'), cen: X(f9, 0, bd, 0, 'center'), num: X(f9, 0, bd, 0, 'center'), dec: X(f9, 0, bd, 181, 'center'), pct: X(f9, 0, bd, 9, 'center'), pct1: X(f9, 0, bd, 180, 'center'),
    txtB: X(f9b, gTot, bd, 0, 'left'), numB: X(f9b, gTot, bd, 0, 'center'), decB: X(f9b, gTot, bd, 181, 'center'), pctB: X(f9b, gTot, bd, 9, 'center'), pct1B: X(f9b, gTot, bd, 180, 'center')
  };
  return { st, S };
}
function cellX(ref, s, v) {
  const sa = s != null ? ' s="' + s + '"' : '';
  if (v == null || v === '') return '<c r="' + ref + '"' + sa + '/>';
  if (typeof v === 'number') return '<c r="' + ref + '"' + sa + '><v>' + (isFinite(v) ? v : 0) + '</v></c>';
  if (typeof v === 'object') return typeof v.v === 'string' ? '<c r="' + ref + '"' + sa + ' t="str"><f>' + esc(v.f) + '</f><v>' + esc(v.v) + '</v></c>' : '<c r="' + ref + '"' + sa + '><f>' + esc(v.f) + '</f><v>' + (isFinite(v.v) ? v.v : 0) + '</v></c>';
  return '<c r="' + ref + '"' + sa + ' t="inlineStr"><is><t xml:space="preserve">' + esc(v) + '</t></is></c>';
}
function sheetXml(rows, widths, merges, landscape) {
  const ncol = Math.max(...rows.map(r => r.c.length));
  const body = rows.map((r, i) => '<row r="' + (i + 1) + '"' + (r.ht ? ' ht="' + r.ht + '" customHeight="1"' : '') + '>' +
    r.c.map((c, j) => c ? cellX(colL(j) + (i + 1), c[1], c[0]) : '').join('') + '</row>').join('');
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:' + colL(ncol - 1) + rows.length + '"/><sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="13.2"/>' +
    '<cols>' + widths.map((w, i) => '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>').join('') + '</cols>' +
    '<sheetData>' + body + '</sheetData>' + (merges.length ? '<mergeCells count="' + merges.length + '">' + merges.map(m => '<mergeCell ref="' + m + '"/>').join('') + '</mergeCells>' : '') +
    '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="' + (landscape ? 'landscape' : 'portrait') + '" fitToHeight="0"/></worksheet>';
}
const lineHt = (txt, perLine) => Math.max(13.2, String(txt || '').split('\n').reduce((s, l) => s + Math.max(1, Math.ceil(l.length / perLine)), 0) * 12);
const R2 = v => Math.round(v * 10000) / 10000;
const pctF = (c, t, a, b) => 'IFERROR(' + c + t + '/SUM($' + a + '$' + t + ':$' + b + '$' + t + '),0)';

function buildTable31A(m, S) {
  const rows = [], merges = [], N = 21, blank = () => new Array(N).fill(null);
  const put = (r, j, v, s) => { r.c[j] = [v, s]; };
  rows.push({ ht: 20, c: [['Table 3-1A: Faculty Sufficiency, Qualifications and Basis for Qualification (60/30/10 Model), January–December ' + m.year, S.title]] });
  rows.push({ c: [['Teaching share = teaching hours ÷ 270 × 60% (max 60%). Research share = ICs ÷ 4 × 30% (max 30%); each faculty author gets full credit (1 per output) in Tables 3-1 and 3-1A. Service: 10% full-time, 5% part-time.', S.sub]] });
  const g = { ht: 30, c: blank().map(() => [null, S.hdr]) };
  put(g, 0, 'Faculty Portfolio by Discipline', S.hdr); put(g, 4, 'Faculty Sufficiency Related to Teaching (SCH)', S.hdr); put(g, 6, 'Faculty Qualifications with Respect to Percent of Time Devoted to Mission', S.hdr);
  put(g, 11, 'Brief Description of Basis for Qualification', S.hdr); put(g, 12, 'Time Devoted to Mission (60/30/10 Model)', S.hdr); put(g, 16, 'Workload and Status', S.hdr); put(g, 20, 'Note', S.hdr);
  rows.push(g); merges.push('A3:D3', 'E3:F3', 'G3:K3', 'L3:L4', 'M3:P3', 'Q3:T3', 'U3:U4');
  rows.push({ ht: 52, c: ['Faculty Member Name', 'Specialty Field', 'Highest Degree', 'Year Earned', 'Participating Faculty Teaching Productivity (P)', 'Supporting Faculty Teaching Productivity (S)', 'Scholarly Academic (SA)', 'Practice Academic (PA)', 'Scholarly Practitioner (SP)', 'Instructional Practitioner (IP)', 'Additional Faculty (A)', null, 'Total', 'Teaching (≤60%)', 'Research (≤30%)', 'Service (≤10%)', 'Teaching hours / year', 'ICs, full credit (' + (m.year - 5) + '–' + m.year + ')', 'FT / PT', 'FTE', null].map(v => [v, S.hdr]) });
  const CL = { SA: 6, PA: 7, SP: 8, IP: 9, A: 10 }, totRows = [], grand = new Array(N).fill(0); let count = 0;
  const isDec = j => (j >= 6 && j <= 10) || j === 19;
  m.disciplines.forEach(d => {
    const dr = rows.length + 1; rows.push({ ht: 16, c: blank().map((_, j) => [j ? null : 'Discipline ' + d.letter + ': ' + d.name + (d.offered ? '  (degrees or majors offered)' : ''), S.disc]) }); merges.push('A' + dr + ':U' + dr);
    const first = rows.length + 1, sum = new Array(N).fill(0);
    d.faculty.forEach(f => {
      const r = { ht: lineHt(f.basis, 60), c: blank() };
      [f.name, f.spec || '', f.deg, f.yr != null ? String(f.yr) : ''].forEach((v, j) => put(r, j, v, S.txt));
      put(r, 4, f.P || 0, S.num); put(r, 5, f.S || 0, S.num);
      [6, 7, 8, 9, 10].forEach(j => put(r, j, 0, S.dec)); put(r, CL[f.cls], f.pct, S.dec);
      put(r, 11, f.basis || '', S.txt);
      const sp = f.split || [0, 0, 0];
      put(r, 12, R2(f.pct / 100), S.pct1); put(r, 13, R2(sp[0] / 100), S.pct1); put(r, 14, R2(sp[1] / 100), S.pct1); put(r, 15, R2(sp[2] / 100), S.pct1);
      put(r, 16, f.hours || 0, S.num); put(r, 17, f.icFull || 0, S.num); put(r, 18, f.ft === 'PT' ? 'Part-time' : 'Full-time', S.cen); put(r, 19, R2(f.fte || 0), S.dec); put(r, 20, f.note || '', S.txt);
      sum[4] += f.P || 0; sum[5] += f.S || 0; sum[CL[f.cls]] += f.pct; sum[16] += f.hours || 0; sum[17] += f.icFull || 0; sum[19] += f.fte || 0;
      rows.push(r);
    });
    const last = rows.length, t = last + 1, tr = { c: blank().map(() => [null, S.numB]) };
    put(tr, 0, 'Total ' + d.name, S.txtB); put(tr, 1, d.faculty.length, S.numB);
    [4, 5, 6, 7, 8, 9, 10, 16, 17, 19].forEach(j => { const c = colL(j); put(tr, j, { f: 'SUM(' + c + first + ':' + c + last + ')', v: sum[j] }, isDec(j) ? S.decB : S.numB); grand[j] += sum[j]; });
    rows.push(tr); totRows.push(t);
    const ps = sum[4] + sum[5], qs = sum[6] + sum[7] + sum[8] + sum[9] + sum[10];
    const pr = { c: blank().map(() => [null, S.pctB]) }; put(pr, 0, d.name + ' %', S.txtB);
    [4, 5].forEach(j => put(pr, j, { f: pctF(colL(j), t, 'E', 'F'), v: div(sum[j], ps) }, S.pctB));
    [6, 7, 8, 9, 10].forEach(j => put(pr, j, { f: pctF(colL(j), t, 'G', 'K'), v: div(sum[j], qs) }, S.pctB));
    rows.push(pr);
    const rr = { ht: 26, c: blank().map(() => [null, S.txtB]) }; put(rr, 0, d.name + ' Ratio', S.txtB);
    put(rr, 4, metTxt('P', div(sum[4], ps), 60), S.txtB); put(rr, 6, metTxt('SA', div(sum[6], qs), 40) + '\n' + metTxt('SA+PA+SP+IP', div(qs - sum[10], qs), 90), S.txtB);
    rows.push(rr); const rn = rows.length; merges.push('E' + rn + ':F' + rn, 'G' + rn + ':K' + rn);
    count += d.faculty.length;
  });
  const gt = rows.length + 1, gr = { c: blank().map(() => [null, S.numB]) };
  put(gr, 0, 'Grand Total', S.txtB); put(gr, 1, count, S.numB);
  [4, 5, 6, 7, 8, 9, 10, 16, 17, 19].forEach(j => put(gr, j, { f: totRows.map(t => colL(j) + t).join('+') || '0', v: grand[j] }, isDec(j) ? S.decB : S.numB));
  rows.push(gr);
  const ps = grand[4] + grand[5], qs = grand[6] + grand[7] + grand[8] + grand[9] + grand[10];
  const gp = { c: blank().map(() => [null, S.pctB]) }; put(gp, 0, 'Overall %', S.txtB);
  [4, 5].forEach(j => put(gp, j, { f: pctF(colL(j), gt, 'E', 'F'), v: div(grand[j], ps) }, S.pctB));
  [6, 7, 8, 9, 10].forEach(j => put(gp, j, { f: pctF(colL(j), gt, 'G', 'K'), v: div(grand[j], qs) }, S.pctB));
  rows.push(gp);
  const go = { ht: 40, c: blank().map(() => [null, S.txtB]) }; put(go, 0, 'Overall Ratio', S.txtB);
  put(go, 4, metTxt('P', div(grand[4], ps), 75), S.txtB); put(go, 6, metTxt('SA', div(grand[6], qs), 40) + '\n' + metTxt('SA+PA+SP+IP', div(qs - grand[10], qs), 90) + '\nA <= 10% ' + (div(grand[10], qs) <= 0.1 + 1e-9 ? 'met' : 'not met') + ' (' + pctTxt(div(grand[10], qs)) + ')', S.txtB);
  rows.push(go); const gn = rows.length; merges.push('E' + gn + ':F' + gn, 'G' + gn + ':K' + gn);
  return sheetXml(rows, [24, 18, 10, 7, 11, 11, 9, 9, 9, 9, 9, 48, 8, 9, 9, 9, 10, 10, 9, 7, 24], merges, true);
}

function buildTable32b(m, S) {
  const rows = [], merges = [];
  rows.push({ ht: 20, c: [['Table 3-2b: Faculty Sufficiency and Qualifications Ratios by Discipline, January–December ' + m.year, S.title]] });
  rows.push({ c: [['Ratios from Table 3-1. Sufficiency uses teaching hours (SCH); qualifications use percent of time devoted to mission.', S.sub]] });
  rows.push({ ht: 30, c: [['Discipline', S.hdr], ['Faculty Sufficiency (SCH)', S.hdr], [null, S.hdr], ['Faculty Qualifications with Respect to Percent of Time Devoted to Mission', S.hdr], [null, S.hdr], [null, S.hdr], [null, S.hdr], [null, S.hdr]] });
  merges.push('A3:A4', 'B3:C3', 'D3:H3');
  rows.push({ ht: 40, c: [null, 'Participating Faculty (P)', 'Supporting Faculty (S)', 'Scholarly Academic (SA)', 'Practice Academic (PA)', 'Scholarly Practitioner (SP)', 'Instructional Practitioner (IP)', 'Additional Faculty (A)'].map(v => [v, S.hdr]) });
  const K = ['SA', 'PA', 'SP', 'IP', 'A'], g = { P: 0, S: 0, SA: 0, PA: 0, SP: 0, IP: 0, A: 0 };
  m.disciplines.forEach(d => {
    const s = { P: 0, S: 0, SA: 0, PA: 0, SP: 0, IP: 0, A: 0 };
    d.faculty.forEach(f => { s.P += f.P || 0; s.S += f.S || 0; s[f.cls] += f.pct; });
    Object.keys(s).forEach(k => { g[k] += s[k]; });
    const ps = s.P + s.S, qs = K.reduce((t, k) => t + s[k], 0);
    rows.push({ c: [[d.name, S.txt], [div(s.P, ps), S.pct], [div(s.S, ps), S.pct]].concat(K.map(k => [div(s[k], qs), S.pct])) });
  });
  const ps = g.P + g.S, qs = K.reduce((t, k) => t + g[k], 0);
  rows.push({ c: [['Total', S.txtB], [div(g.P, ps), S.pctB], [div(g.S, ps), S.pctB]].concat(K.map(k => [div(g[k], qs), S.pctB])) });
  rows.push({ ht: 26, c: [['Minimum Standard', S.txtB], ['P ≥ 75% overall; ≥ 60% per discipline', S.txtB], [null, S.txtB], ['SA ≥ 40%; SA+PA+SP+IP ≥ 90%', S.txtB], [null, S.txtB], [null, S.txtB], [null, S.txtB], ['A ≤ 10%', S.txtB]] });
  const n = rows.length; merges.push('B' + n + ':C' + n, 'D' + n + ':G' + n);
  const ok = b => (b ? 'Met' : 'Not met');
  rows.push({ c: [['Status', S.txtB], [ok(div(g.P, ps) >= 0.75 - 1e-9), S.txtB], [null, S.txtB], [ok(div(g.SA, qs) >= 0.4 - 1e-9 && div(qs - g.A, qs) >= 0.9 - 1e-9), S.txtB], [null, S.txtB], [null, S.txtB], [null, S.txtB], [ok(div(g.A, qs) <= 0.1 + 1e-9), S.txtB]] });
  const n2 = rows.length; merges.push('B' + n2 + ':C' + n2, 'D' + n2 + ':G' + n2);
  return sheetXml(rows, [22, 13, 13, 13, 13, 13, 13, 13], merges, true);
}

function buildTable81A(m, S) {
  const rows = [], merges = [], N = 18, blank = () => new Array(N).fill(null);
  rows.push({ ht: 20, c: [['Table 8-1A: Intellectual Contributions by Individual Faculty — Data Date Range: Jan ' + (m.year - 5) + ' – Dec ' + m.year, S.title]] });
  rows.push({ c: [['Credit: 1.0 sole CamEd author; 0.5 shared with a CamEd colleague. P/S and FTE from Table 3-1.', S.sub]] });
  const g = { ht: 30, c: blank().map(() => [null, S.hdr]) };
  g.c[0] = ['Discipline', S.hdr]; g.c[1] = ['Faculty Name', S.hdr]; g.c[2] = ['Portfolio of Intellectual Contributions', S.hdr]; g.c[6] = ['Types of Intellectual Contributions', S.hdr]; g.c[10] = ['% of Faculty Producing Intellectual Contributions', S.hdr];
  g.c[15] = ['Notes', S.hdr]; g.c[16] = ['IC summary (by policy code)', S.hdr]; g.c[17] = ['Professional engagement summary (P codes)', S.hdr];
  rows.push(g); merges.push('A3:A4', 'B3:B4', 'C3:F3', 'G3:J3', 'K3:O3', 'P3:P4', 'Q3:Q4', 'R3:R4');
  rows.push({ ht: 66, c: [null, null, 'Basic or Discovery Scholarship', 'Applied or Integration/ Application Scholarship', 'Teaching and Learning Scholarship', 'Total', 'Peer-reviewed journal articles', 'Additional peer- or editorial-reviewed intellectual contributions', 'All other intellectual contributions', 'Total', 'P or S', 'ICs (Yes/No)', 'FTE', 'Percentage of participating faculty producing ICs', 'Percentage of total FTE faculty producing ICs', null, null, null].map(v => [v, S.hdr]) });
  const subs = [], G = { B: 0, C: 0, D: 0, F: 0, G: 0, H: 0, n: 0, P: 0, S: 0, yes: 0, pYes: 0, fte: 0, fteYes: 0 };
  m.disciplines.forEach(d => {
    const first = rows.length + 1, s = { B: 0, C: 0, D: 0, F: 0, G: 0, H: 0, P: 0, S: 0, yes: 0, pYes: 0, fte: 0, fteYes: 0 };
    d.faculty.forEach((f, i) => {
      const b = f.icBy || {}, r = rows.length + 1, has = (f.nRes || 0) > 0, role = f.P > 0 ? 'P' : 'S';
      const v = k => R2(b[k] || 0);
      rows.push({ ht: Math.max(lineHt(f.icText, 48), lineHt(f.peText, 48)), c: [[i ? '' : d.name, S.txt], [f.name, S.txt], [v('B'), S.num], [v('C'), S.num], [v('D'), S.num], [{ f: 'SUM(C' + r + ':E' + r + ')', v: v('B') + v('C') + v('D') }, S.num],
        [v('F'), S.num], [v('G'), S.num], [v('H'), S.num], [{ f: 'SUM(G' + r + ':I' + r + ')', v: v('F') + v('G') + v('H') }, S.num], [role, S.cen], [has ? 'Yes' : 'No', S.cen], [R2(f.fte || 0), S.dec], [null, S.cen], [null, S.cen],
        [f.note || '', S.txt], [f.icText || '', S.txt], [f.peText || '', S.txt]] });
      ['B', 'C', 'D', 'F', 'G', 'H'].forEach(k => { s[k] += b[k] || 0; });
      s[role]++; if (has) { s.yes++; if (role === 'P') s.pYes++; s.fteYes += f.fte || 0; } s.fte += f.fte || 0;
    });
    const last = rows.length, t = last + 1, rg = c => c + first + ':' + c + last;
    const sm = (c, val, st) => [{ f: 'SUM(' + rg(c) + ')', v: val }, st || S.numB];
    rows.push({ c: [['Subtotal ' + d.name, S.txtB], [d.faculty.length, S.numB], sm('C', s.B), sm('D', s.C), sm('E', s.D), sm('F', s.B + s.C + s.D), sm('G', s.F), sm('H', s.G), sm('I', s.H), sm('J', s.F + s.G + s.H),
      ['P=' + s.P + ', S=' + s.S, S.numB], ['Yes=' + s.yes, S.numB], sm('M', R2(s.fte), S.decB),
      [{ f: 'IFERROR(COUNTIFS(' + rg('K') + ',"P",' + rg('L') + ',"Yes")/COUNTIF(' + rg('K') + ',"P"),0)', v: div(s.pYes, s.P) }, S.pctB],
      [{ f: 'IFERROR(SUMIFS(' + rg('M') + ',' + rg('L') + ',"Yes")/SUM(' + rg('M') + '),0)', v: div(s.fteYes, s.fte) }, S.pctB], [null, S.txtB], [null, S.txtB], [null, S.txtB]] });
    subs.push(t); Object.keys(s).forEach(k => { G[k] += s[k]; }); G.n += d.faculty.length;
  });
  const gs = (c, val, st) => [{ f: subs.map(t => c + t).join('+') || '0', v: val }, st || S.numB];
  rows.push({ c: [['Grand Total', S.txtB], [G.n, S.numB], gs('C', G.B), gs('D', G.C), gs('E', G.D), gs('F', G.B + G.C + G.D), gs('G', G.F), gs('H', G.G), gs('I', G.H), gs('J', G.F + G.G + G.H),
    ['P=' + G.P + ', S=' + G.S, S.numB], ['Yes=' + G.yes, S.numB], gs('M', R2(G.fte), S.decB), [div(G.pYes, G.P), S.pctB], [div(G.fteYes, G.fte), S.pctB], [null, S.txtB], [null, S.txtB], [null, S.txtB]] });
  return sheetXml(rows, [16, 24, 10, 12, 10, 8, 10, 13, 11, 8, 7, 8, 7, 12, 12, 26, 40, 40], merges, true);
}

export async function buildWorkbook(JSZip, templateBuf, model) {
  const zip = await JSZip.loadAsync(templateBuf);
  const rd = p => zip.file(p).async('string');
  zip.file('xl/worksheets/sheet1.xml', buildTable31(await rd('xl/worksheets/sheet1.xml'), model));
  zip.file('xl/worksheets/sheet2.xml', buildTable32(await rd('xl/worksheets/sheet2.xml'), model));
  zip.file('xl/worksheets/sheet3.xml', buildTable81(await rd('xl/worksheets/sheet3.xml'), model));
  const { st, S } = addStyles(await rd('xl/styles.xml'));
  zip.file('xl/styles.xml', st);
  const extra = [['Table 3-1A', 'sheet101.xml', buildTable31A(model, S)], ['Table 3-2b', 'sheet102.xml', buildTable32b(model, S)], ['Table 8-1A', 'sheet103.xml', buildTable81A(model, S)]];
  extra.forEach(([, f, xml]) => zip.file('xl/worksheets/' + f, xml));
  zip.remove('xl/calcChain.xml');
  let ct = (await rd('[Content_Types].xml')).replace(/<Override PartName="\/xl\/calcChain.xml"[^>]*\/>/, '');
  ct = ct.replace('</Types>', () => extra.map(([, f]) => '<Override PartName="/xl/worksheets/' + f + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('') + '</Types>');
  zip.file('[Content_Types].xml', ct);
  let rels = (await rd('xl/_rels/workbook.xml.rels')).replace(/<Relationship [^>]*calcChain[^>]*\/>/, '');
  rels = rels.replace('</Relationships>', () => extra.map(([, f], i) => '<Relationship Id="rIdX' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/' + f + '"/>').join('') + '</Relationships>');
  zip.file('xl/_rels/workbook.xml.rels', rels);
  let wb = (await rd('xl/workbook.xml')).replace(/activeTab="\d+"/, 'activeTab="0"').replace(/<calcPr ([^/]*)\/>/, (mm, a) => '<calcPr ' + a.replace(/ ?fullCalcOnLoad="[^"]*"/, '') + ' fullCalcOnLoad="1"/>');
  const old = wb.match(/<sheet [^>]*\/>/g), maxId = Math.max(...old.map(s => +s.match(/sheetId="(\d+)"/)[1]));
  const nw = extra.map(([n], i) => '<sheet name="' + n + '" sheetId="' + (maxId + i + 1) + '" r:id="rIdX' + (i + 1) + '"/>');
  const order = [old[0], nw[0], old[1], nw[1], old[2], nw[2]], pos = [0, 2, 4];
  wb = wb.replace(/<sheets>[\s\S]*?<\/sheets>/, () => '<sheets>' + order.join('') + '</sheets>').replace(/localSheetId="(\d+)"/g, (mm, i) => 'localSheetId="' + (pos[+i] ?? i) + '"');
  zip.file('xl/workbook.xml', wb);
  if (zip.file('docProps/app.xml')) zip.file('docProps/app.xml', (await rd('docProps/app.xml')).replace(/<HeadingPairs>[\s\S]*?<\/HeadingPairs>/, '').replace(/<TitlesOfParts>[\s\S]*?<\/TitlesOfParts>/, ''));
  return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
