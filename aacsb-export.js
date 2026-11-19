// Fills the AACSB tables workbook template (Table 3-1, 3-2, 8-1) from portal data.
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
      const v = { A: f.name, B: f.spec || '', C: f.deg, D: f.yr != null ? String(f.yr) : '', E: f.P || null, F: f.S || null, L: f.basis };
      v[CLS_COL[f.cls]] = f.pct;
      const styles = {}; COLS.forEach(c => { styles[c] = v[c] != null ? 28 : 29; });
      ['E', 'F'].forEach(c => { sum[c] += v[c] || 0; }); sum[CLS_COL[f.cls]] += f.pct;
      const lines = String(f.basis).split('\n').reduce((s, l) => s + Math.max(1, Math.ceil(l.length / 55)), 0);
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

export async function buildWorkbook(JSZip, templateBuf, model) {
  const zip = await JSZip.loadAsync(templateBuf);
  const rd = p => zip.file(p).async('string');
  zip.file('xl/worksheets/sheet1.xml', buildTable31(await rd('xl/worksheets/sheet1.xml'), model));
  zip.file('xl/worksheets/sheet2.xml', buildTable32(await rd('xl/worksheets/sheet2.xml'), model));
  zip.file('xl/worksheets/sheet3.xml', buildTable81(await rd('xl/worksheets/sheet3.xml'), model));
  zip.remove('xl/calcChain.xml');
  zip.file('[Content_Types].xml', (await rd('[Content_Types].xml')).replace(/<Override PartName="\/xl\/calcChain.xml"[^>]*\/>/, ''));
  zip.file('xl/_rels/workbook.xml.rels', (await rd('xl/_rels/workbook.xml.rels')).replace(/<Relationship [^>]*calcChain[^>]*\/>/, ''));
  zip.file('xl/workbook.xml', (await rd('xl/workbook.xml')).replace(/activeTab="\d+"/, 'activeTab="0"').replace(/<calcPr ([^/]*)\/>/, (mm, a) => `<calcPr ${a.replace(/ ?fullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"/>`));
  return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
