const d = require('docx');
const { Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle, HeadingLevel, ShadingType, PageBreak, VerticalAlign } = d;
const FONT = 'Times New Roman';
const IND = 709; // 1,25 см
// разметка: **жирный**, [[заполнитель]]
function runs(text, base = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|\[\[[^\]]+\]\])/g;
  let last = 0, m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), font: FONT, ...base }));
    const t = m[0];
    if (t.startsWith('**')) out.push(new TextRun({ text: t.slice(2, -2), bold: true, font: FONT, ...base }));
    else out.push(new TextRun({ text: '[' + t.slice(2, -2) + ']', highlight: 'yellow', font: FONT, ...base }));
    last = m.index + t.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), font: FONT, ...base }));
  return out;
}
const p = (t, o = {}) => new Paragraph({ children: runs(t, o.run || {}), alignment: o.align || AlignmentType.JUSTIFIED, indent: o.noIndent ? undefined : { firstLine: IND }, spacing: { after: o.after ?? 60, line: 276 }, keepNext: o.keepNext });
const pc = (t, o = {}) => new Paragraph({ children: runs(t, o.run || {}), alignment: AlignmentType.CENTER, spacing: { after: o.after ?? 0, before: o.before ?? 0 }, keepNext: true });
const li = (t) => new Paragraph({ children: runs(t), alignment: AlignmentType.JUSTIFIED, numbering: { reference: 'dash', level: 0 }, spacing: { after: 40, line: 276 } });
const ol = (t, ref = 'num') => new Paragraph({ children: runs(t), alignment: AlignmentType.JUSTIFIED, numbering: { reference: ref, level: 0 }, spacing: { after: 40, line: 276 } });
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t, font: FONT, bold: true, size: 32 })], indent: { firstLine: IND }, spacing: { before: 240, after: 120 }, keepNext: true });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t, font: FONT, bold: true, size: 32 })], indent: { firstLine: IND }, spacing: { before: 160, after: 80 }, keepNext: true });
const note = (t) => new Paragraph({ children: runs(t, { size: 24 }), alignment: AlignmentType.JUSTIFIED, indent: { firstLine: IND }, spacing: { after: 80 } });
let tableNo = 0;
const border = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const borders = { top: border, bottom: border, left: border, right: border };
// мягкие переносы в словах, не помещающихся в ширину графы (упрощённые правила русского переноса)
function hyph(word) {
  const cls = word.split('').map(ch => /[аеёиоуыэюя]/i.test(ch) ? 'g' : /[йьъ]/i.test(ch) ? 'x' : 's').join('');
  const cut = new Set();
  const rules = [['xgg', 1], ['xgs', 1], ['xsg', 1], ['xss', 1], ['gssssg', 3], ['gsssg', 2], ['gssg', 2], ['sgsg', 2], ['sggs', 2], ['sggg', 2]];
  for (const [pat, pos] of rules) {
    let i = cls.indexOf(pat);
    while (i !== -1) { cut.add(i + pos); i = cls.indexOf(pat, i + 1); }
  }
  let out = '';
  for (let i = 0; i < word.length; i++) {
    if (cut.has(i) && i >= 2 && word.length - i >= 2) out += '\u00AD';
    out += word[i];
  }
  return out;
}
function fit(line, w, bold) {
  const maxChars = Math.floor((w - 160) / (bold ? 135 : 125));
  return line.replace(/[А-Яа-яЁё]{5,}/g, m => m.length > maxChars ? hyph(m) : m);
}
function cell(t, w, o = {}) {
  const lines = (Array.isArray(t) ? t : String(t).split('\n')).map(l => fit(String(l), w, !!o.header));
  return new TableCell({
    width: { size: w, type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.CENTER,
    shading: o.header ? { type: ShadingType.CLEAR, fill: 'F2F2F2', color: 'auto' } : undefined,
    margins: { top: 40, bottom: 40, left: 80, right: 80 }, columnSpan: o.span,
    children: lines.map(l => new Paragraph({ children: runs(l, { size: 24, bold: !!o.header }), alignment: o.header || o.center ? AlignmentType.CENTER : AlignmentType.LEFT })),
  });
}
// caption: "Таблица N – ..." слева над таблицей (ГОСТ Р 2.105-2019, 6.9)
function table(caption, widths, header, rows, o = {}) {
  const out = [];
  if (caption) { tableNo++; out.push(new Paragraph({ children: runs(`Таблица ${tableNo} – ${caption}`), spacing: { before: 120, after: 60 }, keepNext: true })); }
  const total = widths.reduce((a, b) => a + b, 0);
  const trs = [];
  const plain = rows.filter(r => !r.span);
  const auto = widths.map((_, i) => plain.length > 0 && plain.some(r => String(r[i] ?? '').trim() !== '') && plain.every(r => String(r[i] ?? '').replace(/\[\[|\]\]|\*\*/g, '').length <= 14));
  if (header) trs.push(new TableRow({ tableHeader: true, cantSplit: true, children: header.map((h, i) => cell(h, widths[i], { header: true, size: o.size })) }));
  rows.forEach(r => {
    if (r.span) trs.push(new TableRow({ cantSplit: true, children: [cell(r.span, total, { span: widths.length, size: o.size, header: true })] }));
    else trs.push(new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, widths[i], { size: o.size, center: auto[i] || (o.centerCols && o.centerCols.includes(i)) })) }));
  });
  out.push(new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: trs }));
  out.push(new Paragraph({ children: [], spacing: { after: 60 } }));
  return out;
}
function appendix(letter, status, title) {
  return [
    new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, alignment: AlignmentType.CENTER, children: [new TextRun({ text: `Приложение ${letter}`, font: FONT, bold: true, size: 32 })], spacing: { after: 0 } }),
    pc(`(${status})`, { run: { size: 28 } }),
    pc(title, { run: { bold: true, size: 32 }, after: 160, before: 60 }),
  ];
}
module.exports = { d, FONT, IND, runs, p, pc, li, ol, h1, h2, note, table, appendix, cell };
