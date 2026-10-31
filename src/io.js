import ExcelJS from 'exceljs';
import SSF from 'ssf';
import { cellText } from './core.js';
import {LIMITS,VERSION,validateText} from './limits.js';
import {precheckXlsx} from './xlsx-precheck.js';

export {LIMITS};

// Strict CSV parsing keeps the starting physical line of each quoted record.
export function parseCSV(text, delimiter) {
  text = text.replace(/^\uFEFF/, '');
  if (!delimiter) {
    const counts = { ',': 0, ';': 0, '\t': 0 };
    let quoted = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (c === '"') { if (quoted && text[i + 1] === '"') i++; else quoted = !quoted; }
      if (!quoted && (c === '\n' || c === '\r')) break;
      if (!quoted && Object.hasOwn(counts, c)) counts[c]++;
    }
    delimiter = Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0];
  }
  const rows = [];
  let values = [], value = '', mode = 'start', line = 1, startLine = 1;
  let totalCells=0;
  function emitCell() {
    validateText(value);
    if(values.length>=LIMITS.columns)throw new Error('CSV 超过 200 列，请先裁减');
    if(++totalCells>LIMITS.cells)throw new Error('CSV 超过 150,000 个单元格（含标题），请先拆分');
    values.push({ raw: value, text: value, kind: 'text' }); value = ''; mode = 'start';
  }
  function emitRow() { emitCell(); rows.push({ rowNumber: startLine, cells: values }); values = []; startLine = line + 1; }
  for (let i = 0; i < text.length; i++) {
    const c = text[i], newline = c === '\n' || c === '\r';
    if (mode === 'quoted') {
      if (c === '"') {
        if (text[i + 1] === '"') { value += '"'; i++; } else mode = 'closed';
      } else if (newline) {
        if (c === '\r' && text[i + 1] === '\n') i++;
        value += '\n'; line++;
      } else value += c;
      continue;
    }
    if (c === delimiter) { emitCell(); continue; }
    if (newline) {
      if (c === '\r' && text[i + 1] === '\n') i++;
      emitRow(); line++;
      if (rows.length > LIMITS.rows + LIMITS.headerRecords) throw new Error('CSV 超过 25,100 条记录（含标题/说明），请先拆分文件');
      continue;
    }
    if (mode === 'closed') throw new Error(`CSV 第 ${line} 行引号后有非法字符，无法可靠解析`);
    if (c === '"') {
      if (mode !== 'start') throw new Error(`CSV 第 ${line} 行引号位置错误，无法可靠解析`);
      mode = 'quoted';
    } else { value += c; mode = 'plain'; }
  }
  if (mode === 'quoted') throw new Error(`CSV 第 ${startLine} 行起的引号未闭合，导入已停止`);
  if (value !== '' || values.length || mode === 'closed') { emitCell(); rows.push({ rowNumber: startLine, cells: values }); }
  if (!rows.length) throw new Error('文件为空');
  return rows;
}

function excelCell(cell,unsafe) {
  const v = cell.value;
  if(unsafe?.has(cell.address)){const raw=unsafe.get(cell.address);return {raw,text:raw,unsafeNumeric:true,kind:'number'};}
  if (v == null) return { raw: '', text: '', kind: 'blank' };
  if (typeof v === 'object' && ('formula' in v || 'sharedFormula' in v)) return { raw: v.result ?? '', text: v.result == null ? '' : String(v.result), formula: cell.formula || v.formula || v.sharedFormula, kind: 'formula' };
  if (typeof v === 'object' && 'error' in v) return { raw: v.error, text: v.error, error: v.error, kind: 'error' };
  if (v instanceof Date) return { raw: v.toISOString(), text: cell.text, kind: 'date' };
  if (typeof v === 'object') {
    const s = v.richText ? v.richText.map(x => x.text).join('') : String(v.text ?? cell.text ?? '');
    validateText(s);return { raw: s, text: s, kind: 'text' };
  }
  let text = String(v);
  if (typeof v === 'number' && cell.numFmt && cell.numFmt !== 'General') {
    try { text = SSF.format(cell.numFmt, v); } catch { /* keep stored numeric value */ }
  }
  if(typeof v==='string')validateText(v);return { raw: v, text, kind: typeof v };
}

export async function importBytes(bytes, fileName, encoding = 'utf-8') {
  const buffer = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (buffer.byteLength > LIMITS.bytes) throw new Error('文件超过 20 MB，请先拆分后再导入');
  if (/\.csv$/i.test(fileName)) {
    let text;
    try { text = new TextDecoder(encoding, { fatal: true }).decode(buffer); }
    catch { throw new Error('CSV 编码解码失败，请尝试 GB18030 或另存为 UTF-8 CSV'); }
    return { fileName, sheets: [{ name: 'CSV', rows: parseCSV(text), merged: [] }], format: 'csv' };
  }
  if (!/\.xlsx$/i.test(fileName)) throw new Error('仅支持 .xlsx 和 .csv；不接受宏工作簿或旧版 .xls');
  if (buffer[0] !== 0x50 || buffer[1] !== 0x4b) throw new Error('XLSX 文件格式无效，未参与核对');
  const unsafeCells=await precheckXlsx(buffer);
  const workbook = new ExcelJS.Workbook();
  try { await workbook.xlsx.load(buffer, { ignoreNodes: ['drawing', 'picture', 'extLst'] }); }
  catch { throw new Error('XLSX 解析失败：文件可能损坏、加密或不受支持；未参与核对'); }
  let totalCells=0;
  const sheets = workbook.worksheets.map(ws => {
    const height=ws.rowCount,width=ws.columnCount;
    if (height > LIMITS.rows + LIMITS.headerRecords || width > LIMITS.columns) throw new Error(`工作表「${ws.name}」超出首版范围（25,000 行 / 200 列）`);
    totalCells+=height*width;if(totalCells>LIMITS.cells)throw new Error('工作簿读取矩形区域超过 150,000 个单元格，请裁减空格式列或拆分');
    const rows = [];
    for (let n = 1; n <= height; n++) {
      const row = ws.getRow(n), cells = [];
      for (let c = 1; c <= width; c++) cells.push(excelCell(row.getCell(c),unsafeCells.get(ws.name)));
      rows.push({ rowNumber: n, cells });
    }
    return { name: ws.name, rows, merged: ws.model.merges || [] };
  });
  if (!sheets.length) throw new Error('工作簿没有可用工作表');
  return { fileName, sheets, format: 'xlsx' };
}

export function selectTable(book, sheetIndex, headerRow) {
  const sheet = book.sheets[sheetIndex];
  if (!sheet) throw new Error('工作表不存在');
  if (!Number.isInteger(headerRow) || headerRow < 1 || headerRow > sheet.rows.length) throw new Error('标题行超出工作表范围');
  if (sheet.merged.some(range => Number(range.match(/\d+$/)?.[0]) >= sheet.rows[headerRow - 1].rowNumber)) throw new Error('标题或数据区域有合并单元格，请先在副本中整理为单行标题、取消合并并填齐数据');
  const header = sheet.rows[headerRow - 1];
  let width = header.cells.length;
  while (width > 0 && cellText(header.cells[width - 1]).trim() === '') width--;
  if (!width) throw new Error('所选标题行为空，请重新选择');
  if (width > LIMITS.columns) throw new Error('超过 200 列，请先裁减');
  const columns = Array.from({ length: width }, (_, i) => ({ index: i, label: cellText(header.cells[i]) || `未命名列 ${i + 1}` }));
  const rows = sheet.rows.slice(headerRow).filter(r => r.cells.some(c => c.formula || c.error || cellText(c).trim() !== ''));
  if (rows.length > LIMITS.rows) throw new Error('数据超过 25,000 行，请先拆分');
  if (!rows.length) throw new Error('标题行之后没有明细记录，请重新选择');
  if (rows.some(r => r.cells.slice(width).some(c => c.formula || c.error || cellText(c).trim() !== ''))) throw new Error('标题行右侧存在无标题数据，请补齐标题或选择正确标题行');
  const warnings = [];
  if (rows.some(r => r.cells.some(c => c.formula))) warnings.push('检测到公式：如映射为主键、状态或金额，将隔离该记录；请在副本中转为值。');
  if (rows.some(r => r.cells.some(c => typeof c.raw === 'number'))) warnings.push('数字编号仅能保留文件中已有的文本或补零格式；源文件已丢失的前导零无法恢复。');
  return { fileName: book.fileName, sheetName: sheet.name, headerRow: header.rowNumber, columns, rows, warnings };
}

export const safeText = value => {
  const s = (value == null ? '' : String(value)).replace(/_(?=x[0-9A-Fa-f]{4}_)/g,'_x005F_');
  return /^[\s\uFEFF]*[=+@-]/u.test(s) || /^[\t\r\n]/.test(s) ? "'" + s : s;
};

function exportValue(value){
  if(typeof value==='number')return value;
  const s=value==null?'':String(value);validateText(s);const encoded=safeText(s);
  if(encoded.length>32767)throw new Error('安全转义后的文本超过 XLSX 单元格上限，请精简源文本；未截断原值');
  return encoded;
}

export async function reportBytes(result) {
  if(result.differences.length+result.issues.length>LIMITS.resultRows)throw new Error('异常明细超过 25,000 条，无法生成完整报告，请拆分');
  const workbook = new ExcelJS.Workbook();
  workbook.creator = '台账对对看（本机核对）';
  workbook.created = new Date(result.createdAt);
  const overview = workbook.addWorksheet('概览'), s = result.summary;
  const rows = [
    ['台账对对看 · 独立核对报告', VERSION],
    ['结论', result.allConsistent ? '所选字段全部一致' : s.issueCount ? '存在待处理记录，不能确认全部一致' : '发现差异'],
    ['生成时间（UTC）', result.createdAt],
    ['A 文件', result.sources.left.fileName], ['A 工作表', result.sources.left.sheetName], ['A 标题原始行', result.sources.left.headerRow],
    ['B 文件', result.sources.right.fileName], ['B 工作表', result.sources.right.sheetName], ['B 标题原始行', result.sources.right.headerRow],
    ['A 明细记录数', s.left], ['B 明细记录数', s.right], ['可核对配对数', s.matched], ['配对一致数', s.agreed], ['有差异配对数', s.changed],
    ['仅 A 有记录数', s.missingRight], ['仅 B 有记录数', s.missingLeft], ['差异明细条数', s.differenceCount], ['重复/无效明细条数', s.issueCount], ['待处理源记录数', s.invalidRows], ['同侧重复源记录数', s.duplicateRows],
    ['金额容差（元）', result.config.tolerance], ['金额比较', '差额=A−B；绝对差额≤容差视为相同；空白不等于零'],
    ['主键规则', '逐分量精确文本匹配；不忽略大小写或首尾空格；不做模糊匹配'],
    ['隔离规则', '重复键和无效字段的键两侧均不强行匹配；空键/不完整主键无法建立关联'],
    ['原始行号', 'XLSX 为工作表行号；CSV 为记录起始物理行号（标题选择使用记录序号）'],
    ['安全说明', '报告为新文件，不修改源表；危险文本前置单引号，所有文本以字符串写入；无公式、宏或外部链接'],
    ['范围限制', '未映射字段不核对；不判断哪张表正确，不重算公式；仅使用源文件保存值'],
  ];
  result.config.keys.forEach((k, i) => rows.push([`主键 ${i + 1}`, `${k.label}：A 第${k.left + 1}列 ↔ B 第${k.right + 1}列`]));
  result.config.fields.forEach((f, i) => rows.push([`核对字段 ${i + 1}`, `${f.label}：A 第${f.left + 1}列 ↔ B 第${f.right + 1}列；${f.type === 'amount' ? `金额，A=${f.leftUnit}，B=${f.rightUnit}` : '状态/文本'}`]));
  rows.forEach(r => overview.addRow(r.map(exportValue)));
  overview.columns = [{ width: 30 }, { width: 110 }];
  const keys = ['type','key','field','leftFile','leftSheet','leftRow','leftValue','rightFile','rightSheet','rightRow','rightValue','delta','note'];
  const headers = ['异常类型','主键','字段','A 文件名','A 工作表','A 原始行号','A 原值','B 文件名','B 工作表','B 原始行号','B 原值','金额差（元，A−B）','说明'];
  for (const [name, data] of [['差异明细', result.differences], ['重复与无效记录', result.issues]]) {
    const ws = workbook.addWorksheet(name);
    ws.addRow(headers);
    data.forEach(item => ws.addRow(keys.map(k => exportValue(item[k]))));
    ws.columns.forEach((col, i) => { col.width = [22, 42, 20, 28, 20, 14, 30, 28, 20, 14, 30, 24, 85][i]; });
    ws.autoFilter = { from: 'A1', to: `M${Math.max(1, ws.rowCount)}` };
    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }
  for (const ws of workbook.worksheets) {
    ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF163E39' } };
    ws.eachRow(row => { row.alignment = { vertical: 'top', wrapText: true }; });
  }
  return workbook.xlsx.writeBuffer();
}
