import Decimal from 'decimal.js';
import {LIMITS} from './limits.js';

Decimal.set({ precision: 80 });
export const UNITS = { '元': '1', '千元': '1000', '万元': '10000', '亿元': '100000000' };
const blank = v => v == null || String(v).trim() === '';
export const cellText = c => c?.text ?? (c?.raw == null ? '' : String(c.raw));
export const cellRaw = c => c?.raw == null ? '' : String(c.raw);

export function parseAmount(cell, unit = '元') {
  if (!Object.hasOwn(UNITS, unit)) return { error: '未选择合法金额单位' };
  if (cell?.formula) return { error: '关键字段含公式，请在原表副本中转为值后重新导入' };
  if (cell?.error) return { error: '单元格含 Excel 错误值' };
  if(cell?.unsafeNumeric)return {error:'XLSX 原数值存在超过 15 位有效数字或读取舍入风险，请在来源中核实并转为文本数值'};
  if (blank(cell?.raw)) return { blank: true };
  if (cell?.kind === 'date' || typeof cell?.raw === 'boolean') return { error: '日期或布尔值不能作为金额' };
  let s = String(cell.raw).trim();
  if (typeof cell.raw === 'number') {
    if (!Number.isFinite(cell.raw)) return { error: '金额不是有限数字' };
    s = new Decimal(s).toFixed();
  }
  if (!/^[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$/.test(s)) return { error: '非法金额：仅接受数字、小数点、正负号及规范千位逗号' };
  s = s.replaceAll(',', '');
  if (s.replace(/[^0-9]/g, '').length > 40 || (s.split('.')[1]?.length ?? 0) > 12) return { error: '金额超过首版支持精度（40 位数字、12 位小数）' };
  return { value: new Decimal(s).times(UNITS[unit]) };
}

function checkConfig(left, right, config) {
  if (!left?.rows || !right?.rows) throw new Error('请先导入并确认两份台账');
  if (!config.keys?.length) throw new Error('至少选择一个主键字段');
  if (!config.fields?.length) throw new Error('至少选择一个核对字段');
  if(config.keys.length>LIMITS.keys||config.fields.length>LIMITS.fields)throw new Error('首版最多 8 个主键分量、20 个核对字段，请精简映射');
  const used = { left: new Set(), right: new Set() };
  for (const mapping of [...config.keys, ...config.fields]) {
    for (const side of ['left', 'right']) {
      const index = mapping[side], data = side === 'left' ? left : right;
      if (!Number.isInteger(index) || index < 0 || index >= data.columns.length) throw new Error('字段映射不完整，请逐项选择两边列');
      if (used[side].has(index)) throw new Error(`${side === 'left' ? 'A' : 'B'} 表同一列被重复映射，请检查`);
      used[side].add(index);
    }
  }
  for (const f of config.fields) {
    if (!['text', 'amount'].includes(f.type)) throw new Error('未知核对字段类型');
    if (f.type === 'amount' && (!Object.hasOwn(UNITS, f.leftUnit) || !Object.hasOwn(UNITS, f.rightUnit))) throw new Error('金额单位未选择');
  }
  const tol = parseAmount({ raw: config.tolerance ?? '0.01' });
  if (!tol.value || tol.value.isNegative()) throw new Error('容差须为大于或等于零的合法金额（单位：元）');
  return tol.value;
}

export function reconcile(left, right, config) {
  const tolerance = checkConfig(left, right, config);
  const issues = [], differences = [];
  const sources = { left, right }, groups = { left: new Map(), right: new Map() };
  const blocked = new Set(), invalidRecords = new Set(), duplicateRecords = new Set();
  function trace(type, key, field, l, r, leftValue = '', rightValue = '', delta = '', note = '') {
    if(issues.length+differences.length>=LIMITS.resultRows)throw new Error('异常明细超过 25,000 条，请拆分台账或减少核对字段；未输出部分结果');
    if([key,field,leftValue,rightValue,delta,note].some(v=>String(v).length>32766))throw new Error('异常原值组合超过 XLSX 单元格长度，请精简字段或整理源表；未截断原值');
    return { type, key, field,
      leftFile: left.fileName, leftSheet: left.sheetName, leftRow: l?.rowNumber ?? '', leftValue,
      rightFile: right.fileName, rightSheet: right.sheetName, rightRow: r?.rowNumber ?? '', rightValue, delta, note };
  }
  function addIssue(side, row, type, key, field, value, note) {
    issues.push(trace(type, key, field, side === 'left' ? row : null, side === 'right' ? row : null,
      side === 'left' ? value : '', side === 'right' ? value : '', '', note));
    invalidRecords.add(`${side}:${row.rowNumber}`);
  }
  for (const side of ['left', 'right']) {
    for (const row of sources[side].rows) {
      const keyCells = config.keys.map(k => row.cells[k[side]]);
      const keyValues = keyCells.map(cellText);
      const displayKey = config.keys.map((k, i) => `${k.label || `键${i + 1}`}=${keyValues[i]}`).join(' · ');
      const entry = { row, displayKey, amounts: new Map(), valid: true };
      // Audit mapped fields even when the same row has no usable key.
      for (let i = 0; i < config.fields.length; i++) {
        const f = config.fields[i], cell = row.cells[f[side]], field = f.label || `字段${i + 1}`;
        let error = cell?.formula ? '关键字段含公式，请在原表副本中转为值后重新导入' : cell?.error ? '单元格含 Excel 错误值' : null;
        if (f.type === 'amount') {
          const amount = parseAmount(cell, f[side + 'Unit']);
          entry.amounts.set(i, amount);
          error ||= amount.error;
        }
        if (error) {
          entry.valid = false;
          addIssue(side, row, cell?.formula ? '关键字段含公式' : f.type === 'amount' ? '非法金额' : '无效字段', displayKey, field,
            cell?.formula ? `=${cell.formula}${cell.raw == null || cell.raw === '' ? ' [无缓存]' : ` [缓存：${cellRaw(cell)}]`}` : cellRaw(cell), error);
        }
      }
      if (keyCells.some(c => c?.formula)) {
        addIssue(side, row, '关键字段含公式', displayKey, '主键', keyCells.map(c => c?.formula ? `=${c.formula}${c.raw === '' ? ' [无缓存]' : ''}` : cellRaw(c)).join(' | '), '主键含公式（包括有缓存公式），请在原表副本转为值；未参与匹配');
        continue;
      }
      const emptyCount = keyValues.filter(blank).length;
      if (emptyCount) {
        addIssue(side, row, emptyCount === keyValues.length ? '空键' : '复合键不完整', displayKey, '主键', JSON.stringify(keyValues), '未参与匹配；请补齐全部主键分量');
        continue;
      }
      const keyError = keyCells.find(c => c?.formula || c?.error || c?.unsafeNumeric || cellText(c).length>LIMITS.keyChars || (typeof c?.raw === 'number' && (!Number.isSafeInteger(c.raw) || c.raw < 0)));
      if (keyError) {
        addIssue(side, row, keyError.formula ? '关键字段含公式' : '无效主键', displayKey, '主键', JSON.stringify(keyValues),
          keyError.formula ? '主键含公式（包括有缓存公式），请在原表副本转为值；未参与匹配' : '主键含错误值、精度风险、负数、非整数、超安全整数或分量超过 256 字符，请核实文本编号');
        continue;
      }
      const id = JSON.stringify(keyValues);
      if (!groups[side].has(id)) groups[side].set(id, []);
      groups[side].get(id).push(entry);
      if (!entry.valid) blocked.add(id);
    }
    for (const [id, entries] of groups[side]) {
      if (entries.length < 2) continue;
      blocked.add(id);
      const rowPreview=entries.slice(0,20).map(x=>x.row.rowNumber).join('、');
      const duplicateNote=`同侧共 ${entries.length} 条；原始行${entries.length>20?'（前 20 条）':''}：${rowPreview}；该键两侧均不参与匹配。每条原始行在本表逐条保留`;
      for (const e of entries) {
        duplicateRecords.add(`${side}:${e.row.rowNumber}`);
        addIssue(side, e.row, '重复键', e.displayKey, '主键', id, duplicateNote);
      }
    }
  }
  for (const side of ['left', 'right']) {
    for (const [id, entries] of groups[side]) {
      if (!blocked.has(id) || entries.length !== 1 || !entries[0].valid) continue;
      const e = entries[0];
      addIssue(side, e.row, '对侧记录待处理', e.displayKey, '主键', id, '对侧存在重复键或无效核对字段；未强行匹配');
    }
  }
  let matched = 0, agreed = 0, changed = 0, missingLeft = 0, missingRight = 0;
  for (const id of new Set([...groups.left.keys(), ...groups.right.keys()])) {
    if (blocked.has(id)) continue;
    const a = groups.left.get(id)?.[0], b = groups.right.get(id)?.[0];
    if (!a || !b) {
      if (!a) missingLeft++; else missingRight++;
      differences.push(trace(a ? '仅 A 有' : '仅 B 有', (a || b).displayKey, '整条记录', a?.row, b?.row,
        a ? JSON.stringify(a.row.cells.map(cellRaw)) : '', b ? JSON.stringify(b.row.cells.map(cellRaw)) : '', '', '仅按主键精确匹配；请结合重复/无效记录检查可能关联项'));
      continue;
    }
    matched++;
    const before = differences.length;
    config.fields.forEach((f, i) => {
      const ac = a.row.cells[f.left], bc = b.row.cells[f.right], field = f.label || `字段${i + 1}`;
      if (f.type === 'text') {
        if (cellText(ac) !== cellText(bc)) differences.push(trace('状态/文本不同', a.displayKey, field, a.row, b.row, cellRaw(ac), cellRaw(bc), '', '精确比较文本，保留大小写和首尾空格'));
      } else {
        const av = a.amounts.get(i), bv = b.amounts.get(i);
        if (av.blank && bv.blank) return;
        if (av.blank || bv.blank) {
          differences.push(trace('金额空值不同', a.displayKey, field, a.row, b.row, cellRaw(ac), cellRaw(bc), '', `空白不等于零；A 单位 ${f.leftUnit}，B 单位 ${f.rightUnit}`));
          return;
        }
        const delta = av.value.minus(bv.value);
        if (delta.abs().gt(tolerance)) differences.push(trace('金额不同', a.displayKey, field, a.row, b.row, cellRaw(ac), cellRaw(bc), delta.toFixed(),
          `统一为元：A=${av.value.toFixed()}；B=${bv.value.toFixed()}；差额=A−B；容差=${tolerance.toFixed()} 元（等于容差视为相同）`));
      }
    });
    if (differences.length > before) changed++; else agreed++;
  }
  return {
    createdAt: new Date().toISOString(), config: JSON.parse(JSON.stringify(config)),
    sources: { left: { fileName: left.fileName, sheetName: left.sheetName, headerRow: left.headerRow, rowCount: left.rows.length }, right: { fileName: right.fileName, sheetName: right.sheetName, headerRow: right.headerRow, rowCount: right.rows.length } },
    summary: { left: left.rows.length, right: right.rows.length, matched, agreed, changed, missingLeft, missingRight, differenceCount: differences.length, issueCount: issues.length, invalidRows: invalidRecords.size, duplicateRows: duplicateRecords.size },
    allConsistent: differences.length === 0 && issues.length === 0 && matched > 0, differences, issues,
  };
}
