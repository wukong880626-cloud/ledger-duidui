import {test} from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {importBytes,selectTable,reportBytes,parseCSV} from '../src/io.js';
import {reconcile,parseAmount} from '../src/core.js';
import {LIMITS} from '../src/limits.js';
const enc=s=>new TextEncoder().encode(s);
const cfg={keys:[{label:'编号',left:0,right:0}],fields:[{label:'状态',type:'text',left:1,right:1},{label:'金额',type:'amount',left:2,right:2,leftUnit:'元',rightUnit:'元'}],tolerance:'0.01'};
const csv=async(s,name='A.csv')=>selectTable(await importBytes(enc(s),name),0,1);
async function numericWorkbook(value,coordinate='C2'){
 const wb=new ExcelJS.Workbook(),ws=wb.addWorksheet('明细');ws.addRow(['编号','状态','金额']);ws.addRow([coordinate==='A2'?1:'001','签约',1]);
 const zip=await JSZip.loadAsync(await wb.xlsx.writeBuffer()),path='xl/worksheets/sheet1.xml';
 zip.file(path,(await zip.file(path).async('string')).replace(new RegExp(`(<c r="${coordinate}"[^>]*><v>)1(</v>)`),`$1${value}$2`));
 return zip.generateAsync({type:'nodebuffer'});
}
test('XLSX 读取前识别超精度数值，保留 XML 原数值并隔离',async()=>{
 for(const value of ['1.00000000000000001','123456789012345.67','0.1234567890123456']){
  const a=selectTable(await importBytes(await numericWorkbook(value),'precision.xlsx'),0,1);
  assert.equal(a.rows[0].cells[2].raw,value);assert.equal(a.rows[0].cells[2].unsafeNumeric,true);
  const r=reconcile(a,await csv('编号,状态,金额\n001,签约,1','B.csv'),cfg);assert.equal(r.allConsistent,false);assert.ok(r.issues.some(x=>x.type==='非法金额'&&x.leftValue===value));
 }
});
test('不超过 15 有效位的正常数值不会被误拦截',async()=>{
 for(const value of ['0.1','1234567890123.45','1000000.01']){
  const a=selectTable(await importBytes(await numericWorkbook(value),'valid.xlsx'),0,1);assert.ok(!a.rows[0].cells[2].unsafeNumeric);
  assert.equal(reconcile(a,await csv(`编号,状态,金额\n001,签约,${value}`,'B.csv'),cfg).allConsistent,true);
 }
});
test('数值型超精度主键也被隔离，不使用舍入后的键',async()=>{
 const a=selectTable(await importBytes(await numericWorkbook('9007199254740993','A2'),'key.xlsx'),0,1),b=await csv('编号,状态,金额\n9007199254740992,签约,1','B.csv');
 const r=reconcile(a,b,cfg);assert.ok(r.issues.some(x=>x.type==='无效主键'));assert.equal(r.summary.matched,0);
});
test('高精度文本金额及单位换算保持十进制精度',async()=>{
 const a=await csv('编号,状态,金额\n001,签约,1234567890123456789012345678.123456789012'),b=await csv('编号,状态,金额\n001,签约,12345678901234567890.12345678','B.csv');
 const c=structuredClone(cfg);c.fields[1].rightUnit='亿元';c.tolerance='0';const r=reconcile(a,b,c);assert.equal(r.differences[0].delta,'0.123456789012');
 assert.equal(parseAmount({raw:'0.000001'},'万元').value.toFixed(),'0.01');
});
test('单位换算后的容差上下界、负数和空白不等零',async()=>{
 const a=await csv('编号,状态,金额\n001,签约,-0.01\n002,签约,0.010001\n003,签约,'),b=await csv('编号,状态,金额\n001,签约,0\n002,签约,0\n003,签约,0','B.csv');
 const r=reconcile(a,b,cfg);assert.equal(r.summary.agreed,1);assert.deepEqual(r.differences.map(x=>x.type),['金额不同','金额空值不同']);
});
test('5,000 条重复键说明保持有界，全部源行逐条导出',async()=>{
 const a=await csv('编号,状态,金额\n'+Array.from({length:5000},()=> '001,签约,1').join('\n')),b=await csv('编号,状态,金额\n001,签约,1','B.csv');
 const r=reconcile(a,b,cfg);assert.equal(r.summary.duplicateRows,5000);assert.equal(r.summary.issueCount,5001);assert.ok(Math.max(...r.issues.map(x=>x.note.length))<250);
 const wb=new ExcelJS.Workbook();await wb.xlsx.load(await reportBytes(r));const ws=wb.getWorksheet('重复与无效记录');assert.equal(ws.rowCount,5002);assert.equal(ws.getCell('F5001').value,5001);
});
test('CSV 控制字符及 DEL 明确拒绝，避免导出原值悄悄变化',async()=>{
 for(const control of ['\u0000','\u0001','\u000B','\u007F'])await assert.rejects(importBytes(enc(`编号,状态,金额\n001,含${control}文本,1`),'control.csv'),/控制字符/);
 assert.equal(parseCSV('编号,状态,金额\n001,"含\n换行\t制表",1')[1].cells[1].raw,'含\n换行\t制表');
});
test('XLSX 编码控制字符拒绝；原生转义文本导出往返不变',async()=>{
 const wb=new ExcelJS.Workbook(),ws=wb.addWorksheet('明细');ws.addRow(['编号','状态','金额']);ws.addRow(['001','bad_x0001_text',1]);
 await assert.rejects(importBytes(await wb.xlsx.writeBuffer(),'control.xlsx'),/控制字符/);
 const values=['_x0041_','_x0001_','_x005F_x0041_','_x0041__x0042_','_x00af_'];
 const a=await csv('编号,状态,金额\n'+values.map((v,i)=>`00${i},${v},1`).join('\n')),b=await csv('编号,状态,金额\n'+values.map((_,i)=>`00${i},普通,1`).join('\n'),'B.csv');
 const out=new ExcelJS.Workbook();await out.xlsx.load(await reportBytes(reconcile(a,b,cfg)));assert.deepEqual(out.getWorksheet('差异明细').getRows(2,values.length).map(r=>r.getCell('G').value),values);
});
test('文本长度上限与安全转义后长度上限不截断',async()=>{
 await assert.rejects(importBytes(enc('编号,文本\n001,'+'a'.repeat(LIMITS.cellChars+1)),'long.csv'),/32,766/);
 const s='_x0041_'.repeat(4000),a=await csv(`编号,状态,金额\n001,${s},1`),b=await csv('编号,状态,金额\n001,普通,1','B.csv');
 await assert.rejects(reportBytes(reconcile(a,b,cfg)),/安全转义/);
});
test('CSV 提前限制列数、总单元格及所选记录数',async()=>{
 await assert.rejects(importBytes(enc(Array.from({length:201},()=> '列').join(',')),'wide.csv'),/200 列/);
 await assert.rejects(importBytes(enc(Array.from({length:1501},()=>Array.from({length:100},()=> 'x').join(',')).join('\n')),'cells.csv'),/150,000/);
 const book=await importBytes(enc('编号,金额\n'+Array.from({length:25001},(_,i)=>`${i},1`).join('\n')),'rows.csv');assert.throws(()=>selectTable(book,0,1),/25,000/);
});
test('XLSX 稀疏大矩形和声明解压体积在解析/分配前拦截',async()=>{
 const wb=new ExcelJS.Workbook(),ws=wb.addWorksheet('宽表');ws.getCell('GR1').value='末列';ws.getCell('A751').value='末行';
 const bytes=await wb.xlsx.writeBuffer();await assert.rejects(importBytes(bytes,'wide.xlsx'),/150,000/);
 const altered=new Uint8Array(await numericWorkbook('1'));const view=new DataView(altered.buffer);let central=-1;for(let i=0;i<altered.length-4;i++)if(view.getUint32(i,true)===0x02014b50){central=i;break;}assert.ok(central>=0);view.setUint32(central+24,LIMITS.expandedBytes+1,true);
 await assert.rejects(importBytes(altered,'expanded.xlsx'),/解压体积/);
});
test('伪装成 XLSX 的宏工作簿和过多工作表拒绝',async()=>{
 const zip=await JSZip.loadAsync(await numericWorkbook('1'));zip.file('xl/vbaProject.bin','synthetic-only');await assert.rejects(importBytes(await zip.generateAsync({type:'nodebuffer'}),'renamed.xlsx'),/宏/);
 const wb=new ExcelJS.Workbook();for(let n=0;n<21;n++)wb.addWorksheet('表'+n).addRow(['合成']);await assert.rejects(importBytes(await wb.xlsx.writeBuffer(),'sheets.xlsx'),/20 张/);
});
test('超出异常总量时明确拒绝完整核对，不返回部分一致结果',async()=>{
 const a=await csv('编号,状态,金额\n'+Array.from({length:12501},(_,i)=>`${i},签约,1`).join('\n')),b=await csv('编号,状态,金额\n'+Array.from({length:12501},(_,i)=>`${i},认购,2`).join('\n'),'B.csv');
 assert.throws(()=>reconcile(a,b,cfg),/超过 25,000.*未输出部分结果/);
});
test('超长主键隔离，过多映射拒绝',async()=>{
 const a=await csv('编号,状态,金额\n'+'a'.repeat(257)+',签约,1'),b=await csv('编号,状态,金额\n001,签约,1','B.csv');assert.equal(reconcile(a,b,cfg).issues[0].type,'无效主键');
 const c=structuredClone(cfg);c.keys=Array.from({length:9},()=>c.keys[0]);assert.throws(()=>reconcile(a,b,c),/最多 8/);
});
test('导出层也拒绝不合法控制字符，不能静默删改原值',async()=>{
 const r=reconcile(await csv('编号,状态,金额\n001,签约,1'),await csv('编号,状态,金额\n001,认购,1','B.csv'),cfg);r.differences[0].leftValue='bad\u007F';await assert.rejects(reportBytes(r),/控制字符/);
});
