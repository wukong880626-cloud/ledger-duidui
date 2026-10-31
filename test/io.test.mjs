import {test} from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {importBytes,selectTable,parseCSV,reportBytes,safeText,LIMITS} from '../src/io.js';
import {reconcile} from '../src/core.js';
import {demoA,demoB,demoConfig} from '../src/fixtures.js';
const enc=s=>new TextEncoder().encode(s);
const cfg={keys:[{label:'编号',left:0,right:0}],fields:[{label:'状态',type:'text',left:1,right:1},{label:'金额',type:'amount',left:2,right:2,leftUnit:'元',rightUnit:'元'}],tolerance:'0.01'};
async function fromCSV(s,name='A.csv',header=1){return selectTable(await importBytes(enc(s),name),0,header);}

test('CSV BOM、引号逗号、双引号、跨行、CRLF 和物理行号',()=>{
 const rows=parseCSV('\uFEFF编号,备注,金额\r\n001,"甲,乙\r\n""丙""",100\r\n002,普通,0\r\n');
 assert.equal(rows.length,3);assert.equal(rows[1].cells[1].raw,'甲,乙\n"丙"');assert.equal(rows[2].rowNumber,4);assert.equal(rows[1].cells[0].raw,'001');
});
test('CSV 制表符、分号、末行无换行与空尾字段',()=>{
 assert.equal(parseCSV('编号;状态\n001;签约')[1].cells[0].raw,'001');assert.equal(parseCSV('编号\t金额\n001\t')[1].cells[1].raw,'');
});
test('CSV 引号缺失、嵌套非法、引号后垃圾均失败，不能给一致结论',async()=>{
 for(const s of ['编号,金额\n001,"10','编号,金额\n001,1"0','编号,金额\n001,"10"x'])await assert.rejects(importBytes(enc(s),'bad.csv'),/引号/);
});
test('非 UTF-8 解码失败；GB18030 可显式读取合成中文',async()=>{
 const bytes=Uint8Array.from([0xb1,0xe0,0xba,0xc5,0x2c,0xbd,0xf0,0xb6,0xee,10,48,48,49,44,49,48]);
 await assert.rejects(importBytes(bytes,'gb.csv'),/编码/);const book=await importBytes(bytes,'gb.csv','gb18030');assert.equal(book.sheets[0].rows[0].cells[0].raw,'编号');
});
test('XLSX 选择工作表与标题行，数字补零与文本前导零保留',async()=>{
 const wb=new ExcelJS.Workbook();wb.addWorksheet('说明').addRow(['请选择明细']);const ws=wb.addWorksheet('明细');
 ws.addRow(['合成说明']);ws.addRow(['编号','状态','金额']);ws.addRow([1,'签约',100]);ws.getCell('A3').numFmt='0000';ws.addRow(['0002','认购',200]);
 const book=await importBytes(await wb.xlsx.writeBuffer(),'leading.xlsx');assert.equal(book.sheets.length,2);assert.throws(()=>selectTable(book,0,1),/没有明细/);
 const data=selectTable(book,1,2);assert.equal(data.headerRow,2);assert.equal(data.rows[0].rowNumber,3);assert.equal(data.rows[0].cells[0].text,'0001');assert.equal(data.rows[1].cells[0].text,'0002');
 const other=await fromCSV('编号,状态,金额\n0002,认购,200\n0001,签约,100');assert.equal(reconcile(data,other,cfg).allConsistent,true);
});
test('XLSX 公式和 Excel 错误值无论缓存均保留并隔离',async()=>{
 const wb=new ExcelJS.Workbook(),ws=wb.addWorksheet('明细');ws.addRow(['编号','状态','金额']);ws.addRow(['001','签约',{formula:'1+1'}]);ws.addRow(['002',{formula:'"签约"',result:'签约'},2]);ws.addRow(['003','签约',{error:'#VALUE!'}]);ws.addRow([{formula:'"004"'},'签约',2]);
 const data=selectTable(await importBytes(await wb.xlsx.writeBuffer(),'formulas.xlsx'),0,1);assert.ok(data.warnings.some(x=>x.includes('公式')));assert.equal(data.rows[0].cells[2].raw,'');
 const b=await fromCSV('编号,状态,金额\n001,签约,2\n002,签约,2\n003,签约,0\n004,签约,2');const result=reconcile(data,b,cfg);assert.equal(result.allConsistent,false);assert.equal(result.issues.filter(x=>x.type==='关键字段含公式').length,3);assert.equal(result.summary.matched,0);
});
test('合并表头或数据区域阻止导入，标题之前说明合并允许',async()=>{
 for(const range of ['A2:B2','A3:A4']){
  const wb=new ExcelJS.Workbook(),ws=wb.addWorksheet('明细');ws.addRow(['说明']);ws.addRow(['编号','状态','金额']);ws.addRow(['001','签约',100]);ws.addRow(['002','签约',100]);ws.mergeCells(range);
  const book=await importBytes(await wb.xlsx.writeBuffer(),'merged.xlsx');assert.throws(()=>selectTable(book,0,2),/合并/);
 }
 const wb=new ExcelJS.Workbook(),ws=wb.addWorksheet('明细');ws.mergeCells('A1:C1');ws.getCell('A1').value='说明';ws.addRow(['编号','状态','金额']);ws.addRow(['001','签约',100]);assert.equal(selectTable(await importBytes(await wb.xlsx.writeBuffer(),'ok.xlsx'),0,2).rows.length,1);
});
test('损坏、加密/非 ZIP、宏/旧格式和超限文件拒绝',async()=>{
 await assert.rejects(importBytes(enc('bad'),'bad.xlsx'),/格式无效/);await assert.rejects(importBytes(enc('PKbad'),'bad.xlsx'),/解析失败/);
 for(const name of ['macro.xlsm','old.xls','plain.txt'])await assert.rejects(importBytes(enc('x'),name),/仅支持/);
 await assert.rejects(importBytes(new Uint8Array(LIMITS.bytes+1),'big.csv'),/20 MB/);
});
test('空标题、越界、无数据和无标题右侧数据拒绝',async()=>{
 for(const [s,header] of [['\n001,1',1],['编号,金额\n001,1',3],['编号,金额',1],['说明\n编号,金额\n001,1',1]]){
  const book=await importBytes(enc(s),'bad.csv');assert.throws(()=>selectTable(book,0,header));
 }
});
test('完整异常样本 CSV 与 XLSX 得到相同摘要和行号',async()=>{
 const a=await fromCSV(demoA,'异常示例_A.csv',2),b=await fromCSV(demoB,'异常示例_B.csv');const r=reconcile(a,b,demoConfig);
 assert.deepEqual(r.summary,{left:13,right:10,matched:7,agreed:3,changed:4,missingLeft:1,missingRight:1,differenceCount:6,issueCount:7,invalidRows:7,duplicateRows:2});
 const ax=selectTable(await importBytes(await readFile('samples/异常示例_A.xlsx'),'异常示例_A.xlsx'),1,2),bx=selectTable(await importBytes(await readFile('samples/异常示例_B.xlsx'),'异常示例_B.xlsx'),1,1);
 assert.deepEqual(reconcile(ax,bx,demoConfig).summary,r.summary);assert.equal(r.differences.find(x=>x.type==='状态/文本不同').leftRow,4);
});
test('报告往返读取，三张表、完整来源、金额差与全部异常不丢失',async()=>{
 const r=reconcile(await fromCSV(demoA,'A.csv',2),await fromCSV(demoB,'B.csv'),demoConfig),bytes=await reportBytes(r),wb=new ExcelJS.Workbook();await wb.xlsx.load(bytes);
 assert.deepEqual(wb.worksheets.map(x=>x.name),['概览','差异明细','重复与无效记录']);assert.equal(wb.getWorksheet('差异明细').rowCount,7);assert.equal(wb.getWorksheet('重复与无效记录').rowCount,8);
 const rows=wb.getWorksheet('差异明细').getRows(2,6);assert.ok(rows.some(row=>row.getCell(12).value==="'-0.02"));assert.ok(rows.some(row=>row.getCell(6).value===4&&row.getCell(10).value===9));assert.equal(wb.getWorksheet('概览').getCell('B2').value,'存在待处理记录，不能确认全部一致');
});
test('恶意文本导出无公式、无外链；防止 Excel 注入与 HTML 载荷丢失',async()=>{
 const a=await fromCSV('编号,状态,金额\n001,=HYPERLINK(恶意),1\n002,<img src=x onerror=alert(1)>,2','@坏文件.csv');
 const b=await fromCSV('编号,状态,金额\n001,普通,0\n002,普通,0','B.csv');const bytes=await reportBytes(reconcile(a,b,cfg));
 const zip=await JSZip.loadAsync(bytes);for(const name of Object.keys(zip.files).filter(n=>/\.xml$/.test(n))){const xml=await zip.file(name).async('string');assert.doesNotMatch(xml,/<f[ >]/);assert.doesNotMatch(xml,/<externalLink/);}
 const wb=new ExcelJS.Workbook();await wb.xlsx.load(bytes);const row=wb.getWorksheet('差异明细').getRow(2);assert.match(row.getCell(7).value,/^'=HYPERLINK/);assert.equal(typeof row.getCell(7).value,'string');assert.equal(row.getCell(4).value,"'@坏文件.csv");
 for(const s of ['=1',' +1','-1','@cmd','\tcalc','\n=cmd'])assert.equal(safeText(s),"'"+s);
});
