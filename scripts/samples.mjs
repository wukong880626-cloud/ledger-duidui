import ExcelJS from 'exceljs';
import { mkdir, writeFile } from 'node:fs/promises';
import { demoA, demoB, demoConfig } from '../src/fixtures.js';
import { importBytes, selectTable, reportBytes, parseCSV } from '../src/io.js';
import { reconcile } from '../src/core.js';
await mkdir('samples',{recursive:true});
await writeFile('samples/异常示例_A.csv','\uFEFF'+demoA);
await writeFile('samples/异常示例_B.csv','\uFEFF'+demoB);
for(const [side,text] of [['A',demoA],['B',demoB]]){
  const wb=new ExcelJS.Workbook();wb.addWorksheet('使用提示').addRow(['纯合成数据，请选择“销售明细”工作表；A 标题第2行，B 标题第1行。']);
  const ws=wb.addWorksheet('销售明细');parseCSV(text).forEach(r=>ws.addRow(r.cells.map(c=>c.raw)));ws.columns.forEach(c=>c.width=22);
  await wb.xlsx.writeFile(`samples/异常示例_${side}.xlsx`);
}
const cleanA='编号,状态,金额（元）\n0001,已签约,10000\n0002,认购,20000\n';
const cleanB='金额（万元）,房间编号,销售状态\n2,0002,认购\n1,0001,已签约\n';
await writeFile('samples/一致示例_A.csv','\uFEFF'+cleanA);await writeFile('samples/一致示例_B.csv','\uFEFF'+cleanB);
const formula=new ExcelJS.Workbook(),fws=formula.addWorksheet('公式示例');
fws.addRow(['编号','状态','金额']);fws.addRow(['0001','已签约',{formula:'1+1'}]);fws.addRow(['0002',{formula:'"已签约"',result:'已签约'},100]);
fws.addRow([{formula:'"0003"'},'已签约',100]);await formula.xlsx.writeFile('samples/公式拦截示例.xlsx');
await writeFile('samples/解析失败示例.csv','编号,状态,金额\n0001,"未闭合,100\n');
await writeFile('samples/公式对侧示例.csv','\uFEFF编号,状态,金额\n0001,已签约,2\n0002,已签约,100\n0003,已签约,100\n');
for(const side of ['A','B']){
  const rows=Array.from({length:121},(_,i)=>`${String(i+1).padStart(4,'0')},${side==='A'?'已签约':'认购'},100`);
  await writeFile(`samples/分页示例_${side}.csv`,'\uFEFF编号,状态,金额\n'+rows.join('\n'));
}
await writeFile('samples/安全文本_A.csv','\uFEFF编号,状态,金额\n0001,<img src=x onerror=alert(1)>,100\n0002,=1+1,100\n');
await writeFile('samples/安全文本_B.csv','\uFEFF编号,状态,金额\n0001,普通,100\n0002,普通,100\n');
const a=selectTable(await importBytes(new TextEncoder().encode(demoA),'异常示例_A.csv'),0,2),b=selectTable(await importBytes(new TextEncoder().encode(demoB),'异常示例_B.csv'),0,1);
const result=reconcile(a,b,demoConfig);await writeFile('samples/预期核对报告.xlsx',await reportBytes(result));await writeFile('samples/预期结果.json',JSON.stringify(result,null,2));
console.log(JSON.stringify(result.summary));
