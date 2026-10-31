import ExcelJS from 'exceljs';
import Decimal from 'decimal.js';
import {mkdir,writeFile} from 'node:fs/promises';
const root='output/round2/fixtures';await mkdir(root,{recursive:true});
for(const n of [1000,10000,25000,50000]){
 const a=[['楼栋','单元','房号','状态','金额（元）']],b=[['房间编号','金额（万元）','销售状态','单元号','楼号']];
 for(let i=0;i<n;i++){
  const building=String(Math.floor(i/10000)+1).padStart(2,'0'),unit=String(Math.floor(i%10000/1000)+1).padStart(2,'0'),room=String(i%1000+1).padStart(4,'0');
  const amount=new Decimal(1000000).plus(new Decimal(i%100).div(100)).toFixed(2);
  a.push([building,unit,room,'已签约',amount]);b.push([room,new Decimal(amount).div(10000).toFixed(6),'认购',unit,building]);
 }
 const reversed=[b[0],...b.slice(1).reverse()];
 for(const [side,rows] of [['A',a],['B',reversed]]){
  await writeFile(`${root}/${n}_${side}.csv`,'\uFEFF'+rows.map(r=>r.join(',')).join('\n'));
  if(n<=25000){const wb=new ExcelJS.Workbook(),ws=wb.addWorksheet('合成明细');for(const row of rows)ws.addRow(row);await wb.xlsx.writeFile(`${root}/${n}_${side}.xlsx`);}
  if(n<=25000){const wb=new ExcelJS.Workbook(),ws=wb.addWorksheet('合成明细');for(let i=0;i<rows.length;i++){const row=[...rows[i]],column=side==='A'?4:1;if(i)row[column]=Number(row[column]);ws.addRow(row);if(i)ws.getRow(i+1).getCell(column+1).numFmt=side==='A'?'0.00':'0.000000';}await wb.xlsx.writeFile(`${root}/${n}_${side}_numeric.xlsx`);}
 }
 console.log(`已生成 ${n} 条/侧，5 列，复合主键；全部状态不同，金额换算相同。`);
}
