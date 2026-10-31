import JSZip from 'jszip';
import {SaxesParser} from 'saxes';
import Decimal from 'decimal.js';
import {LIMITS} from './limits.js';

// Read declared sizes before decompression. ZIP64/multidisk are outside this prototype.
export function checkZipBudget(bytes){
 const view=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);let end=-1;
 for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--)if(view.getUint32(i,true)===0x06054b50&&i+22+view.getUint16(i+20,true)===bytes.length){end=i;break;}
 if(end<0)throw new Error('XLSX ZIP 目录无效，解析失败');
 const count=view.getUint16(end+10,true),offset=view.getUint32(end+16,true);
 if(view.getUint16(end+4,true)||view.getUint16(end+6,true)||count===65535||offset===0xffffffff)throw new Error('首版不支持分卷或 ZIP64 工作簿');
 if(count>LIMITS.zipEntries)throw new Error('XLSX 内部文件过多（超过 1,000 项），请整理副本');
 let cursor=offset,total=0;
 for(let n=0;n<count;n++){
  if(cursor+46>end||view.getUint32(cursor,true)!==0x02014b50)throw new Error('XLSX ZIP 目录损坏，解析失败');
  const size=view.getUint32(cursor+24,true);if(size===0xffffffff)throw new Error('首版不支持 ZIP64 工作簿');
  total+=size;if(total>LIMITS.expandedBytes)throw new Error('XLSX 声明解压体积超过 50 MB，请拆分工作簿');
  cursor+=46+view.getUint16(cursor+28,true)+view.getUint16(cursor+30,true)+view.getUint16(cursor+32,true);
 }
 return total;
}
const local=name=>name.split(':').pop();
function parser(){const p=new SaxesParser();p.on('doctype',()=>{throw new Error('不支持含 DTD 的 XLSX');});return p;}
function attr(node,name){return Object.entries(node.attributes).find(([k])=>local(k)===name)?.[1];}
function pathInZip(target){
 if(/^[a-z]+:/i.test(target))throw new Error('工作表指向外部地址，无法导入');
 const parts=(target.startsWith('/')?target.slice(1):'xl/'+target).split('/'),stack=[];
 for(const part of parts){if(part==='..'){if(!stack.length)throw new Error('工作表路径无效');stack.pop();}else if(part&&part!=='.')stack.push(part);}
 return stack.join('/');
}
function unsafeNumeric(raw){
 try{const d=new Decimal(raw),n=Number(raw);return !Number.isFinite(n)||d.sd()>15||!d.eq(String(n));}catch{return true;}
}
export async function precheckXlsx(bytes){
 checkZipBudget(bytes);
 let zip;try{zip=await JSZip.loadAsync(bytes);}catch{throw new Error('XLSX ZIP 解析失败');}
 if(Object.keys(zip.files).some(n=>/vbaProject\.bin$/i.test(n)))throw new Error('工作簿含 VBA 宏，首版不接受宏文件');
 const workbook=zip.file('xl/workbook.xml'),relsFile=zip.file('xl/_rels/workbook.xml.rels');
 if(!workbook||!relsFile)throw new Error('缺少 XLSX 工作簿结构，解析失败');
 const sheets=[],rels=new Map();let p=parser();
 p.on('opentag',n=>{if(local(n.name)==='sheet')sheets.push({name:attr(n,'name'),id:attr(n,'id')});});p.write(await workbook.async('string')).close();
 p=parser();p.on('opentag',n=>{if(local(n.name)==='Relationship')rels.set(attr(n,'Id'),{target:attr(n,'Target'),external:attr(n,'TargetMode')==='External'});});p.write(await relsFile.async('string')).close();
 if(sheets.length>LIMITS.sheets)throw new Error('工作簿超过 20 张工作表，请整理副本');
 const result=new Map();let totalCells=0;
 for(const sheet of sheets){
  const rel=rels.get(sheet.id);if(!rel?.target||rel.external)throw new Error('工作表关系无效或指向外部地址');
  const entry=zip.file(pathInZip(rel.target));if(!entry)throw new Error('工作表 XML 缺失');
  const unsafe=new Map();let cell=null,inValue=false;
  p=parser();
  p.on('opentag',n=>{
   const name=local(n.name);
   if(name==='row'&&Number(attr(n,'r'))>LIMITS.rows+LIMITS.headerRecords)throw new Error('工作表行号超过 25,100（含标题/说明行），请拆分');
   if(name==='c'){
    totalCells++;if(totalCells>LIMITS.cells)throw new Error('工作簿超过 150,000 个单元格（含标题及所有工作表），请拆分');
    const ref=attr(n,'r'),type=attr(n,'t')||'n',match=/^([A-Z]+)(\d+)$/.exec(ref||'');
    if(match){let col=0;for(const c of match[1])col=col*26+c.charCodeAt(0)-64;if(col>LIMITS.columns||Number(match[2])>LIMITS.rows+LIMITS.headerRecords)throw new Error('工作表超出 200 列 / 25,100 行范围');}
    cell={ref,type,raw:'',formula:false};
   }
   if(cell&&name==='v')inValue=true;if(cell&&name==='f')cell.formula=true;
  });
  p.on('text',text=>{if(cell&&inValue)cell.raw+=text;});
  p.on('closetag',n=>{const name=local(n.name);if(name==='v')inValue=false;if(name==='c'&&cell){if(cell.type==='n'&&!cell.formula&&cell.raw&&unsafeNumeric(cell.raw))unsafe.set(cell.ref,cell.raw);cell=null;}});
  p.write(await entry.async('string')).close();result.set(sheet.name,unsafe);
 }
 return result;
}
