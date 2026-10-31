import {execFileSync} from 'node:child_process';
import {mkdir,readFile,writeFile} from 'node:fs/promises';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
const url=process.env.LEDGER_QA_URL||'http://127.0.0.1:4175/ledger-duidui/',base=new URL(url),label=process.env.LEDGER_CHECK_LABEL||'local-subpath';
if(!/^[a-z0-9-]+$/.test(label))throw Error('验收标签无效');
await mkdir('output/publish',{recursive:true});const logs=[];
function run(...args){const out=execFileSync(process.execPath,['node_modules/@playwright/cli/playwright-cli.js','-s=ledger-publish',...args],{encoding:'utf8',timeout:60000,maxBuffer:5*1024*1024});logs.push(out);if(out.includes('### Error'))throw Error(out);return out;}
try{
 run('open','about:blank','--browser','chrome');
 run('run-code',`async page=>{await page.context().setOffline(false);await page.context().unroute('**/*');await page.context().addInitScript(()=>{window.__publishRequests=[];});page.context().on('request',r=>{page.__publishRequests??=[];page.__publishRequests.push({url:r.url(),method:r.method(),type:r.resourceType()});});await page.goto(${JSON.stringify(url)});}`);run('snapshot');
 run('run-code',String(async page=>{
  await page.getByLabel('选择台账 A 文件',{exact:true}).setInputFiles('samples/异常示例_A.csv');
  await page.getByRole('spinbutton',{name:'台账 A 标题行'}).fill('2');await page.getByRole('spinbutton',{name:'台账 A 标题行'}).press('Tab');
  await page.getByLabel('选择台账 B 文件',{exact:true}).setInputFiles('samples/异常示例_B.csv');
  await page.getByRole('button',{name:'＋ 添加主键'}).click();await page.getByRole('button',{name:'＋ 添加主键'}).click();
 }));run('snapshot');
 run('run-code',`async page=>{
  for(const [i,l,r] of [[1,'0','4'],[2,'1','3'],[3,'2','0']]){await page.getByRole('combobox',{name:'主键 '+i+' A 列',exact:true}).selectOption(l);await page.getByRole('combobox',{name:'主键 '+i+' B 列',exact:true}).selectOption(r);}
  for(const [name,v] of [['核对字段 1 A 列','3'],['核对字段 1 B 列','2'],['核对字段 2 A 列','4'],['核对字段 2 B 列','1'],['核对字段 2 B 金额单位','万元']])await page.getByRole('combobox',{name,exact:true}).selectOption(v);
  await page.getByRole('button',{name:'开始核对 →'}).click();await page.getByRole('button',{name:'导出 XLSX 报告 ↓'}).waitFor({timeout:15000});
  const stats=await page.locator('.stat strong').allTextContents();if(stats.join(',')!=='7,3,4,1,1,7')throw Error('线上核对结果不符 '+stats);
  const d=page.waitForEvent('download');await page.getByRole('button',{name:'导出 XLSX 报告 ↓'}).click();await(await d).saveAs('output/publish/${label}-report.xlsx');
  await page.context().setOffline(true);await page.getByRole('button',{name:'开始核对 →'}).click();await page.getByRole('button',{name:'导出 XLSX 报告 ↓'}).waitFor();
  const offline=page.waitForEvent('download');await page.getByRole('button',{name:'导出 XLSX 报告 ↓'}).click();await(await offline).saveAs('output/publish/${label}-offline-report.xlsx');
  await page.context().setOffline(false);await page.screenshot({path:'output/playwright/publish-${label}.png',fullPage:true});
  const requests=page.__publishRequests||[],base=new URL(${JSON.stringify(url)});
  if(!requests.some(r=>new URL(r.url).pathname===base.pathname+'worker.js'))throw Error('Worker 未从站点子路径加载');
  if(requests.some(r=>r.method!=='GET'||new URL(r.url).origin!==base.origin||!new URL(r.url).pathname.startsWith(base.pathname)))throw Error('检测到非预期请求，不能确认无上传');
  return {stats,requests,offlineCompareAndDownload:true};
 }`);
 const requests=run('requests','--static');const consoleOutput=run('console');if(!consoleOutput.includes('Errors: 0, Warnings: 0'))throw Error('控制台存在错误/告警');
 const artifacts=[];for(const suffix of ['report','offline-report']){
  const path=`output/publish/${label}-${suffix}.xlsx`,bytes=await readFile(path),wb=new ExcelJS.Workbook();await wb.xlsx.load(bytes);
  if(wb.getWorksheet('差异明细').rowCount!==7||wb.getWorksheet('重复与无效记录').rowCount!==8)throw Error('导出明细丢失');
  if(!wb.getWorksheet('差异明细').getCell('B2').value.includes('002'))throw Error('前导零丢失');
  const zip=await JSZip.loadAsync(bytes);for(const entry of Object.values(zip.files))if(!entry.dir&&/\.xml$/.test(entry.name)&&/<f[ >]|<externalLink/.test(await entry.async('string')))throw Error('报告有公式或外链');
  artifacts.push({path,sheets:wb.worksheets.map(s=>s.name),differences:6,issues:7});
 }
 await writeFile(`output/publish/${label}-verification.json`,JSON.stringify({url,passed:true,realFileImport:true,offlineCompute:true,noFileUploadObserved:true,consoleErrors:0,artifacts},null,2));
 console.log(JSON.stringify({url,passed:true,realImportCompareDownload:true,offline:true,workerSubpath:base.pathname+'worker.js',requests:requests.trim()},null,2));
}finally{await writeFile(`output/publish/${label}-browser.log`,logs.join('\n'));}
