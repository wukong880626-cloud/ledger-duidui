import {execFileSync,spawn} from 'node:child_process';
import {mkdir,writeFile,readFile} from 'node:fs/promises';
import ExcelJS from 'exceljs';
const label=process.argv[2]||'baseline',specs=process.argv.slice(3).length?process.argv.slice(3):['csv:1000','csv:10000','csv:25000','xlsx:1000','xlsx:10000'];
await mkdir('output/round2',{recursive:true});const log=[],results=[];
function run(...args){const out=execFileSync(process.execPath,['node_modules/@playwright/cli/playwright-cli.js','-s=ledger-perf',...args],{encoding:'utf8',timeout:90000,maxBuffer:10*1024*1024});log.push(out);if(out.includes('### Error'))throw new Error(out);return out;}
function result(out){const match=out.match(/### Result\n([\s\S]*?)\n### Ran/);if(!match)throw Error('缺少测量返回值');return JSON.parse(match[1]);}
let rendererIds=[];
function rss(){try{return execFileSync('/bin/ps',['-o','rss=','-p',rendererIds.join(',')],{encoding:'utf8'}).trim().split(/\s+/).reduce((n,v)=>n+Number(v||0),0)/1024;}catch{return null;}}
async function measured(code){
 const samples=[];const sample=()=>{const n=rss();if(n!==null)samples.push(n);};sample();const timer=setInterval(sample,50);
 try{
  const out=await new Promise((resolve,reject)=>{let stdout='',stderr='';const child=spawn(process.execPath,['node_modules/@playwright/cli/playwright-cli.js','-s=ledger-perf','run-code',code]);const timeout=setTimeout(()=>{child.kill();reject(Error('浏览器测量超过90秒'));},90000);child.stdout.on('data',c=>stdout+=c);child.stderr.on('data',c=>stderr+=c);child.on('close',status=>{clearTimeout(timeout);if(status||stdout.includes('### Error'))reject(Error(stdout+stderr));else resolve(stdout);});});
  log.push(out);sample();return {...result(out),sampledRendererRSSPeakMB:samples.length?Math.max(...samples):null,rssSamples:samples.length};
 }finally{clearInterval(timer);}
}
try{
 run('open',process.env.LEDGER_QA_URL||'http://127.0.0.1:4173','--browser','chrome');
 for(const spec of specs){
  const [format,count]=spec.split(':'),n=Number(count),download=`output/round2/${label}_${format}_${n}.xlsx`,extension=format==='xlsxnum'?'xlsx':format,suffix=format==='xlsxnum'?'_numeric':'';
  run('goto',process.env.LEDGER_QA_URL||'http://127.0.0.1:4173');run('snapshot');
  rendererIds=result(run('run-code',String(async page=>{const cdp=await page.context().browser().newBrowserCDPSession();const info=await cdp.send('SystemInfo.getProcessInfo');await cdp.detach();return info.processInfo.filter(p=>p.type==='renderer').map(p=>p.id);})));
  const importCode=`async page=>{
    const cdp=await page.context().newCDPSession(page);await cdp.send('HeapProfiler.collectGarbage');
    await page.evaluate(()=>{window.__perf={last:performance.now(),maxGap:0,longTasks:[]};window.__perf.timer=setInterval(()=>{const t=performance.now();window.__perf.maxGap=Math.max(window.__perf.maxGap,t-window.__perf.last);window.__perf.last=t;},20);window.__perf.observer=new PerformanceObserver(list=>window.__perf.longTasks.push(...list.getEntries().map(e=>e.duration)));window.__perf.observer.observe({entryTypes:['longtask']});});
    const start=Date.now(),before=(await cdp.send('Runtime.getHeapUsage')).usedSize;
    for(const side of ['A','B']){await page.getByLabel('选择台账 '+side+' 文件',{exact:true}).setInputFiles('output/round2/fixtures/${n}_'+side+'${suffix}.${extension}');await page.getByRole('region',{name:'台账 '+side,exact:true}).getByText('${n} 条明细',{exact:true}).waitFor({timeout:60000});}
    await page.waitForTimeout(60);const perf=await page.evaluate(()=>({maxGapMs:window.__perf.maxGap,maxLongTaskMs:Math.max(0,...window.__perf.longTasks)})),after=(await cdp.send('Runtime.getHeapUsage')).usedSize;
    await page.evaluate(()=>{clearInterval(window.__perf.timer);window.__perf.observer.disconnect();});return {ms:Date.now()-start,heapBeforeMB:before/1048576,heapAfterMB:after/1048576,...perf};}`;
  const imported=await measured(importCode);
  run('run-code',String(async page=>{await page.getByRole('button',{name:'＋ 添加主键'}).click();await page.getByRole('button',{name:'＋ 添加主键'}).click();}));run('snapshot');
  const compareCode=`async page=>{
    for(const [i,l,r] of [[1,'0','4'],[2,'1','3'],[3,'2','0']]){await page.getByRole('combobox',{name:'主键 '+i+' A 列',exact:true}).selectOption(l);await page.getByRole('combobox',{name:'主键 '+i+' B 列',exact:true}).selectOption(r);}
    for(const [name,v] of [['核对字段 1 A 列','3'],['核对字段 1 B 列','2'],['核对字段 2 A 列','4'],['核对字段 2 B 列','1'],['核对字段 2 B 金额单位','万元']])await page.getByRole('combobox',{name,exact:true}).selectOption(v);
    const cdp=await page.context().newCDPSession(page);await page.evaluate(()=>{window.__perf={last:performance.now(),maxGap:0,longTasks:[]};window.__perf.timer=setInterval(()=>{const t=performance.now();window.__perf.maxGap=Math.max(window.__perf.maxGap,t-window.__perf.last);window.__perf.last=t;},20);window.__perf.observer=new PerformanceObserver(list=>window.__perf.longTasks.push(...list.getEntries().map(e=>e.duration)));window.__perf.observer.observe({entryTypes:['longtask']});});const start=Date.now();
    await page.getByRole('button',{name:'开始核对 →'}).click();await page.getByRole('button',{name:'导出 XLSX 报告 ↓'}).waitFor({timeout:60000});await page.waitForTimeout(60);const stats=await page.locator('.stat strong').allTextContents();if(stats.join(',')!=='${n},0,${n},0,0,0')throw Error('结果不符 '+stats);const perf=await page.evaluate(()=>({maxGapMs:window.__perf.maxGap,maxLongTaskMs:Math.max(0,...window.__perf.longTasks)})),heap=(await cdp.send('Runtime.getHeapUsage')).usedSize;await page.evaluate(()=>{clearInterval(window.__perf.timer);window.__perf.observer.disconnect();});return {ms:Date.now()-start,heapAfterMB:heap/1048576,stats,...perf};}`;
  const compared=await measured(compareCode);
  const exportCode=`async page=>{
    const cdp=await page.context().newCDPSession(page);await page.evaluate(()=>{window.__perf={last:performance.now(),maxGap:0,longTasks:[]};window.__perf.timer=setInterval(()=>{const t=performance.now();window.__perf.maxGap=Math.max(window.__perf.maxGap,t-window.__perf.last);window.__perf.last=t;},20);window.__perf.observer=new PerformanceObserver(list=>window.__perf.longTasks.push(...list.getEntries().map(e=>e.duration)));window.__perf.observer.observe({entryTypes:['longtask']});});const start=Date.now(),dl=page.waitForEvent('download',{timeout:60000});await page.getByRole('button',{name:'导出 XLSX 报告 ↓'}).click();await(await dl).saveAs('${download}');await page.waitForTimeout(60);const perf=await page.evaluate(()=>({maxGapMs:window.__perf.maxGap,maxLongTaskMs:Math.max(0,...window.__perf.longTasks)})),heap=(await cdp.send('Runtime.getHeapUsage')).usedSize;await page.evaluate(()=>{clearInterval(window.__perf.timer);window.__perf.observer.disconnect();});return {ms:Date.now()-start,heapAfterMB:heap/1048576,...perf};}`;
  const exported=await measured(exportCode),wb=new ExcelJS.Workbook();await wb.xlsx.load(await readFile(download));if(wb.getWorksheet('差异明细').rowCount!==n+1)throw Error('导出记录丢失');
  const entry={format,rowsPerSide:n,columns:5,allStatusDifferent:true,import:imported,compare:compared,export:exported,exportedDifferences:n};results.push(entry);await writeFile(`output/round2/${label}-performance.json`,JSON.stringify(results,null,2));console.log(JSON.stringify(entry));
 }
}finally{await writeFile(`output/round2/${label}-performance.log`,log.join('\n'));}
