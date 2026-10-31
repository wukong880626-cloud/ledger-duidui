import {execFileSync} from 'node:child_process';
import {mkdir,writeFile,readFile} from 'node:fs/promises';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
await mkdir('output/playwright',{recursive:true});
const logs=[],cli='node_modules/@playwright/cli/playwright-cli.js';
function run(...args){
 const out=execFileSync(process.execPath,[cli,'-s=ledger-final',...args],{encoding:'utf8',timeout:90000});logs.push(out);
 if(out.includes('### Error'))throw new Error(out);return out;
}
try{
 run('open',process.env.LEDGER_QA_URL||'http://127.0.0.1:4173','--browser','chrome','--headed');
 run('snapshot');
 run('run-code',String(async page=>{
   await page.getByRole('button',{name:'试用合成示例'}).click();
   await page.getByRole('button',{name:'开始核对 →'}).click();
   await page.getByRole('button',{name:'导出 XLSX 报告 ↓'}).waitFor();
   const stats=await page.locator('.stat strong').allTextContents();
   if(stats.join(',')!=='7,3,4,1,1,7')throw new Error('最终构建结果不符：'+stats);
   await page.screenshot({path:'output/playwright/08-final-demo.png',fullPage:true});
   await page.context().setOffline(true);
   await page.getByRole('button',{name:'清空本次数据'}).click();
   await page.getByLabel('选择台账 A 文件',{exact:true}).setInputFiles('samples/安全文本_A.csv');
   await page.getByLabel('选择台账 B 文件',{exact:true}).setInputFiles('samples/安全文本_B.csv');
   return {finalDemo:stats,offlineImport:true};
 }));
 run('snapshot');
 run('run-code',String(async page=>{
   for(const [name,v] of [['主键 1 A 列','0'],['主键 1 B 列','0'],['核对字段 1 A 列','1'],['核对字段 1 B 列','1'],['核对字段 2 A 列','2'],['核对字段 2 B 列','2']])await page.getByRole('combobox',{name,exact:true}).selectOption(v);
   await page.getByRole('button',{name:'开始核对 →'}).click();
   await page.getByRole('button',{name:'导出 XLSX 报告 ↓'}).waitFor();
   if(await page.locator('#result-table tbody tr').count()!==2)throw new Error('安全文本差异数量错误');
   if(await page.locator('#result-table img').count()!==0)throw new Error('HTML 文本被解释为元素');
   if(!(await page.locator('#result-table').textContent()).includes('<img src=x onerror=alert(1)>'))throw new Error('HTML 原值丢失');
   if(await page.evaluate(()=>navigator.onLine))throw new Error('浏览器未进入离线状态');
   const dl=page.waitForEvent('download');await page.getByRole('button',{name:'导出 XLSX 报告 ↓'}).click();
   await(await dl).saveAs('output/playwright/browser-offline-safe-export.xlsx');
   await page.screenshot({path:'output/playwright/09-offline-safe-text.png',fullPage:true});
   await page.context().setOffline(false);
   return {offlineCompareAndDownload:true,htmlEscaped:true,differences:2};
 }));
 const network=run('requests','--static');
 if(/https?:\/\/(?!127\.0\.0\.1:4173)/.test(network))throw new Error('出现非预期外部网络请求');
 const consoleOutput=run('console');if(!consoleOutput.includes('Errors: 0, Warnings: 0'))throw new Error('浏览器控制台有告警');
 const bytes=await readFile('output/playwright/browser-offline-safe-export.xlsx'),wb=new ExcelJS.Workbook();await wb.xlsx.load(bytes);
 if(wb.getWorksheet('差异明细').getCell('G3').value!=="'=1+1")throw new Error('公式注入防护缺失');
 const zip=await JSZip.loadAsync(bytes);for(const name of Object.keys(zip.files).filter(n=>/\.xml$/.test(n))){if(/<f[ >]|<externalLink/.test(await zip.file(name).async('string')))throw new Error('报告含公式或外链');}
 console.log('最终浏览器回归通过：演示、离线导入/核对/下载、HTML 转义、公式注入防护、无外部请求、控制台零告警。');
}finally{await writeFile('output/playwright/final-browser-check.log',logs.join('\n'));}
