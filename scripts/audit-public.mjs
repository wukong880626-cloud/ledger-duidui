import {readdir,readFile,lstat,mkdir,writeFile} from 'node:fs/promises';
import {execFileSync} from 'node:child_process';
import {createHash} from 'node:crypto';
import JSZip from 'jszip';
const roots=['src','dist','test','scripts','samples','.github'];
const names=['README.md','LICENSE','PRIVACY.md','package.json','package-lock.json','.gitignore','测试结果.md','第二轮验收_v0.1.1.md','更新说明_v0.1.1.md','启动台账对对看.command','打开台账对对看.html','发布验收.md'];
const files=[];
async function walk(path){const s=await lstat(path);if(s.isSymbolicLink())throw Error('公开内容不能包含符号链接：'+path);if(s.isDirectory()){for(const n of(await readdir(path)).sort())await walk(path+'/'+n);}else files.push(path);}
for(const r of roots)await walk(r);for(const n of names){try{await walk(n);}catch(e){if(e.code!=='ENOENT')throw e;}}
const allowed=new Set(files);let tracked=[];try{tracked=execFileSync('git',['ls-files','-z'],{encoding:'utf8',stdio:['ignore','pipe','ignore']}).split('\0').filter(Boolean);}catch{}
for(const p of tracked)if(!allowed.has(p))throw Error('跟踪内容不在公开白名单：'+p);
const personal=new RegExp('[/]'+'(?:Users|home)'+'[/][^\\s<>"\u0027]+','i');
const credentials=/(?:github_pat_[A-Za-z0-9_]{20,}|gh[pousr]_[A-Za-z0-9]{20,}|sk-[A-Za-z0-9]{24,}|AKIA[0-9A-Z]{16}|-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----)/;
const failures=[],manifest=[];function check(t,p){if(personal.test(t))failures.push(p+': 用户目录路径');if(credentials.test(t))failures.push(p+': 疑似凭据');}
for(const p of files){
 if(/(?:^|\/)(?:output|release|node_modules|\.playwright-cli|\.codex|\.agents|\.aws)(?:\/|$)/.test(p))throw Error('公开清单含本机目录：'+p);
 const b=await readFile(p);if(p.endsWith('.xlsx')){const z=await JSZip.loadAsync(b);for(const e of Object.values(z.files))if(!e.dir&&/\.xml$|\.rels$/i.test(e.name))check(await e.async('string'),p+':'+e.name);}else check(b.toString('utf8'),p);
 manifest.push({path:p,bytes:b.length,sha256:createHash('sha256').update(b).digest('hex')});
}
if(failures.length)throw Error('公开审查未通过：\n'+failures.join('\n'));
await mkdir('output/publish',{recursive:true});await writeFile('output/publish/public-manifest.json',JSON.stringify({files:manifest,reviewedAt:new Date().toISOString(),scope:'代码、说明、第三方许可、合成样本；排除本机日志、压测输出、ZIP、用户目录'},null,2));
console.log(`公开白名单审查通过：${files.length} 个文件；未发现用户目录路径或凭据特征。样本来源另由生成脚本与交付校验确认。`);
