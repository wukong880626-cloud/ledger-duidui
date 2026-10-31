import { build } from 'esbuild';
import { mkdir, copyFile, readdir, readFile, writeFile } from 'node:fs/promises';
await mkdir('dist', { recursive: true });
await build({ entryPoints:['src/app.js'], outfile:'dist/app.js', bundle:true, platform:'browser', format:'iife', target:['es2020'], minify:true, legalComments:'external' });
await build({ entryPoints:['src/worker.js'], outfile:'dist/worker.js', bundle:true, platform:'browser', format:'iife', target:['es2020'], minify:true, legalComments:'external' });
for(const file of ['index.html','style.css','privacy.html']) await copyFile(`src/${file}`,`dist/${file}`);
await copyFile('LICENSE','dist/LICENSE.txt');
// Preserve the self-code MIT license and all dependency notices in the bundle.
const notices=[];
async function scan(dir){
  for(const entry of await readdir(dir,{withFileTypes:true})){
    if(!entry.isDirectory()||entry.name.startsWith('.'))continue;
    const path=`${dir}/${entry.name}`;
    if(entry.name.startsWith('@')){await scan(path);continue;}
    try{
      const pkg=JSON.parse(await readFile(`${path}/package.json`,'utf8'));
      const files=await readdir(path),license=files.find(n=>/^licen[sc]e(?:\.md|\.txt)?$/i.test(n));
      notices.push(`\n===== ${pkg.name} ${pkg.version} (${typeof pkg.license==='string'?pkg.license:'see package'}) =====\n${license?await readFile(`${path}/${license}`,'utf8'):''}`);
    }catch{}
    try{await scan(`${path}/node_modules`);}catch{}
  }
}
await scan('node_modules');await writeFile('dist/THIRD_PARTY_NOTICES.txt',notices.join('\n'));
console.log('已生成 dist/index.html；所有运行依赖均为本地文件。');
