import http from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {spawn} from 'node:child_process';
const root=path.resolve(fileURLToPath(new URL('../dist/',import.meta.url)));
const port=Number(process.env.PORT||4173);
const basePath=(process.env.LEDGER_BASE_PATH||'').replace(/\/+$/,'');
if(basePath&&(!basePath.startsWith('/')||basePath.includes('..')))throw new Error('本机服务子路径无效');
const server=http.createServer(async(req,res)=>{
  try{
    if(!['GET','HEAD'].includes(req.method)){res.writeHead(405,{'Allow':'GET, HEAD'});res.end();return;}
    const url=new URL(req.url,'http://localhost');let pathname=url.pathname;
    if(basePath){if(pathname===basePath){res.writeHead(308,{'Location':basePath+'/'});res.end();return;}if(!pathname.startsWith(basePath+'/')){res.writeHead(404);res.end();return;}pathname=pathname.slice(basePath.length);}
    const file=path.resolve(root,'.'+decodeURIComponent(pathname==='/'?'/index.html':pathname));
    if(!file.startsWith(root+path.sep)){res.writeHead(403);res.end();return;}
    if(!(await stat(file)).isFile())throw new Error('not file');
    const mime={'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.txt':'text/plain; charset=utf-8'}[path.extname(file)]||'application/octet-stream';
    res.writeHead(200,{'Content-Type':mime,'Cache-Control':'no-store','X-Content-Type-Options':'nosniff'});res.end(req.method==='HEAD'?undefined:await readFile(file));
  }catch{res.writeHead(404);res.end('未找到文件');}
});
server.listen(port,'127.0.0.1',()=>{
  const url=`http://127.0.0.1:${server.address().port}${basePath}/`;
  console.log(`台账对对看：${url} （仅本机，Ctrl+C 停止）`);
  if(process.argv.includes('--open')&&process.platform==='darwin')spawn('open',[url],{stdio:'ignore'}).on('error',()=>console.log('请手动在浏览器打开上方地址。'));
}).on('error',err=>{console.error(err.message);process.exit(1);});
