// Eager local worker: computation stays on this computer, including offline use.
const pending=new Map();let nextId=0,worker,failed;
let readyResolve,readyReject;
const ready=new Promise((resolve,reject)=>{readyResolve=resolve;readyReject=reject;});
ready.catch(()=>{});
try{
 worker=new Worker(new URL('worker.js',document.currentScript?.src||document.baseURI));
 worker.onmessage=event=>{
  if(event.data.ready){readyResolve();return;}
  const item=pending.get(event.data.id);if(!item)return;pending.delete(event.data.id);
  if(event.data.error)item.reject(new Error(event.data.error));else item.resolve(event.data.result);
 };
 worker.onerror=()=>{
  failed=new Error('无法启动本机计算线程，请用近期 Chrome 经本机 HTTP 入口打开；请勿直接打开 file: 文件');
  readyReject(failed);for(const item of pending.values())item.reject(failed);pending.clear();
 };
}catch{failed=new Error('浏览器不支持本机计算线程，请使用近期 Chrome 和本机 HTTP 入口');readyReject(failed);}
export async function runTask(op,...args){
 await ready;if(failed)throw failed;
 return new Promise((resolve,reject)=>{const id=++nextId;pending.set(id,{resolve,reject});try{worker.postMessage({id,op,args});}catch(error){pending.delete(id);reject(error);}});
}
