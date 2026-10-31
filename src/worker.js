import {importBytes,reportBytes} from './io.js';
import {reconcile} from './core.js';
let queue=Promise.resolve();
self.onmessage=event=>{
 const {id,op,args}=event.data;
 queue=queue.then(async()=>{
  try{
   if(op==='report'){
    const bytes=await reportBytes(...args),buffer=bytes instanceof ArrayBuffer?bytes:bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.byteLength);
    self.postMessage({id,result:buffer},[buffer]);
   }else{
    const result=op==='import'?await importBytes(...args):op==='compare'?reconcile(...args):null;
    if(result===null)throw new Error('未知本机计算任务');
    self.postMessage({id,result});
   }
  }catch(error){self.postMessage({id,error:error.message||'本机计算失败'});}
 });
};
self.postMessage({ready:true});
