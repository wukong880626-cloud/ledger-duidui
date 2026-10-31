import {mkdir,writeFile,readFile} from 'node:fs/promises';
import {execFileSync} from 'node:child_process';
const out='output/round2/fixtures',sample='samples/第二轮';await mkdir(sample,{recursive:true});
const encodeRow=row=>row.map(v=>/[",\r\n]/.test(v)?'"'+v.replaceAll('"','""')+'"':v).join(',');
const values=['_x0041_','_x005F_x0041_','=1+1','@SUM(1)','<img src=x onerror=alert(1)>'];
for(const side of ['A','B'])await writeFile(`${sample}/文本保真_${side}.csv`,'\uFEFF编号,状态,金额\n'+values.map((v,i)=>encodeRow([String(i+1).padStart(4,'0'),side==='A'?v:'普通','100'])).join('\n'));
await writeFile(`${sample}/超精度数值.xlsx`,await readFile(`${out}/超精度数值.xlsx`));
await writeFile(`${sample}/精度对侧.csv`,'\uFEFF编号,状态,金额\n001,签约,1');
await writeFile(`${sample}/控制字符拦截.csv`,'\uFEFF编号,状态,金额\n001,含\u0001控制字符,1');
const gb='编号,状态,金额\n0001,签约,100\n0002,认购,200';
const gbBytes=execFileSync('python3',['-c','import sys;sys.stdout.buffer.write(sys.stdin.read().encode("gb18030"))'],{input:gb});await writeFile(`${sample}/GB18030编码.csv`,gbBytes);await writeFile(`${sample}/UTF8编码.csv`,'\uFEFF'+gb);
for(const side of ['A','B']){
 await writeFile(`${out}/取消_${side}.csv`,'编号,状态,金额\n'+Array.from({length:25000},(_,i)=>`${String(i).padStart(6,'0')},${side==='A'?'签约':'认购'},1`).join('\n'));
 await writeFile(`${out}/异常超限_${side}.csv`,'编号,状态,金额\n'+Array.from({length:12501},(_,i)=>`${i},${side==='A'?'签约':'认购'},${side==='A'?1:2}`).join('\n'));
}
await writeFile(`${out}/行数超限.csv`,'编号,状态,金额\n'+Array.from({length:25001},(_,i)=>`${i},签约,1`).join('\n'));
console.log('已生成第二轮功能合成样本；大表仅保存在 output/round2/fixtures。');
