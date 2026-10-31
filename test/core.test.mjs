import {test} from 'node:test';
import assert from 'node:assert/strict';
import {reconcile,parseAmount} from '../src/core.js';
const cell=raw=>({raw,text:raw==null?'':String(raw)});
const table=(rows,name='A.csv')=>({fileName:name,sheetName:'明细',headerRow:2,columns:[{label:'编号'},{label:'状态'},{label:'金额'},{label:'单元'}],rows:rows.map((r,i)=>({rowNumber:i+3,cells:r.map(v=>typeof v==='object'&&v!==null?v:cell(v))}))});
const cfg=()=>({keys:[{label:'编号',left:0,right:0}],fields:[{label:'状态',type:'text',left:1,right:1},{label:'金额',type:'amount',left:2,right:2,leftUnit:'元',rightUnit:'元'}],tolerance:'0.01'});
const run=(a,b,c=cfg())=>reconcile(table(a),table(b,'B.xlsx'),c);
test('行序不同仍一致，编号前导零按文本保留',()=>{
 const r=run([['001','认购','100'],['002','已签约','200']],[['002','已签约','200'],['001','认购','100']]);assert.equal(r.allConsistent,true);assert.equal(r.summary.agreed,2);
 assert.equal(run([['001','认购','100']],[['1','认购','100']]).summary.matched,0);
});
test('手动映射列序和名称，金额单位转换为元',()=>{
 const c=cfg();c.keys[0].right=2;c.fields[1].right=0;c.fields[1].rightUnit='万元';assert.equal(run([['001','签约','12345']],[['1.2345','签约','001']],c).allConsistent,true);
});
test('复合键逐分量匹配，避免分隔符碰撞',()=>{
 const c=cfg();c.keys.push({label:'单元',left:3,right:3});
 assert.equal(run([['01','签约','1','001'],['01','签约','2','002']],[['01','签约','2','002'],['01','签约','1','001']],c).allConsistent,true);
 assert.equal(run([['a|b','签约','1','c']],[['a','签约','1','b|c']],c).summary.matched,0);
});
test('空键和缺复合键分量隔离，不宣称全部一致',()=>{
 const c=cfg();c.keys.push({label:'单元',left:3,right:3});const r=run([['','','1',''],['01','签约','1','']],[['02','签约','1','1']],c);
 assert.deepEqual(r.issues.map(x=>x.type),['空键','复合键不完整']);assert.equal(r.allConsistent,false);assert.equal(r.issues[1].leftRow,4);
});
test('单侧重复键隔离两边，不误报缺项',()=>{
 const r=run([['001','签约','1'],['001','签约','1']],[['001','签约','1']]);assert.equal(r.summary.matched,0);assert.equal(r.summary.duplicateRows,2);assert.equal(r.differences.length,0);assert.equal(r.issues.length,3);assert.equal(r.allConsistent,false);
});
test('两侧均重复时逐条追溯',()=>{
 const r=run([['001','签约','1'],['001','签约','2']],[['001','签约','1'],['001','签约','2']]);assert.equal(r.summary.duplicateRows,4);assert.equal(r.issues.length,4);
});
test('状态、缺项、两边原值与来源行号完整',()=>{
 const r=run([['001','认购','10'],['002','认购','20']],[['003','认购','30'],['001','签约','11']]);assert.equal(r.summary.changed,1);assert.equal(r.summary.missingLeft,1);assert.equal(r.summary.missingRight,1);
 const d=r.differences.find(x=>x.type==='金额不同');assert.equal(d.delta,'-1');assert.equal(d.leftRow,3);assert.equal(d.rightRow,4);assert.equal(d.leftFile,'A.csv');assert.equal(d.rightSheet,'明细');assert.equal(d.leftValue,'10');assert.equal(d.rightValue,'11');
});
test('容差下方、等值、上方的十进制边界',()=>{
 const r=run([['1','签约','100.0099'],['2','签约','100.01'],['3','签约','100.0101']],[['1','签约','100'],['2','签约','100'],['3','签约','100']]);assert.equal(r.summary.agreed,2);assert.equal(r.differences.length,1);assert.equal(r.differences[0].delta,'0.0101');
});
test('空白不等于零，两边空白一致',()=>{
 const r=run([['1','签约',''],['2','签约',null],['3','签约','0']],[['1','签约',0],['2','签约',''],['3','签约','0']]);assert.equal(r.summary.agreed,2);assert.equal(r.differences[0].type,'金额空值不同');assert.equal(r.differences[0].delta,'');
});
test('非法金额单列并隔离对侧，不当作零',()=>{
 const r=run([['1','签约','待核'],['2','签约','12,34']],[['1','签约','0'],['2','签约','1234']]);assert.equal(r.issues.filter(x=>x.type==='非法金额').length,2);assert.equal(r.summary.matched,0);assert.equal(r.differences.length,0);assert.equal(r.allConsistent,false);
});
test('正负金额、规范千分位、四种单位与非法格式',()=>{
 assert.equal(parseAmount(cell('-1,234.50'),'万元').value.toFixed(),'-12345000');assert.equal(parseAmount(cell('1'),'亿元').value.toFixed(),'100000000');assert.equal(parseAmount(cell('1'),'千元').value.toFixed(),'1000');
 for(const s of ['1e3','￥100','Infinity','NaN','1 000','(100)','--1','=1+1'])assert.ok(parseAmount(cell(s)).error,s);
 assert.ok(parseAmount(cell('1'),'美元').error);assert.ok(parseAmount(cell(true)).error);assert.ok(parseAmount({...cell('2026-01-01'),kind:'date'}).error);
});
test('主键与核对字段的有缓存、无缓存公式都拦截',()=>{
 for(const raw of ['',100]){const r=run([['1','签约',{raw,text:String(raw),formula:'50+50'}]],[['1','签约',100]]);assert.ok(r.issues.some(x=>x.type==='关键字段含公式'));assert.equal(r.allConsistent,false);assert.equal(r.summary.matched,0);}
 const r=run([[{raw:'',text:'',formula:'"001"'},'签约',100]],[['001','签约',100]]);assert.equal(r.issues[0].type,'关键字段含公式');assert.match(r.issues[0].note,/转为值/);
});
test('错误值、非整数与不安全数字编号拒绝匹配',()=>{
 for(const value of [{raw:'#N/A',text:'#N/A',error:'#N/A'},{raw:9007199254740992,text:'9007199254740992'},cell(1.5)]){const r=run([[value,'签约',1]],[['other','签约',1]]);assert.equal(r.issues[0].type,'无效主键');assert.equal(r.allConsistent,false);}
});
test('空格和大小写精确比较',()=>{
 assert.equal(run([['A ','签约',1]],[['A','签约',1]]).summary.matched,0);assert.equal(run([['A','签约 ',1]],[['A','签约',1]]).differences[0].type,'状态/文本不同');assert.equal(run([['a','签约',1]],[['A','签约',1]]).summary.matched,0);
});
test('无字段、重复或不完整映射、非法容差拒绝执行',()=>{
 for(const mutate of [c=>c.keys=[],c=>c.fields=[],c=>c.keys[0].left=null,c=>c.fields[1].left=1,c=>c.tolerance='-0.01',c=>c.tolerance='',c=>c.tolerance='abc']){const c=cfg();mutate(c);assert.throws(()=>run([['1','签约',1]],[['1','签约',1]],c));}
});
test('空数据不报告全部一致',()=>{assert.equal(run([],[]).allConsistent,false);});
test('同一行缺主键且金额非法时保留两类问题，源记录只计一次',()=>{
 const r=run([['','签约','非法']],[['001','签约','1']]);
 assert.deepEqual(new Set(r.issues.map(x=>x.type)),new Set(['非法金额','空键']));assert.equal(r.summary.issueCount,2);assert.equal(r.summary.invalidRows,1);assert.equal(r.allConsistent,false);
});
