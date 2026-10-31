import { cellText, UNITS } from './core.js';
import { importBytes, selectTable } from './io.js';
import { demoA, demoB, demoConfig } from './fixtures.js';
import {runTask} from './tasks.js';
const $ = id => document.getElementById(id);
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const freshSource = () => ({file:null, book:null, table:null, encoding:'utf-8', sheetIndex:0, header:1, error:'', serial:0});
const freshConfig = () => ({keys:[{label:'主键',left:null,right:null}], fields:[{label:'状态',type:'text',left:null,right:null},{label:'金额',type:'amount',left:null,right:null,leftUnit:'元',rightUnit:'元'}],tolerance:'0.01'});
let sources = {left:freshSource(),right:freshSource()}, config = freshConfig(), result = null, tab = 'differences', page = 0, query = '',revision=0;
const PAGE_SIZE = 50;
function message(text='', error=false) {$('message').hidden=!text; $('message').textContent=text; $('message').classList.toggle('failure',error);}
function invalidate() {revision++;result=null; $('results').hidden=true;}
const ready = () => sources.left.table && sources.right.table;
function renderSource(side) {
  const s=sources[side], name=side==='left'?'A':'B', sheet=s.book?.sheets[s.sheetIndex];
  let preview='';
  if(sheet) {
    const rows=s.table?[{rowNumber:s.table.headerRow,cells:s.table.columns.map(c=>({text:c.label}))},...s.table.rows.slice(0,5)]:sheet.rows.slice(0,7);
    preview=`<div class="preview-label"><span>${s.table?'所选标题与前 5 条记录':'原始记录预览，请确认标题行'}</span><span>${s.table?`${s.table.rows.length} 条明细`:'前 7 条'}</span></div><div class="table-scroll"><table><tbody>${rows.map((row,ri)=>{const tag=ri===0&&s.table?'th':'td';return `<tr><${tag} class="row-number">${row.rowNumber}</${tag}>${row.cells.slice(0,10).map(c=>{const v=c.formula?'='+c.formula:cellText(c);return `<${tag} title="${esc(v)}">${esc(v)||'—'}</${tag}>`;}).join('')}</tr>`;}).join('')}</tbody></table></div>`;
  }
  $('source-'+side).innerHTML=`<div class="source-head"><span class="side-label ${side==='right'?'b':''}">${name}</span><h3>台账 ${name}</h3><span class="badge">${s.table?'已就绪':'本机读取'}</span></div>
  <div class="upload"><span class="upload-icon" aria-hidden="true">▤</span>${s.file||s.book?`<span class="filename">${esc(s.file?.name||s.book.fileName)}</span>`:'<p>选择一份需要核对的明细台账</p>'}<label class="button file-picker">${s.book?'更换文件':'选择 XLSX / CSV'}<input id="file-${side}" type="file" accept=".xlsx,.csv" data-action="file" data-side="${side}" aria-label="选择台账 ${name} 文件"></label><p>仅读取，不修改文件 · 最大 20 MB</p></div>
  <label class="encoding">CSV 编码<select data-action="encoding" data-side="${side}" aria-label="台账 ${name} CSV 编码"><option value="utf-8" ${s.encoding==='utf-8'?'selected':''}>UTF-8</option><option value="gb18030" ${s.encoding==='gb18030'?'selected':''}>GB18030</option></select></label>
  <div class="file-config"><label>工作表<select data-action="sheet" data-side="${side}" aria-label="台账 ${name} 工作表" ${!s.book?'disabled':''}>${s.book?s.book.sheets.map((x,i)=>`<option value="${i}" ${s.sheetIndex===i?'selected':''}>${esc(x.name)}</option>`).join(''):'<option>请先选择文件</option>'}</select></label><label>标题${s.book?.format==='csv'?'记录序号':'行'}<input type="number" min="1" step="1" max="${sheet?.rows.length||1}" value="${s.header}" data-action="header" data-side="${side}" aria-label="台账 ${name} 标题行" ${!s.book?'disabled':''}></label></div>
  ${s.book?.format==='csv'?'<div class="muted">CSV 标题按记录序号选择；预览为起始物理行号。</div>':''}${s.error?`<div class="error" role="alert">${esc(s.error)}</div>`:''}${preview}${(s.table?.warnings||[]).map(x=>`<div class="warn">${esc(x)}</div>`).join('')}`;
}
function columnSelect(side,group,i,selected) {
  return `<select data-group="${group}" data-index="${i}" data-prop="${side}" aria-label="${group==='keys'?'主键':'核对字段'} ${i+1} ${side==='left'?'A':'B'} 列"><option value="">请选择列</option>${(sources[side].table?.columns||[]).map(c=>`<option value="${c.index}" ${selected===c.index?'selected':''}>${c.index+1}. ${esc(c.label)}</option>`).join('')}</select>`;
}
function unitSelect(side,i,value) {return `<div class="unit"><label>原表单位</label><select data-group="fields" data-index="${i}" data-prop="${side}Unit" aria-label="核对字段 ${i+1} ${side==='left'?'A':'B'} 金额单位">${Object.keys(UNITS).map(u=>`<option ${u===value?'selected':''}>${u}</option>`).join('')}</select></div>`;}
function renderMappings() {
  $('mapping').disabled=!ready(); $('rules-hint').textContent=ready()?'仅核对明确选定的字段':'导入两份有效台账后开始配置';
  for(const group of ['keys','fields']) $(group).innerHTML=config[group].map((f,i)=>`<div class="mapping-row ${group==='fields'?'field':''}"><label>字段名称<input type="text" maxlength="80" data-group="${group}" data-index="${i}" data-prop="label" value="${esc(f.label)}" aria-label="${group==='keys'?'主键':'核对字段'} ${i+1} 名称"></label>${group==='fields'?`<label>核对类型<select data-group="fields" data-index="${i}" data-prop="type" aria-label="核对字段 ${i+1} 类型"><option value="text" ${f.type==='text'?'selected':''}>状态 / 文本</option><option value="amount" ${f.type==='amount'?'selected':''}>金额</option></select></label>`:''}<div><label>A 表字段${columnSelect('left',group,i,f.left)}</label>${f.type==='amount'?unitSelect('left',i,f.leftUnit):''}</div><span class="arrow">↔</span><div><label>B 表字段${columnSelect('right',group,i,f.right)}</label>${f.type==='amount'?unitSelect('right',i,f.rightUnit):''}</div><button class="delete" data-delete="${group}" data-index="${i}" aria-label="删除${group==='keys'?'主键':'核对字段'} ${i+1}">×</button></div>`).join('');
  $('tolerance').value=config.tolerance;
}
function applySource(side) {
  const s=sources[side]; s.table=null; s.error='';
  try{s.table=selectTable(s.book,s.sheetIndex,s.header);}catch(e){s.error=e.message;}
  invalidate(); config=freshConfig(); renderSource(side); renderMappings();
}
async function loadFile(side,file) {
  if(!file)return; const s=sources[side],serial=++s.serial;
  Object.assign(s,{file,book:null,table:null,error:'',header:1,sheetIndex:0}); invalidate(); config=freshConfig(); renderSource(side); renderMappings(); message(`正在本机读取台账 ${side==='left'?'A':'B'}…`);
  try{
    if(file.size>20*1024*1024)throw new Error('文件超过 20 MB，请先拆分后再导入');
    const book=await runTask('import',await file.arrayBuffer(),file.name,s.encoding);
    if(s!==sources[side]||serial!==s.serial)return;
    s.book=book; applySource(side); message(s.error?'请根据预览调整标题行，再继续配置。':'读取完成，请确认工作表和标题行。');
  }catch(e){if(s!==sources[side]||serial!==s.serial)return;s.error=e.message;renderSource(side);message('导入失败，本次核对未执行。请处理文件后重新导入。',true);}
}
function renderResults() {
  if(!result)return;const s=result.summary;$('results').hidden=false;
  $('results').innerHTML=`<div class="result-head"><div><h2>03 <span>核对结果</span></h2><div class="result-title ${result.allConsistent?'success':'attention'}">${result.allConsistent?'所选字段全部一致':s.issueCount?'存在差异或待处理记录':'发现台账差异'}</div><p class="result-desc">${s.issueCount?'待处理记录尚未完成核对，不能确认全部一致。':'仅基于本次选择的主键与字段。'} A 表 ${s.left} 条 / B 表 ${s.right} 条 · 容差 ${esc(result.config.tolerance)} 元</p></div><button id="export" class="button primary">导出 XLSX 报告 ↓</button></div>
  <div class="stats">${[['已配对',s.matched],['配对一致',s.agreed],['有差异配对',s.changed],['仅 A 有',s.missingRight],['仅 B 有',s.missingLeft],['待处理源记录',s.invalidRows]].map(([label,value],i)=>`<div class="stat ${i>1&&value?'attention':''}"><span>${label}</span><strong>${value}</strong></div>`).join('')}</div>
  <div class="card result-card"><div class="result-controls"><div class="tabs" role="tablist" aria-label="异常分类"><button class="tab ${tab==='differences'?'active':''}" role="tab" aria-selected="${tab==='differences'}" data-tab="differences">差异明细 ${s.differenceCount}</button><button class="tab ${tab==='issues'?'active':''}" role="tab" aria-selected="${tab==='issues'}" data-tab="issues">重复 / 无效记录 ${s.issueCount}</button></div><input id="search-results" type="text" placeholder="搜索主键、字段、原值或行号" aria-label="搜索异常记录" value="${esc(query)}"></div><div id="result-table"></div></div><div class="scope-note">金额差 = A − B，统一为元；空白显示为「（空白）」。页面分页展示，导出包含全部异常及配置。CSV 行号为记录起始物理行。</div>`;
  renderResultTable();
}
function renderResultTable() {
  if(!result)return;
  const filtered=result[tab].filter(x=>!query||Object.values(x).some(v=>String(v).includes(query))),pages=Math.max(1,Math.ceil(filtered.length/PAGE_SIZE));page=Math.min(page,pages-1);
  const rows=filtered.slice(page*PAGE_SIZE,(page+1)*PAGE_SIZE);
  const sourceCell=(x,side)=>`<span class="cell-value">${esc(x[side+'Value'])||'（空白）'}</span><div class="source-ref">${x[side+'Row']?`${esc(x[side+'File'])} / ${esc(x[side+'Sheet'])} / 第 ${x[side+'Row']} 行`:'该侧无关联记录'}</div>`;
  $('result-table').innerHTML=rows.length?`<div class="table-scroll"><table><thead><tr><th>异常类型</th><th>主键 / 字段</th><th>A 原值与来源</th><th>B 原值与来源</th><th>金额差（元）</th><th>说明</th></tr></thead><tbody>${rows.map(x=>`<tr><td>${esc(x.type)}</td><td>${esc(x.key)}<div class="source-ref">${esc(x.field)}</div></td><td>${sourceCell(x,'left')}</td><td>${sourceCell(x,'right')}</td><td>${esc(x.delta)||'—'}</td><td>${esc(x.note)}</td></tr>`).join('')}</tbody></table></div><div class="pagination"><span>共 ${filtered.length} 条 · 第 ${page+1} / ${pages} 页 · 每页 ${PAGE_SIZE} 条</span><div><button class="button small" data-page="-1" ${page===0?'disabled':''}>上一页</button><button class="button small" data-page="1" ${page>=pages-1?'disabled':''}>下一页</button></div></div>`:`<div class="empty">${query?'没有符合搜索条件的异常。':tab==='issues'?'没有重复或无效记录。':'所选字段没有发现差异。'}</div>`;
}
document.addEventListener('change',async e=>{
  const t=e.target;
  if(t.dataset.action){
    const side=t.dataset.side,s=sources[side];
    if(t.dataset.action==='file')await loadFile(side,t.files[0]);
    if(t.dataset.action==='encoding'){s.encoding=t.value;if(s.file)await loadFile(side,s.file);}
    if(t.dataset.action==='sheet'){s.sheetIndex=Number(t.value);s.header=1;applySource(side);message('工作表已改变，请重新确认标题和映射。');}
    if(t.dataset.action==='header'){s.header=Number(t.value);applySource(side);message('标题行已改变，请重新配置字段映射。');}
  }
  if(t.dataset.group){
    const f=config[t.dataset.group][Number(t.dataset.index)],p=t.dataset.prop;
    f[p]=['left','right'].includes(p)?t.value===''?null:Number(t.value):t.value;
    invalidate();message('规则已更新，请重新核对。');
    if(p==='type'){f.leftUnit||='元';f.rightUnit||='元';renderMappings();}
  }
});
document.addEventListener('input',e=>{
  const t=e.target;
  if(t.id==='tolerance'){config.tolerance=t.value;invalidate();}
  if(t.dataset.prop==='label'){config[t.dataset.group][Number(t.dataset.index)].label=t.value;invalidate();}
  if(t.id==='search-results'){query=t.value;page=0;renderResultTable();}
});
document.addEventListener('click',async e=>{
  const t=e.target.closest('button');if(!t)return;
  if(t.id==='demo'){
    t.disabled=true;message('正在载入合成示例…');
    try{
      sources={left:freshSource(),right:freshSource()};invalidate();
      for(const [side,csv,name,header] of [['left',demoA,'合成示例_A.csv',2],['right',demoB,'合成示例_B.csv',1]]){
        const s=sources[side];s.book=await importBytes(new TextEncoder().encode(csv),name);s.header=header;s.table=selectTable(s.book,0,header);renderSource(side);
      }
      config=structuredClone(demoConfig);renderMappings();message('已载入纯合成示例，并填写演示映射。A 金额为元，B 为万元。点击“开始核对”体验异常追溯。');
    }catch(err){message(err.message,true);}finally{t.disabled=false;}
  }
  if(t.id==='reset'){sources={left:freshSource(),right:freshSource()};config=freshConfig();invalidate();renderSource('left');renderSource('right');renderMappings();message('本次数据已从页面移除。');}
  if(t.id==='add-key'||t.id==='add-field'){
    const group=t.id==='add-key'?'keys':'fields';config[group].push({label:'',left:null,right:null,...(group==='fields'?{type:'text',leftUnit:'元',rightUnit:'元'}:{})});invalidate();renderMappings();
  }
  if(t.dataset.delete){config[t.dataset.delete].splice(Number(t.dataset.index),1);invalidate();renderMappings();}
  if(t.id==='run'){
    t.disabled=true;t.textContent='正在核对…';message();await new Promise(resolve=>requestAnimationFrame(()=>setTimeout(resolve,10)));
    const token=revision;
    try{const next=await runTask('compare',sources.left.table,sources.right.table,config);if(token!==revision)return;result=next;tab='differences';page=0;query='';renderResults();$('results').scrollIntoView({behavior:'smooth',block:'start'});}
    catch(err){invalidate();message(err.message,true);}finally{t.disabled=false;t.innerHTML='开始核对 <span>→</span>';}
  }
  if(t.dataset.tab){tab=t.dataset.tab;page=0;renderResults();}
  if(t.dataset.page){page+=Number(t.dataset.page);renderResultTable();}
  if(t.id==='export'&&result){
    t.disabled=true;t.textContent='正在生成报告…';const snapshot=result;
    try{
      const bytes=await runTask('report',snapshot);if(snapshot!==result){message('设置已改变，本次报告已取消，请重新核对。');return;}const url=URL.createObjectURL(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
      const a=document.createElement('a');a.href=url;a.download=`台账核对报告_${new Date().toISOString().replace(/[:.]/g,'-')}.xlsx`;a.click();setTimeout(()=>URL.revokeObjectURL(url),30000);message('独立 XLSX 报告已生成，请在浏览器下载列表或保存位置查看。');
    }catch(err){message('报告生成失败：'+err.message,true);}finally{t.disabled=false;t.textContent='导出 XLSX 报告 ↓';}
  }
});
renderSource('left');renderSource('right');renderMappings();
