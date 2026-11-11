import './style.css';
import {combineSources,expandRows,selectLocations,availabilityWarning,validPrintDate,localDateString,type SourceData,type Selection} from './model';
import type {PrintOptions} from './excel';
interface Upload {id:number;name:string;size?:number;data?:SourceData;error?:string}
const app=document.querySelector<HTMLDivElement>('#app')!;
app.innerHTML=`
 <header class="topbar"><div class="brand"><span class="brand-icon">R</span><span>LOCATION <b>STUDIO</b></span></div><span class="private"><span class="dot"></span> 本机处理 · 无需登录</span></header>
 <main><div class="intro"><div><p class="eyebrow">WAREHOUSE OPERATIONS / 库位作业</p><h1>把库存导出，变成现场作业单。</h1><p class="subtitle">多个单库位文件，或一个多库位文件。一次整理，每库位一张纸。</p></div><button id="demo" class="secondary">查看演示数据 <span>↗</span></button></div>
 <div class="workspace"><section class="controls"><div class="card"><div class="step-title"><span class="step">01</span><h2>导入库存</h2><span id="mode-badge" class="badge">XLSX</span></div>
 <label id="dropzone" class="dropzone" for="files"><span class="upload-icon">↥</span><strong>选择或拖入 Excel 文件</strong><span>支持同时选择多个文件，每个文件 ≤ 10 MB</span><span class="file-button">选择 .xlsx 文件</span></label><input id="files" type="file" accept=".xlsx" multiple class="visually-hidden"><div id="file-list" class="file-list"></div><p class="help">读取文件内部的库位编号，不依赖文件名。原文件保持不变。</p></div>
 <div class="card"><div class="step-title"><span class="step">02</span><h2>设置作业单</h2></div>
 <fieldset class="scope"><legend>库位范围</legend><label><input type="radio" name="scope" value="all" checked> 文件中的全部 REI 库位</label><label><input type="radio" name="scope" value="range"> 指定连续范围</label><div class="range-row"><label><span>起始编号</span><div class="number-box"><span>REI-</span><input id="from" type="number" min="1" max="100" value="10" aria-label="起始库位编号" disabled></div></label><span class="range-dash">—</span><label><span>结束编号</span><div class="number-box"><span>REI-</span><input id="to" type="number" min="1" max="100" value="20" aria-label="结束库位编号" disabled></div></label></div></fieldset>
 <div class="field-grid"><div class="form-label date-field"><label for="date">打印日期</label><div class="date-controls"><input id="date" type="date" aria-label="打印日期" required><button id="pick-date" type="button" class="date-button" aria-label="选择打印日期">选日期</button><button id="today-date" type="button" class="date-button">今天</button></div><small class="date-hint">用于所有标签页的 Datum 和下载文件名</small></div><label class="form-label">作业人 <span class="muted">可留空</span><input id="name" maxlength="80" placeholder="现场填写，或预先输入"></label></div>
 <label id="missing-option" class="missing-option" hidden><input id="allow-missing" type="checkbox"> 为缺失库位生成提示页（不代表无货）</label>
 <div class="print-rule"><span>▤</span><div><strong>A4 纵向 · 每库位一页</strong><p>按库存数量展开，一件一行。Location 留空供现场填写。</p></div></div>
 <button id="download" class="primary" disabled>生成并下载 Excel <span>↓</span></button><p id="download-status" role="status" class="download-status">先导入文件，或查看演示数据</p></div></section>
 <section class="results"><div class="summary"><div><span>库位 / 标签页</span><strong id="total-locations">—</strong></div><div><span>作业件数</span><strong id="total-pieces">—</strong></div><div><span>有数据的库位</span><strong id="available-locations">—</strong></div></div>
 <div id="issues" role="alert" aria-live="polite"></div>
 <div class="preview-card"><div class="preview-heading"><div><p class="eyebrow">WORKSHEET PREVIEW</p><h2>作业单预览</h2></div><span class="badge">A4 / 1 PAGE</span></div><div id="tabs" class="tabs" role="tablist" aria-label="库位标签页"></div><div id="empty" class="empty"><span>▤</span><h3>作业单，从这里开始</h3><p>导入库存文件后，按 REI 编号自动分组。<br>每张作业单都可以预览，再统一导出。</p><button id="empty-demo" class="text-button">用演示数据试一下 →</button></div>
 <div id="paper-wrap" class="paper-wrap" hidden><div id="paper" class="paper"></div></div><div id="sheet-footer" class="sheet-footer" hidden></div></div>
 </section></div><footer>REI-1 — REI-100 <span>·</span> 文件在浏览器本地处理 <span>·</span> 输出格式沿用 RT 作业模板</footer></main>`;
const $=<T extends HTMLElement=HTMLElement>(id:string)=>document.getElementById(id) as T;
let uploads:Upload[]=[],selection:Selection|null=null,activeLocation='',busy=false,serial=0,demo=false;
$('date').setAttribute('value',localDateString());
function node<K extends keyof HTMLElementTagNameMap>(tag:K,text='',className=''):HTMLElementTagNameMap[K]{const el=document.createElement(tag);el.textContent=text;if(className)el.className=className;return el;}
const options=():PrintOptions=>({date:$<HTMLInputElement>('date').value,name:$<HTMLInputElement>('name').value.trim(),allowMissing:$<HTMLInputElement>('allow-missing').checked});
function notice(text:string,error=false){return node('p',text,error?'notice error':'notice');}
function setBusy(value:boolean){busy=value;document.querySelectorAll<HTMLInputElement>('input').forEach(input=>input.disabled=value);$<HTMLInputElement>('files').disabled=value;$<HTMLButtonElement>('demo').disabled=value;$<HTMLButtonElement>('pick-date').disabled=value;$<HTMLButtonElement>('today-date').disabled=value;$<HTMLButtonElement>('empty-demo').disabled=value;document.querySelectorAll<HTMLButtonElement>('.remove-file').forEach(button=>button.disabled=value);update();}
function update(){
 const issue=$('issues');issue.replaceChildren();selection=null;
 const good=uploads.flatMap(upload=>upload.data?[upload.data]:[]),bad=uploads.filter(upload=>upload.error);
 for(const upload of bad)issue.append(notice(upload.error!,true));
 try{
  const rows=combineSources(good);
  const ranged=document.querySelector<HTMLInputElement>('input[name="scope"]:checked')?.value==='range';
  $<HTMLInputElement>('from').disabled=!ranged||busy;$<HTMLInputElement>('to').disabled=!ranged||busy;
  if(rows.length)selection=selectLocations(rows,ranged?{start:Number($<HTMLInputElement>('from').value),end:Number($<HTMLInputElement>('to').value)}:undefined);
 }catch(error){issue.append(notice(error instanceof Error?error.message:String(error),true));}
 if(demo)issue.prepend(notice('当前是演示数据，用于查看效果，不代表真实库存。'));
 const ignored=good.reduce((sum,source)=>sum+source.ignored,0);if(ignored)issue.append(notice(`已排除 ${ignored} 条非 REI-1～REI-100 的记录。`));
 for(const warning of [...new Set(good.flatMap(source=>source.warnings))].slice(0,15))issue.append(notice(warning));
 const countWarnings=[...new Set(good.flatMap(source=>source.warnings))].length;if(countWarnings>15)issue.append(notice(`另有 ${countWarnings-15} 条编码提示，请检查源文件。`));
 if(selection?.missing.length)issue.append(notice(`缺少数据：${selection.missing.join('、')}。缺少数据不能判断为无货。`));
 for(const warning of selection?.warnings||[])issue.append(notice(warning));
 const availability=selection?.groups.flatMap(group=>group.rows.map(availabilityWarning).filter((warning):warning is string=>!!warning))||[];
 if(availability.length){const details=node('details','','notice availability-warnings');details.open=true;details.append(node('summary',`可用库存预警 · ${availability.length} 条（不影响打印）`));const list=node('ul');availability.forEach(warning=>list.append(node('li',warning)));details.append(list);issue.append(details);}
 if(selection&&!validPrintDate(options().date))issue.append(notice('请选择有效的打印日期。',true));
 $('missing-option').hidden=!selection?.missing.length;
 $('total-locations').textContent=selection?String(selection.groups.length):'—';$('total-pieces').textContent=selection?String(selection.pieces):'—';$('available-locations').textContent=selection?String(selection.groups.length-selection.missing.length):'—';
 const valid=selection&&selection.groups.some(group=>!group.missing)&&!bad.length&&validPrintDate(options().date)&&(!selection.missing.length||options().allowMissing);
 $<HTMLButtonElement>('download').disabled=busy||!valid;
 $('download-status').textContent=busy?'正在处理，请稍候…':valid?`将生成 ${selection!.groups.length} 个标签页，按库位数字顺序排列`:selection?.missing.length?'补充缺失数据，或勾选生成缺失提示页':'先导入文件，或查看演示数据';
 $('mode-badge').textContent=demo?'演示':uploads.length?`${uploads.length} 个文件`:'XLSX';
 const list=$('file-list');list.replaceChildren();for(const upload of uploads){const row=node('div','','file-item'),meta=node('div');meta.append(node('strong',upload.name),node('span',upload.error?'读取失败':`${upload.data?.rows.length||0} 条 REI 记录`));const remove=node('button','×','remove-file');remove.type='button';remove.disabled=busy;remove.setAttribute('aria-label',`移除 ${upload.name}`);remove.onclick=()=>{uploads=uploads.filter(file=>file.id!==upload.id);update();};row.append(meta,remove);list.append(row);}
 const tabs=$('tabs');tabs.replaceChildren();if(!selection?.groups.some(group=>group.location===activeLocation))activeLocation=selection?.groups[0]?.location||'';
 for(const group of selection?.groups||[]){const tab=node('button',`${group.location}${group.missing?' · 缺失':''}`,group.location===activeLocation?'tab active':'tab');tab.setAttribute('role','tab');tab.setAttribute('aria-selected',String(group.location===activeLocation));tab.onclick=()=>{activeLocation=group.location;update();};tabs.append(tab);}
 renderPaper();
}
function renderPaper(){
 const group=selection?.groups.find(group=>group.location===activeLocation);$('empty').hidden=!!group;$('paper-wrap').hidden=!group;$('sheet-footer').hidden=!group;if(!group)return;
 const paper=$('paper');paper.replaceChildren();paper.dataset.location=group.location;
 const header=node('table','','paper-header');const meta=[['STARTZEIT','','ENDZEIT',''],['Name',options().name,'Datum',options().date.split('-').reverse().join('.')]];
 for(const values of meta){const tr=node('tr');values.forEach(value=>tr.append(node('td',value)));header.append(tr);}paper.append(header,node('h3',group.location));
 const table=node('table','','stock-table'),thead=node('thead'),head=node('tr');['SKU','Barcode','Rei','Location'].forEach(label=>head.append(node('th',label)));thead.append(head);table.append(thead);const tbody=node('tbody');
 const expanded=expandRows(group),shown=expanded.slice(0,80);
 for(const values of shown){const tr=node('tr');values.forEach(value=>tr.append(node('td',value)));tbody.append(tr);}
 if(!shown.length){const tr=node('tr'),td=node('td',group.missing?'Keine Quelldaten / 未提供该库位数据':'Quellbestand 0 / 源表库存为 0','missing-cell');td.colSpan=4;tr.append(td);tbody.append(tr);}
 for(let i=Math.max(shown.length,1);i<18;i++){const tr=node('tr');for(let c=0;c<4;c++)tr.append(node('td'));tbody.append(tr);}table.append(tbody);paper.append(table);
 if(expanded.length>shown.length)paper.append(node('p',`预览仅展示前 ${shown.length} 行；Excel 将保留全部 ${expanded.length} 行。`,'preview-limit'));
 $('sheet-footer').textContent=`${group.location} · ${group.missing?'缺少来源数据':`${group.rows.length} 条库存记录 → ${group.pieces} 行货物`} · Excel 中每张工作表设置为一页`;
}
async function importFiles(files:File[]){
 if(busy||!files.length)return;
 if(files.length+(demo?0:uploads.length)>100||files.reduce((sum,file)=>sum+file.size,0)+(demo?0:uploads.reduce((sum,file)=>sum+(file.size||0),0))>50*1024*1024){$('issues').replaceChildren(notice('每次最多选择 100 个文件，总大小不超过 50 MB。',true));return;}
 if(demo){uploads=[];demo=false;}
 setBusy(true);
 try{const {readSource}=await import('./excel');for(const file of files){const upload:Upload={id:++serial,name:file.name,size:file.size};try{upload.data=await readSource(file.name,await file.arrayBuffer());}catch(error){upload.error=error instanceof Error?error.message:String(error);}uploads.push(upload);}}
 catch(error){uploads.push({id:++serial,name:'读取组件',error:`读取失败：${error instanceof Error?error.message:String(error)}`});}
 finally{setBusy(false);}
}
$<HTMLInputElement>('files').onchange=event=>{const input=event.target as HTMLInputElement;void importFiles(Array.from(input.files||[]));input.value='';};
const zone=$('dropzone');zone.ondragover=event=>{event.preventDefault();zone.classList.add('dragover');};zone.ondragleave=()=>zone.classList.remove('dragover');zone.ondrop=event=>{event.preventDefault();zone.classList.remove('dragover');void importFiles(Array.from(event.dataTransfer?.files||[]));};
function loadDemo(){if(busy)return;const rows=Array.from({length:11},(_,i)=>[{location:`REI-${i+10}`,sku:`DEMO-SKU-${i+1}-A`,barcode:`DEMO-${String(i+1).padStart(5,'0')}`,qty:i%3+1,hall:'H1',customer:'DEMO',source:'演示数据',availableQty:i===0?0:i%3+1,sourceRow:i*2+2},{location:`REI-${i+10}`,sku:`DEMO-SKU-${i+1}-B`,barcode:`DEMO-B${String(i+1).padStart(4,'0')}`,qty:1,hall:'H1',customer:'DEMO',source:'演示数据',availableQty:1,sourceRow:i*2+3}]).flat();uploads=[{id:++serial,name:'演示 · REI-10 至 REI-20',data:{name:'演示数据',rows,ignored:0,warnings:[]}}];demo=true;activeLocation='';update();}
$('pick-date').onclick=()=>{if(busy)return;const date=$<HTMLInputElement>('date');try{if(typeof date.showPicker==='function')date.showPicker();else date.focus();}catch{date.focus();}};
$('today-date').onclick=()=>{if(busy)return;$<HTMLInputElement>('date').value=localDateString();update();};
$('demo').onclick=loadDemo;$('empty-demo').onclick=loadDemo;
document.querySelectorAll<HTMLInputElement>('input[name="scope"],#from,#to,#date,#name,#allow-missing').forEach(input=>input.addEventListener('input',update));
$('download').onclick=async()=>{
 if(busy||!selection)return;const chosen=selection,settings=options();setBusy(true);
 try{const {buildWorkbook}=await import('./excel');const workbook=await buildWorkbook(chosen,settings),bytes=await workbook.xlsx.writeBuffer();const blob=new Blob([new Uint8Array(bytes)],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});const url=URL.createObjectURL(blob),a=node('a');a.href=url;a.download=`${demo?'DEMO_':''}REI_${chosen.groups[0].location.slice(4)}-${chosen.groups.at(-1)!.location.slice(4)}_${settings.date}.xlsx`;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);setBusy(false);$('download-status').textContent='已生成 Excel。打印时选择「整个工作簿」，并先检查打印预览。';}
 catch(error){setBusy(false);$('issues').append(notice(error instanceof Error?error.message:String(error),true));}
};
update();
