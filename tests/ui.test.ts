import {afterEach,beforeEach,expect,it,vi} from 'vitest';
import {JSDOM} from 'jsdom';
import ExcelJS from 'exceljs';
let dom:JSDOM;
const query=<T extends HTMLElement=HTMLElement>(selector:string)=>document.querySelector<T>(selector)!;
beforeEach(async()=>{
 vi.resetModules();dom=new JSDOM('<div id="app"></div>',{url:'http://localhost'});
 Object.assign(globalThis,{document:dom.window.document,window:dom.window,HTMLElement:dom.window.HTMLElement});
 await import('../src/main');
});
afterEach(()=>dom.window.close());
async function upload(files:Array<{name:string;location:string;sku?:string;available?:ExcelJS.CellValue}>){
 const objects=[];
 for(const file of files){const book=new ExcelJS.Workbook(),sheet=book.addWorksheet('Stock');sheet.addRow(['SKU编码','Barcode','库区','库位','库位库存','可用库存']);sheet.addRow([file.sku||'S1','00123','H1',file.location,2,file.available===undefined?2:file.available]);const bytes=await book.xlsx.writeBuffer();objects.push({name:file.name,arrayBuffer:async()=>new Uint8Array(bytes).buffer,size:bytes.byteLength});}
 Object.defineProperty(query('#files'),'files',{configurable:true,value:objects});query('#files').dispatchEvent(new dom.window.Event('change'));
 await vi.waitFor(()=>expect(query<HTMLInputElement>('#files').disabled).toBe(false));
}
it('shows eleven numerically ordered demo tabs and explicitly labels synthetic data',()=>{
 query<HTMLButtonElement>('#demo').click();expect(query('#total-locations').textContent).toBe('11');
 expect(Array.from(document.querySelectorAll('[role="tab"]')).map(el=>el.textContent)).toEqual(Array.from({length:11},(_,i)=>`REI-${10+i}`));
 expect(query('#issues').textContent).toContain('不代表真实库存');expect(query<HTMLButtonElement>('#download').disabled).toBe(false);
});
it('imports multiple separate files, previews quantity expansion and keeps SKU content as text',async()=>{
 await upload([{name:'second.xlsx',location:'REI-12'},{name:'first.xlsx',location:'REI-11',sku:'<img src=x onerror=alert(1)>'}]);
 expect(query('#total-locations').textContent).toBe('2');expect(query('#total-pieces').textContent).toBe('4');expect(query('#paper h3').textContent).toBe('REI-11');expect(query('#paper img')).toBeNull();expect(query('#paper').textContent).toContain('<img src=x onerror=alert(1)>');
});
it('blocks overlapping snapshots until the conflicting file is removed',async()=>{
 await upload([{name:'old.xlsx',location:'REI-10'},{name:'new.xlsx',location:'REI-10'}]);expect(query('#issues').textContent).toContain('同时出现在');expect(query<HTMLButtonElement>('#download').disabled).toBe(true);
 document.querySelectorAll<HTMLButtonElement>('.remove-file')[1].click();expect(query<HTMLButtonElement>('#download').disabled).toBe(false);expect(query('#total-pieces').textContent).toBe('2');
});
it('requires an explicit missing-data-sheet option for gaps rather than declaring them empty',async()=>{
 await upload([{name:'one.xlsx',location:'REI-10'}]);const range=query<HTMLInputElement>('input[value="range"]');range.checked=true;range.dispatchEvent(new dom.window.Event('input'));
 expect(query('#total-locations').textContent).toBe('11');expect(query('#issues').textContent).toContain('REI-11');expect(query<HTMLButtonElement>('#download').disabled).toBe(true);
 const allow=query<HTMLInputElement>('#allow-missing');allow.checked=true;allow.dispatchEvent(new dom.window.Event('input'));expect(query<HTMLButtonElement>('#download').disabled).toBe(false);
 (document.querySelectorAll('[role="tab"]')[1] as HTMLButtonElement).click();expect(query('#paper').textContent).toContain('未提供该库位数据');
});
it('replaces demo data when a real file is imported and does not transmit files',async()=>{
 const fetch=vi.fn();vi.stubGlobal('fetch',fetch);query<HTMLButtonElement>('#demo').click();await upload([{name:'real.xlsx',location:'REI-22'}]);
 expect(query('#total-locations').textContent).toBe('1');expect(query('#paper h3').textContent).toBe('REI-22');expect(query('#issues').textContent).not.toContain('演示数据');expect(fetch).not.toHaveBeenCalled();vi.unstubAllGlobals();
});

it('offers an explicit calendar control, updates the date preview, and resets to today',()=>{
 const date=query<HTMLInputElement>('#date'),today=date.value,picker=vi.fn();Object.defineProperty(date,'showPicker',{configurable:true,value:picker});
 query<HTMLButtonElement>('#pick-date').click();expect(picker).toHaveBeenCalledOnce();query<HTMLButtonElement>('#demo').click();date.value='2026-12-25';date.dispatchEvent(new dom.window.Event('input'));expect(query('.paper-header').textContent).toContain('25.12.2026');
 query<HTMLButtonElement>('#today-date').click();expect(date.value).toBe(today);date.value='';date.dispatchEvent(new dom.window.Event('input'));expect(query<HTMLButtonElement>('#download').disabled).toBe(true);expect(query('#issues').textContent).toContain('有效的打印日期');
});
it('shows zero and blank availability warnings only for selected locations and keeps download enabled',async()=>{
 await upload([{name:'zero.xlsx',location:'REI-10',sku:'ZERO',available:0},{name:'blank.xlsx',location:'REI-11',sku:'BLANK',available:null},{name:'full.xlsx',location:'REI-12',sku:'FULL',available:2}]);
 expect(query('#issues').textContent).toContain('可用库存 0 < 库位库存 2');expect(query('#issues').textContent).toContain('可用库存为空白');expect(query<HTMLButtonElement>('#download').disabled).toBe(false);expect(query('#total-pieces').textContent).toBe('6');
 const start=query<HTMLInputElement>('#from'),end=query<HTMLInputElement>('#to'),range=query<HTMLInputElement>('input[value="range"]');range.checked=true;start.value='12';end.value='12';range.dispatchEvent(new dom.window.Event('input'));
 expect(query('.availability-warnings')).toBeNull();expect(query('#total-pieces').textContent).toBe('2');expect(query<HTMLButtonElement>('#download').disabled).toBe(false);
});
