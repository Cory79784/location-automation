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
async function upload(files:Array<{name:string;location:string;sku?:string}>){
 const objects=[];
 for(const file of files){const book=new ExcelJS.Workbook(),sheet=book.addWorksheet('Stock');sheet.addRow(['SKU编码','Barcode','库区','库位','库位库存']);sheet.addRow([file.sku||'S1','00123','H1',file.location,2]);const bytes=await book.xlsx.writeBuffer();objects.push({name:file.name,arrayBuffer:async()=>new Uint8Array(bytes).buffer,size:bytes.byteLength});}
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
