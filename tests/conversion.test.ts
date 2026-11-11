import {describe,it,expect} from 'vitest';
import ExcelJS from 'exceljs';
import {readSource,buildWorkbook} from '../src/excel';
import {combineSources,selectLocations,canonicalLocation,availabilityWarning,validPrintDate,type StockRow,type SourceData} from '../src/model';
const stock=(location:string,qty=1,sku='TEST-SKU'):StockRow=>({location,qty,sku,barcode:'001234567890',hall:'H1',customer:'TEST',source:'test.xlsx',sourceRow:2});
const source=(name:string,rows:StockRow[]):SourceData=>({name,rows,ignored:0,warnings:[]});
async function input(values:ExcelJS.CellValue[][],name='input.xlsx'){
 const book=new ExcelJS.Workbook(),sheet=book.addWorksheet('库位库存数据');sheet.addRow(['SKU编码','Barcode','库区','库位','库位库存']);for(const row of values)sheet.addRow(row);
 const bytes=await book.xlsx.writeBuffer();return readSource(name,new Uint8Array(bytes).buffer);
}
describe('WMS Excel → per-location RT worksheets',()=>{
 it('accepts a consolidated export and sorts tabs numerically, with normalized REI codes',async()=>{
  const parsed=await input([['A','00001','H1','ReI-20',1],['B','00002','H1','rei-10',2],['C','00003','H1','REI-11',1],['X','00004','H1','101-001-1',9]]);
  expect(parsed.rows[1]).toMatchObject({location:'REI-10',barcode:'00002',qty:2});expect(parsed.ignored).toBe(1);
  const selected=selectLocations(combineSources([parsed]));expect(selected.groups.map(g=>g.location)).toEqual(['REI-10','REI-11','REI-20']);expect(selected.pieces).toBe(4);
 });
 it('accepts multiple single-location files and expands quantity, including repeated source rows',async()=>{
  const a=await input([['A','0001','H1','REI-12',2],['A','0001','H1','REI-12',1]],'a.xlsx');
  const b=await input([['B','0002','H1','REI-11',1]],'b.xlsx');
  const book=await buildWorkbook(selectLocations(combineSources([a,b])),{date:'2026-11-11',name:'Test',allowMissing:false});
  expect(book.worksheets.map(s=>s.name)).toEqual(['REI-11','REI-12']);const sheet=book.worksheets[1];
  expect([5,6,7].map(row=>sheet.getCell(`A${row}`).value)).toEqual(['A','A','A']);expect(sheet.getCell('A8').value).toBeNull();expect(sheet.getCell('D5').value).toBe('');
 });
 it('rejects overlapping files instead of double counting, and keeps distinct halls from mixing',()=>{
  expect(()=>combineSources([source('a',[stock('REI-10')]),source('b',[stock('REI-10')])])).toThrow('同时出现在');
  expect(()=>combineSources([source('a',[stock('REI-10'),{...stock('REI-10'),hall:'H2'}])])).toThrow('跨多个库区');
 });
 it('rejects the same location repeated across data worksheets in a single workbook',async()=>{
  const book=new ExcelJS.Workbook();for(const name of ['Snapshot A','Snapshot B']){const sheet=book.addWorksheet(name);sheet.addRow(['SKU编码','Barcode','库位','库位库存']);sheet.addRow(['A','B','REI-10',1]);}
  const bytes=await book.xlsx.writeBuffer();await expect(readSource('overlap.xlsx',new Uint8Array(bytes).buffer)).rejects.toThrow('移除重叠数据页');
 });
 it('makes a 10–20 inclusive range of eleven sheets; distinguishes missing data from reported zero',async()=>{
  const selected=selectLocations([stock('REI-10',0),stock('REI-20',1)],{start:10,end:20});expect(selected.groups).toHaveLength(11);expect(selected.missing).toHaveLength(9);expect(selected.groups[0].missing).toBe(false);
  await expect(buildWorkbook(selected,{date:'2026-11-11',name:'',allowMissing:false})).rejects.toThrow('缺少库位数据');
  const book=await buildWorkbook(selected,{date:'2026-11-11',name:'',allowMissing:true});expect(book.worksheets[0].getCell('A5').value).toContain('源表库存为 0');expect(book.worksheets[1].getCell('A5').value).toContain('未提供该库位数据');
  expect(book.worksheets.map(s=>s.name)).toEqual(Array.from({length:11},(_,i)=>`REI-${i+10}`));
 });
 it('writes original identifiers as text and makes every sheet A4 portrait, one page, with a bounded print area',async()=>{
  const book=await buildWorkbook(selectLocations([stock('REI-22',2,'0000123')]),{date:'2026-11-11',name:'Alice',allowMissing:false});
  const bytes=await book.xlsx.writeBuffer(),reopened=new ExcelJS.Workbook();await reopened.xlsx.load(bytes);const sheet=reopened.worksheets[0];
  expect(sheet.pageSetup).toMatchObject({paperSize:9,orientation:'portrait',fitToPage:true,fitToWidth:1,fitToHeight:1,printArea:'A1:D22'});
  expect(sheet.getCell('A5').value).toBe('0000123');expect(sheet.getCell('B5').value).toBe('001234567890');expect(sheet.getCell('A5').numFmt).toBe('@');expect(sheet.getCell('C5').value).toBe('REI-22');
  expect(sheet.getCell('A1').value).toBe('STARTZEIT');expect(sheet.getCell('C1').value).toBe('ENDZEIT');expect(sheet.getCell('B2').value).toBe('Alice');expect(sheet.getCell('D2').value).toBe('11.11.2026');expect(sheet.getCell('D5').value).toBe('');
 });
 it('preserves every row for high-density stock without zero-height or hidden rows',async()=>{
  const selected=selectLocations([stock('REI-1',100)]);expect(selected.warnings).toHaveLength(1);
  const book=await buildWorkbook(selected,{date:'2026-11-11',name:'',allowMissing:false}),sheet=book.worksheets[0];expect(sheet.pageSetup.printArea).toBe('A1:D104');expect(sheet.pageSetup.fitToHeight).toBe(1);expect(sheet.getRow(104).height).toBeGreaterThan(0);expect(sheet.getCell('A104').value).toBe('TEST-SKU');
 });
 it('rejects malformed quantities, templates and formulas, and reports the physical source row',async()=>{
  for(const qty of [-1,1.5,'oops',''])await expect(input([['A','B','H1','REI-10',qty]])).rejects.toThrow('第 2 行');
  await expect(input([['A','B','H1','REI-10',{formula:'1+1',result:2}]])).rejects.toThrow('含公式');
  const book=new ExcelJS.Workbook();book.addWorksheet('RT').addRow(['SKU','Barcode','Rei','Location']);const bytes=await book.xlsx.writeBuffer();await expect(readSource('template.xlsx',new Uint8Array(bytes).buffer)).rejects.toThrow('而非打印模板');
 });
 it('preserves numeric identifiers with zero-padding and warns on blank barcodes',async()=>{
  const book=new ExcelJS.Workbook(),sheet=book.addWorksheet('Stock');sheet.addRow(['SKU编码','Barcode','库位','库位库存']);sheet.addRow([12,'','REI-1',1]);sheet.getCell('A2').numFmt='000000';const bytes=await book.xlsx.writeBuffer();const parsed=await readSource('numeric.xlsx',new Uint8Array(bytes).buffer);
  expect(parsed.rows[0].sku).toBe('000012');expect(parsed.warnings[0]).toContain('缺少 Barcode');
 });
 it('validates REI boundaries and ranges, and prevents oversized expansion',()=>{
  expect(canonicalLocation(' ReI-001 ')).toBe('REI-1');for(const value of ['REI-0','REI-101','REI-1-2','XREI-1'])expect(canonicalLocation(value)).toBeNull();
  expect(()=>selectLocations([],{start:20,end:10})).toThrow('有效范围');expect(()=>selectLocations([stock('REI-1',20001)])).toThrow('上限');
 });
});

it('warns about zero and partial availability without changing the printed unit count',async()=>{
 const book=new ExcelJS.Workbook(),sheet=book.addWorksheet('Stock');sheet.addRow(['SKU编码','Barcode','库位','库位库存','可用库存']);
 sheet.addRow(['ZERO','B0','REI-10',3,0]);sheet.addRow(['PARTIAL','B1','REI-10',2,1]);sheet.addRow(['FULL','B2','REI-10',1,1]);sheet.addRow(['BLANK','B3','REI-10',2,null]);
 const bytes=await book.xlsx.writeBuffer(),parsed=await readSource('stock.xlsx',new Uint8Array(bytes).buffer);
 expect(parsed.rows[0].availableQty).toBe(0);expect(availabilityWarning(parsed.rows[0])).toContain('可用库存 0 < 库位库存 3');expect(availabilityWarning(parsed.rows[1])).toContain('可用库存 1 < 库位库存 2');expect(availabilityWarning(parsed.rows[2])).toBeNull();expect(availabilityWarning(parsed.rows[3])).toContain('为空白');expect(parsed.rows[3].availableQty).toBeNull();
 const output=await buildWorkbook(selectLocations(parsed.rows),{date:'2026-12-25',name:'',allowMissing:false});
 expect([5,6,7,8,9,10,11,12].map(row=>output.worksheets[0].getCell(`A${row}`).value)).toEqual(['ZERO','ZERO','ZERO','PARTIAL','PARTIAL','FULL','BLANK','BLANK']);
 expect(output.worksheets[0].getCell('A13').value).toBeNull();expect(output.worksheets[0].getCell('D2').value).toBe('25.12.2026');expect(output.worksheets[0].pageSetup.fitToHeight).toBe(1);
});
it('treats absent, invalid and formula availability as advisory and never blocks stock printing',async()=>{
 const missing=await input([['A','B','H1','REI-10',2]]);expect(availabilityWarning(missing.rows[0])).toContain('未提供可用库存列');
 const book=new ExcelJS.Workbook(),sheet=book.addWorksheet('Stock');sheet.addRow(['SKU编码','Barcode','库位','库位库存','可用库存']);
 for(const available of ['?',1.5,{formula:'0',result:0},-1] as ExcelJS.CellValue[])sheet.addRow(['A','B','REI-10',2,available]);
 const bytes=await book.xlsx.writeBuffer(),parsed=await readSource('advisory.xlsx',new Uint8Array(bytes).buffer);
 expect(parsed.rows.slice(0,3).every(row=>row.availabilityIssue==='invalid')).toBe(true);expect(availabilityWarning(parsed.rows[3])).toContain('可用库存 -1');
 const out=await buildWorkbook(selectLocations(parsed.rows),{date:'2026-11-11',name:'',allowMissing:false});expect(out.worksheets[0].getCell('A12').value).toBe('A');expect(out.worksheets[0].getCell('A13').value).toBeNull();
});
it('validates calendar dates and applies one chosen date to all output tabs',async()=>{
 expect(validPrintDate('2024-02-29')).toBe(true);for(const date of ['','2026-02-29','2026-13-01','2026-12-32'])expect(validPrintDate(date)).toBe(false);
 const selected=selectLocations([stock('REI-1'),stock('REI-2')]);await expect(buildWorkbook(selected,{date:'2026-02-29',name:'',allowMissing:false})).rejects.toThrow('有效的打印日期');
 const out=await buildWorkbook(selected,{date:'2026-12-25',name:'',allowMissing:false});expect(out.worksheets.every(sheet=>sheet.getCell('D2').value==='25.12.2026')).toBe(true);
});
