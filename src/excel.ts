import ExcelJS from 'exceljs';
import {canonicalLocation,expandRows,MAX_PIECES,type Selection,type SourceData,type StockRow} from './model';
const aliases={sku:['SKU编码','SKU'],barcode:['Barcode','条码'],location:['库位','库位编码','Location'],qty:['库位库存','库存数量','System quantity'],hall:['库区','库厅','Hall'],customer:['客户编码','Customer']};
function raw(cell:ExcelJS.Cell):unknown {
 const value=cell.value;
 if(value&&typeof value==='object'&&('formula' in value||'sharedFormula' in value))throw new Error(`单元格 ${cell.address} 含公式，请使用 WMS 原始导出值。`);
 if(value&&typeof value==='object'&&'richText' in value)return value.richText.map(part=>part.text).join('');
 if(value&&typeof value==='object'&&'text' in value)return value.text;
 return value;
}
function text(cell:ExcelJS.Cell):string {
 const value=raw(cell);
 if(value===null||value===undefined)return '';
 if(typeof value==='number'){
  if(!Number.isSafeInteger(value))throw new Error(`单元格 ${cell.address} 的编码不是安全整数，请在源文件中保留文本编码。`);
  const format=cell.numFmt;
  if(/^0+$/.test(format))return String(value).padStart(format.length,'0');
  if(format!=='General'&&format!=='@')throw new Error(`单元格 ${cell.address} 使用了特殊数字格式，请将编码保存为文本后重试。`);
 }
 if(typeof value!=='string'&&typeof value!=='number')throw new Error(`单元格 ${cell.address} 的编码格式无法识别。`);
 return String(value).trim();
}
export async function readSource(name:string,bytes:ArrayBuffer):Promise<SourceData>{
 if(!/\.xlsx$/i.test(name))throw new Error(`${name}：只支持 .xlsx 文件。`);
 if(bytes.byteLength>10*1024*1024)throw new Error(`${name}：文件超过 10 MB，请先筛选所需库位。`);
 const workbook=new ExcelJS.Workbook();
 try{await workbook.xlsx.load(bytes);}catch{throw new Error(`${name}：无法读取 Excel，请确认文件未加密且格式为 .xlsx。`);}
 const rows:StockRow[]=[],warnings:string[]=[],locationSheets=new Map<string,string>();let ignored=0,dataSheets=0,scanned=0;
 for(const sheet of workbook.worksheets){
  const columns:Record<string,number>={};
  const headers=sheet.getRow(1).values as ExcelJS.CellValue[];
  for(const [key,names] of Object.entries(aliases)){
   const found=headers.flatMap((value,index)=>names.some(alias=>String(value??'').trim().toLowerCase()===alias.toLowerCase())?[index]:[]);
   if(found.length>1)throw new Error(`${name} / ${sheet.name}：字段 ${key} 重复。`);
   columns[key]=found[0]||0;
  }
  if(!['sku','barcode','location','qty'].every(key=>columns[key]))continue;
  dataSheets++;
  if(sheet.rowCount>50001)throw new Error(`${name}：工作表超过 50,000 行，请缩小导出范围。`);
  for(let index=2;index<=sheet.rowCount;index++){
   if(++scanned>50000)throw new Error(`${name}：数据超过 50,000 行，请缩小导出范围。`);
   const row=sheet.getRow(index);if(!row.hasValues)continue;
   const get=(key:string)=>columns[key]?text(row.getCell(columns[key])):'';
   const inputLocation=get('location'),location=canonicalLocation(inputLocation);
   if(!location){ignored++;continue;}
   const priorSheet=locationSheets.get(location);if(priorSheet&&priorSheet!==sheet.name)throw new Error(`${name}：${location} 同时出现在工作表「${priorSheet}」和「${sheet.name}」。请移除重叠数据页，避免重复计数。`);locationSheets.set(location,sheet.name);
   try{
    const sku=get('sku'),barcode=get('barcode');if(!sku)throw new Error('SKU 为空');
    const quantity=raw(row.getCell(columns.qty));
    if(!/^[0-9]+$/.test(String(quantity??'').trim())||!Number.isSafeInteger(Number(quantity))||Number(quantity)>MAX_PIECES)throw new Error('库位库存必须是 0～20,000 的整数');
    if(!barcode)warnings.push(`${location} / ${sku} 缺少 Barcode，输出将留空。`);
    if(sku.length>35||barcode.length>30)warnings.push(`${location} 存在较长编码，打印时会缩小以完整显示。`);
    rows.push({location,sku,barcode,qty:Number(quantity),hall:get('hall'),customer:get('customer'),source:`${name} / ${sheet.name}`,sourceRow:index});
   }catch(error){throw new Error(`${name} / ${sheet.name} 第 ${index} 行：${error instanceof Error?error.message:String(error)}`);}
  }
 }
 if(!dataSheets)throw new Error(`${name}：找不到包含「SKU编码、Barcode、库位、库位库存」表头的数据页，请选择 WMS 库位库存导出，而非打印模板。`);
 if(!rows.length)throw new Error(`${name}：没有找到 REI-1～REI-100 的库存记录（已排除 ${ignored} 行）。`);
 if(rows.reduce((sum,row)=>sum+row.qty,0)>MAX_PIECES)throw new Error(`${name}：超过 ${MAX_PIECES} 件，请先在 WMS 缩小导出范围。`);
 return {name,rows,ignored,warnings:[...new Set(warnings)]};
}
export interface PrintOptions {date:string;name:string;allowMissing:boolean}
export async function buildWorkbook(selection:Selection,options:PrintOptions):Promise<ExcelJS.Workbook>{
 if(!selection.groups.length)throw new Error('没有可导出的库位。');
 if(selection.missing.length&&!options.allowMissing)throw new Error('所选范围缺少库位数据，请补充文件或选择生成缺失提示页。');
 const date=/^(\d{4})-(\d{2})-(\d{2})$/.exec(options.date);if(!date)throw new Error('请选择日期。');
 const displayDate=`${date[3]}.${date[2]}.${date[1]}`;
 const workbook=new ExcelJS.Workbook();workbook.creator='Location Automation';
 for(const group of selection.groups){
  const sheet=workbook.addWorksheet(group.location,{views:[{showGridLines:false,zoomScale:85}],pageSetup:{paperSize:9,orientation:'portrait',fitToPage:true,fitToWidth:1,fitToHeight:1,horizontalCentered:true,margins:{left:0.35,right:0.35,top:0.35,bottom:0.35,header:0,footer:0}}});
  sheet.columns=[{width:29},{width:21},{width:13},{width:25}];
  sheet.getRow(1).values=['STARTZEIT','','ENDZEIT',''];sheet.getRow(1).height=28;
  sheet.getRow(2).values=['Name',options.name,'Datum',displayDate];sheet.getRow(2).height=30;
  sheet.mergeCells('A3:D3');sheet.getCell('A3').value=group.location;sheet.getRow(3).height=28;
  sheet.getRow(4).values=['SKU','Barcode','Rei','Location'];sheet.getRow(4).height=30;
  const expanded=expandRows(group),lastRow=Math.max(22,expanded.length+4);
  const bodyHeight=Math.max(12,Math.min(32,Math.floor(650/(lastRow-4)))),fontSize=bodyHeight>=22?12:bodyHeight>=16?10:8;
  expanded.forEach((values,index)=>{sheet.getRow(index+5).values=values;});
  if(!expanded.length){sheet.mergeCells('A5:D5');sheet.getCell('A5').value=group.missing?'Keine Quelldaten / 未提供该库位数据':'Quellbestand 0 / 源表库存为 0';}
  for(let r=1;r<=lastRow;r++){
   if(r>4)sheet.getRow(r).height=bodyHeight;
   for(let c=1;c<=4;c++){
    const cell=sheet.getCell(r,c);cell.font={name:'Arial',size:r<=4?12:fontSize,bold:r===3||r===4};
    cell.alignment={vertical:'middle',horizontal:r===3?'center':'left',shrinkToFit:true};
    if(r!==3)cell.border={top:{style:'thin'},left:{style:'thin'},bottom:{style:'thin'},right:{style:'thin'}};
    if(r>=5&&c<=3)cell.numFmt='@';
   }
  }
  sheet.pageSetup.printArea=`A1:D${lastRow}`;
 }
 return workbook;
}
