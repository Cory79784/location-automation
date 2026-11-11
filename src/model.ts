export interface StockRow {location:string;sku:string;barcode:string;qty:number;availableQty?:number|null;availabilityIssue?:'missing-column'|'blank'|'invalid';hall:string;customer:string;source:string;sourceRow:number}
export interface SourceData {name:string;rows:StockRow[];ignored:number;warnings:string[]}
export interface LocationGroup {location:string;rows:StockRow[];pieces:number;missing:boolean}
export interface Selection {groups:LocationGroup[];warnings:string[];pieces:number;missing:string[]}
export const MAX_PIECES=20000;
export function localDateString():string {
 const date=new Date();return `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`;
}
export function validPrintDate(value:string):boolean {
 if(!/^\d{4}-\d{2}-\d{2}$/.test(value))return false;
 const date=new Date(`${value}T00:00:00Z`);return !Number.isNaN(date.getTime())&&date.toISOString().slice(0,10)===value;
}
export function availabilityWarning(row:StockRow):string|null {
 const prefix=`${row.location} / ${row.sku}`;
 if(row.availabilityIssue){
  const reason=row.availabilityIssue==='missing-column'?'源表未提供可用库存列':row.availabilityIssue==='blank'?'可用库存为空白':'可用库存格式无效';
  return `${prefix}：${reason}，无法判断可用数量；仍按库位库存 ${row.qty} 件打印。`;
 }
 if(typeof row.availableQty==='number'&&row.availableQty<row.qty)return `${prefix}：可用库存 ${row.availableQty} < 库位库存 ${row.qty}，请留意占用或冻结情况；仍打印 ${row.qty} 件。`;
 return null;
}
export function locationNumber(value:string):number|null {
 const match=/^REI-(\d{1,3})$/i.exec(value.trim());
 if(!match)return null;
 const number=Number(match[1]);return number>=1&&number<=100?number:null;
}
export const canonicalLocation=(value:string)=>{const number=locationNumber(value);return number===null?null:`REI-${number}`;};
export function combineSources(sources:SourceData[]):StockRow[] {
 const owner=new Map<string,string>(),rows:StockRow[]=[];
 for(const source of sources){
  for(const location of new Set(source.rows.map(row=>row.location))){
   const previous=owner.get(location);
   if(previous!==undefined)throw new Error(`${location} 同时出现在「${previous}」和「${source.name}」中。请移除重叠文件，避免重复计数。`);
   owner.set(location,source.name);
  }
  rows.push(...source.rows);
 }
 for(const location of owner.keys()){
  const halls=new Set(rows.filter(row=>row.location===location).map(row=>row.hall));
  if(halls.size>1)throw new Error(`${location} 跨多个库区（${[...halls].join('、')}）。请在 WMS 先筛选唯一库区后重新导出。`);
 }
 return rows;
}
export function selectLocations(rows:StockRow[],range?:{start:number;end:number}):Selection {
 if(range&&(!Number.isInteger(range.start)||!Number.isInteger(range.end)||range.start<1||range.end>100||range.start>range.end))throw new Error('请输入有效范围：1～100，起始库位不能大于结束库位。');
 const map=new Map<string,StockRow[]>();for(const row of rows){const list=map.get(row.location)||[];list.push(row);map.set(row.location,list);}
 const locations=range?Array.from({length:range.end-range.start+1},(_,index)=>`REI-${range.start+index}`):[...map.keys()].sort((a,b)=>locationNumber(a)!-locationNumber(b)!);
 const groups=locations.map(location=>{const own=map.get(location)||[];return {location,rows:own,pieces:own.reduce((sum,row)=>sum+row.qty,0),missing:!map.has(location)};});
 const pieces=groups.reduce((sum,group)=>sum+group.pieces,0);
 if(pieces>MAX_PIECES)throw new Error(`本次共 ${pieces} 件，超过 ${MAX_PIECES} 件的批次上限，请缩小库位范围。`);
 const warnings=groups.filter(group=>group.pieces>35).map(group=>`${group.location} 有 ${group.pieces} 件；仍按一页输出，打印字体和手写空间会缩小，请先检查打印预览。`);
 return {groups,pieces,missing:groups.filter(group=>group.missing).map(group=>group.location),warnings};
}
export function expandRows(group:LocationGroup):Array<[string,string,string,string]> {
 return group.rows.flatMap(row=>Array.from({length:row.qty},()=>[row.sku,row.barcode,group.location,''] as [string,string,string,string]));
}
