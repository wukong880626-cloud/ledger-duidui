export const VERSION='0.1.1';
export const LIMITS={bytes:20*1024*1024,rows:25000,headerRecords:100,columns:200,cells:150000,cellChars:32766,keyChars:256,keys:8,fields:20,resultRows:25000,zipEntries:1000,expandedBytes:50*1024*1024,sheets:20};
export function validateText(value){
  if(value.length>LIMITS.cellChars)throw new Error('单元格文本超过 32,766 字符，请先整理；未截断原值');
  if(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/.test(value))throw new Error('文件含不能无损写入 XLSX 的控制字符，请先在副本中清理');
}
