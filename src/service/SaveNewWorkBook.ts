import * as fs from 'fs';
import * as XLSX  from 'xlsx';

const st = require('sheetjs-style')

export default async function SaveNewWorkBook(sheets: XLSX.WorkSheet[], dir:string,file:string){

  const pathFormated = `${dir}/Novas-Planilhas/${file.toUpperCase()}.xlsx`

  const workBook = await XLSX.utils.book_new()

  sheets.forEach((sheet) => {
    XLSX.utils.book_append_sheet(workBook, sheet)
  });
  

  await XLSX.writeFile(workBook, pathFormated)

  console.log('saved')
}