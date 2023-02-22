import * as XLSX  from 'xlsx-js-style';

export default async function SaveNewWorkBook(sheets: XLSX.WorkSheet[], dir:string,file:string, namePlan: string){

  const pathFormated = `${dir}${namePlan}/${file.toUpperCase()}`

  const workBook = await XLSX.utils.book_new()

  sheets.forEach((sheet) => {
    XLSX.utils.book_append_sheet(workBook, sheet)
  });
  
  await XLSX.writeFile(workBook, pathFormated)

  console.log('saved')
}