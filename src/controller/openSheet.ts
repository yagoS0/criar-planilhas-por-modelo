import * as XLSX  from 'xlsx';
import * as fs from 'fs';
import SaveNewWorkBook from '../service/SaveNewWorkBook';
import AlteredRows from './AlteredRows'
import sheetToArray from '../functions/ArrayFunctions/sheetToArray';

export default async function execute(dir: string) {
  console.log(dir)
  const files = fs.readdirSync(dir);

  fs.mkdirSync(`${dir}/Novas-Planilhas`)

  files.forEach(async (file) => {
    console.log(file)

    const workbook = XLSX.readFile(`${dir}/${file}`);

    const sheet_range = workbook.SheetNames.length - 3
    
    const sheets = []
    
    for(let i = 0; i <= sheet_range; i++ ){
      if (workbook.Sheets[workbook.SheetNames[i]] === undefined) {
        console.log(`Nao existe nada aqui - ${dir}/${file}`)
        return null
      }
      const bookSheets =  workbook.Sheets[workbook.SheetNames[i]]

      const sheetArray = await sheetToArray(bookSheets)

      await AlteredRows(sheetArray, i, dir, file)


      const sheet =  XLSX.utils.aoa_to_sheet(sheetArray)
      sheets.push(sheet)

      if (sheet_range === 11) {
        await SaveNewWorkBook(sheets,dir,file)
      }
    }   
   
    
     
  })
}
