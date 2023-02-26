import * as XLSX  from 'xlsx-js-style';
import * as fs from 'fs';

import SaveNewWorkBook from '../service/SaveNewWorkBook';
import SpaceStyle from '../functions/SpaceStyle';
import styleSheet from '../functions/styleSheet'
import MapRows from '../functions/MapRows'

export default async function execute(dir) {

  const year = 2023

  const arrayMonthText  = [
    `JAN-${year}`,
    `FEV-${year}`,
    `MAR-${year}`,
    `ABR-${year}`,
    `MAI-${year}`,
    `JUN-${year}`,
    `JUL-${year}`,
    `AGO-${year}`,
    `SET-${year}`,
    `OUT-${year}`,
    `NOV-${year}`,
    `DEZ-${year}`,
  ]


  console.log(dir)
  const files = fs.readdirSync(dir);
  
  const dirNovasPlanilhas = `/Novas-Planilhas${Math.random() * 10}`

  fs.mkdirSync(`${dir}${dirNovasPlanilhas}`)

  files.forEach(async (file) => {
    console.log(file)

    const workbook = XLSX.readFile(`${dir}/${file}`);

    const newWorkBook =  XLSX.utils.book_new()

    const pathFormated = `${dir}${dirNovasPlanilhas}/${file.toUpperCase()}`
       
    for(let sheetNumber = 0; sheetNumber <= 11; sheetNumber++ ){
      if (workbook.Sheets[workbook.SheetNames[sheetNumber]] === undefined) {
        console.log(`Nao existe nada aqui - ${dir}/${file}`)
        return null
      }
      const bookSheets =  workbook.Sheets[workbook.SheetNames[sheetNumber]]

      const rowObject = await MapRows(sheetNumber,bookSheets)
      
      await SpaceStyle(bookSheets)

      await styleSheet(bookSheets, rowObject, sheetNumber)
      

      XLSX.utils.book_append_sheet(newWorkBook, bookSheets,arrayMonthText[sheetNumber])

    } 
    await XLSX.writeFile(newWorkBook, pathFormated)

  })
}

