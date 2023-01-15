import * as XLSX  from 'xlsx-js-style';
import * as fs from 'fs';

import SaveNewWorkBook from '../service/SaveNewWorkBook';
import SpaceStyle from '../functions/SpaceStyle';
import styleSheet from '../functions/styleSheet'
import MapRows from '../functions/MapRows'

export default async function execute(dir) {

  console.log(dir)
  const files = fs.readdirSync(dir);

  const dirNovasPlanilhas = `/Novas-Planilhas${Math.random() * 10}`

  fs.mkdirSync(`${dir}${dirNovasPlanilhas}`)

  files.forEach(async (file) => {
    console.log(file)

    const workbook = XLSX.readFile(`${dir}/${file}`);

    const sheet_range = workbook.SheetNames.length - 3
    
    const sheets = []
       
    for(let sheetNumber = 0; sheetNumber <= 11; sheetNumber++ ){
      if (workbook.Sheets[workbook.SheetNames[sheetNumber]] === undefined) {
        console.log(`Nao existe nada aqui - ${dir}/${file}`)
        return null
      }
      const bookSheets =  workbook.Sheets[workbook.SheetNames[sheetNumber]]

      const rowObject = await MapRows(sheetNumber,bookSheets)

      
      await styleSheet(bookSheets, rowObject)
      
      await SpaceStyle(bookSheets)

        
      sheets.push(bookSheets)
        
      if (sheet_range === 11) {
        await SaveNewWorkBook(sheets,dir,file, dirNovasPlanilhas)
      }
    }    
  })
}


// bookSheets["A1"].s = {
//   font: { sz: 50, color: { rgb: '00f' } },
//   border: { top: { style: 'bold' }, bottom: { style: 'bold' } }
// }