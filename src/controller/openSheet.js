import * as XLSX  from 'sheetjs-style';
import * as fs from 'fs';
import SaveNewWorkBook from '../service/SaveNewWorkBook';
import DateFormat from '../utils/DateFormat';

import AlteredRows from './AlteredRows'
import sheetToArray from '../functions/ArrayFunctions/sheetToArray';
import StyleSheets from '../functions/StyleSheets';

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

    
       
    for(let i = 0; i <= 11; i++ ){
      if (workbook.Sheets[workbook.SheetNames[i]] === undefined) {
        console.log(`Nao existe nada aqui - ${dir}/${file}`)
        return null
      }
      const bookSheets =  workbook.Sheets[workbook.SheetNames[i]]
        for (let r = 0; r <= 100; r ++) {
          
          // Debitos e Creditos
          const rowDebito = XLSX.utils.encode_cell({c: 1, r: r})
          const rowCredito = XLSX.utils.encode_cell({c: 2, r: r})

          const debitoValue = parseInt(bookSheets[rowDebito]?.v)
          const creditoValue = parseInt(bookSheets[rowCredito]?.v)
         

          // Altera a data
            const date = await DateFormat(i)

            const rowDate = XLSX.utils.encode_cell({c: 0, r: r})

            const debitoString = JSON.stringify(debitoValue)
            const creditoString = JSON.stringify(creditoValue)

            console.log(debitoString, creditoString)

            if(debitoString !== 'null'|| creditoString !== 'null'){

              bookSheets[rowDate] = {
                v: `${date.lastDay}/${date.dateMonth}/${date.year + 1}`
              } 

              if(debitoValue === 232 & creditoValue === 5){
              
                bookSheets[rowDate] = {
                  v: `05/${date.dateMonth + 1}/${date.year + 1}`
                } 
              } 
            }

           

            ///////////////////////////////////////////////////////

            // Troca data do texto
            const rowTexto = XLSX.utils.encode_cell({c: 3, r: r})

            const cellText = bookSheets[rowTexto]?.v

            const splitBarra = cellText?.split('/')

            if(splitBarra?.length === 2){
              const splitSpace = splitBarra[1]?.split(" ")
              const yearInt = parseInt(splitSpace[0])

              if (yearInt !== NaN) {
                bookSheets[rowTexto] = {
                  v: cellText?.replace(yearInt, '2023') 
                } 
              }
            }
            ////////////////////////////////////////////////
      
            // Apagando valores das celulas
            const rowValue = XLSX.utils.encode_cell({c: 4, r: r})
            const cellValue = bookSheets[rowValue]?.v
           

            if (cellValue !==  'VALOR') {
              bookSheets[rowValue] = {
                v: ' '
              } 
            }
            ///////////////////////////////////////////////////////
        }
    
   

      ////// ESSE CODIGO A BAIXO FUNCIONA////////////////////
      // const address = XLSX.utils.encode_cell({c: 3, r: 1});
      // const cell = bookSheets['D2'];
      // const value = XLSX.utils.format_cell(cell);
      // console.log(value)
      // cell?.v = 'Essa passou'
      ////////////////////////////////////////////////////////////

      // await StyleSheets(sheet)
      
      sheets.push(bookSheets)

      if (sheet_range === 11) {
        await SaveNewWorkBook(sheets,dir,file, dirNovasPlanilhas)
      }
    }   
   
    
     
  })
}
