import * as XLSX  from 'xlsx-js-style';

import DateFormat from '../utils/DateFormat';

async function AtualizaPlanilha(bookSheets, r,sheetNumber){

  const rowDebito = XLSX.utils.encode_cell({c: 1, r: r})
  const rowCredito = XLSX.utils.encode_cell({c: 2, r: r})

  const debitoValue = parseInt(bookSheets[rowDebito]?.v)
  const creditoValue = parseInt(bookSheets[rowCredito]?.v)

   // Altera a data
    const date = await DateFormat(sheetNumber)

    const rowDate = XLSX.utils.encode_cell({c: 0, r: r})

    const debitoString = JSON.stringify(debitoValue)
    const creditoString = JSON.stringify(creditoValue)

    console.log(debitoString, creditoString)

    if(debitoString !== 'null'|| creditoString !== 'null'){

      bookSheets[rowDate] = {
        v: `${date.lastDay}/${date.dateMonth}/${date.year}`
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


    return bookSheets
}

export default AtualizaPlanilha