import XLSX from 'xlsx-js-style'

import AtualizaPlanilha from './AtualizaPlanilha'

export default async function MapRows(sheetNumber:number, bookSheets:  XLSX.WorkSheet, ){


    const arrayDebitoCredito: {
      dataValue: string;
      debitoValue: number;
      creditoValue: number;
      textValue: string;
      value: string;
    }[] = []

  for (let r = 0; r <= 100; r ++) {

    await AtualizaPlanilha(bookSheets, r, sheetNumber)


    const rowData = XLSX.utils.encode_cell({c: 0, r: r})
    const rowDebito = XLSX.utils.encode_cell({c: 1, r: r})
    const rowCredito = XLSX.utils.encode_cell({c: 2, r: r})
    const rowText = XLSX.utils.encode_cell({c: 3, r: r})
    const rowValue = XLSX.utils.encode_cell({c: 4, r: r})
    

    const dataValue =   bookSheets[rowData]?.v
    const debitoValue = parseInt(bookSheets[rowDebito]?.v)
    const creditoValue = parseInt(bookSheets[rowCredito]?.v)
    const textValue = bookSheets[rowText]?.v
    const value = bookSheets[rowValue]?.v
    
    arrayDebitoCredito.push({
        dataValue,
        debitoValue,
        creditoValue,
        textValue,
        value

    })
   
    
  }

  return arrayDebitoCredito

}