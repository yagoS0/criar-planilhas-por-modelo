import XLSX from 'xlsx-js-style'


export default async function InsertFolhaInTable(
  sheet:XLSX.WorkSheet,celulas:string[][],indice: number[]){
    const letras: number[] = [0,1,2,3,4]

    const inicio = indice[0] + 5
    console.log(inicio)
    sheet[celulas[inicio][0]] = {
      v: "Teste de celula",
      s:{
        font: {
          sz: 12,
          bold: true,
          color: {
            rgb: '000',
          },
        },
      alignment: {
          horizontal: 'center',
        },
      border: {
          top: {
            style: 'thick',
          },
          bottom: {
            style: 'thick',
          },
          right: {
            style: 'thick',
          },
          left: {
            style: 'thick',
          },
      },
    }}
  
  return sheet
} 


