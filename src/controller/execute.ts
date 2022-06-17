import * as XLSX  from 'xlsx';
import {format} from 'date-fns'
import ptBR from 'date-fns/locale/pt-BR'

import * as fs from 'fs';
import sheetToArray from '../service/sheetToArray';

export default async function execute(dir: string) {
  console.log(dir)
  const files = fs.readdirSync(dir);

  files.forEach(async (file) => {
    console.log(file)
    const workbook =  XLSX.readFile(`${dir}/${file}`);

    // for (let i = 0; i < workbook.SheetNames.length; i++) {


    if (workbook.Sheets[workbook.SheetNames[1]] === undefined) {
      console.log(`Nao existe nada aqui - ${dir}/${file}`)
      return null
    }

      const sheet = await workbook.Sheets[workbook.SheetNames[1]]

      const sheetArray = await sheetToArray(sheet)

      var dateYear = 2022
      var dateMonth = 0

      const date = new Date(dateYear,dateMonth,1)

      var ultimoDia = format( new Date(date.getFullYear(), date.getMonth() + 1, 0), 'dd/MM/yyyy');
      var monthNumber = date.toLocaleString('pt-BR', { month: '2-digit'});

      var month = date.toLocaleString('pt-BR', { month: 'long'});
      var year = date.toLocaleString('pt-BR', { year: 'numeric'});

      sheetArray.forEach(async (row: Array<string>)=> {
        const row1 = parseInt(row[1])
        const row2 = parseInt(row[2])

        if(row1 || row2) {
          row[0] =  ultimoDia
        }

        // console.log(row1String,row2String)
        if (row[0] === '') {
            const dataDeTroca = row[3].split("/")
            if(dataDeTroca.length === 2) {
              row[3] = row[3].replace(dataDeTroca[0], month.toUpperCase())
            }
        }
        var searchDate = /(\d{2})[-.\/](\d{4})/.exec(row[3]);

        if(searchDate){
          console.log(searchDate[0])
          row[3] = row[3].replace(searchDate[0], `${monthNumber}/${year}`)
        }

        const arrayBuffer =  JSON.stringify(sheetArray)
        await fs.writeFileSync(`${__dirname}/textArray.txt`,arrayBuffer)

      })
    // }

  });

}
