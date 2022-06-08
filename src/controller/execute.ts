import * as XLSX  from 'xlsx';

import * as fs from 'fs';
import sheetToArray from '../service/sheetToArray';
import { parse } from 'path';
import { types } from '@babel/core';
import { lastDayOfDecade } from 'date-fns';

export default async function execute(dir: string) {
  console.log(dir)

  const files = fs.readdirSync(dir);

  files.forEach(async (file) => {
    console.log(file)
    const workbook =  XLSX.readFile(`${dir}/${file}`);

    if (workbook.Sheets[workbook.SheetNames[1]] === undefined) {
      console.log(`Nao existe nada aqui - ${dir}/${file}`)
      return null
    }
    const sheet = await workbook.Sheets[workbook.SheetNames[1]]

    const sheetArray = await sheetToArray(sheet)

    const meses = [
      '',
      'JANEIRO',
      'FEVEREIRO',
      'MARCO',
      'ABRIL',
      'MAIO',
      'JUNHO',
      'JULHO',
      'AGOSTO',
      'SETEMBRO',
      'OUTUBRO',
      'NOVEMBRO',
      'DEZEMBRO'
    ]

    sheetArray.forEach(async (row: Array<string>)=> {

      const row1 = parseInt(row[1])
      const row2 = parseInt(row[2])

      if(row1 || row2) {
        row[0] = "31/03/2022"
      }

      for (let index = 1; index < meses.length; index++) {
        if(row[3].indexOf(meses[index]) === 0){
          row[3] = row[3].replace(meses[index], meses[index - 1])
            console.log(row[3])
         }
      }



    })

    const arrayBuffer =  JSON.stringify(sheetArray)
    await fs.writeFileSync(`${__dirname}/textArray.txt`,arrayBuffer)
  });

}
