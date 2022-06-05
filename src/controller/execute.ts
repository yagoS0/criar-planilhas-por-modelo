import * as XLSX  from 'xlsx';

import * as fs from 'fs';
import sheetToArray from '../service/sheetToArray';

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

    const arrayBuffer = JSON.stringify(sheetArray)

    fs.writeFileSync(`${__dirname}/textArray.txt`,arrayBuffer)
  });

}
