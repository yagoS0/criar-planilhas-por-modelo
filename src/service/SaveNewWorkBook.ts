import * as fs from 'fs';

export default async function SaveNewWorkBook(sheetArray: Array<string[]>,month: number){

  const arrayBuffer = JSON.stringify(sheetArray)
  fs.writeFileSync(`${__dirname}/textArray${month += 1}.txt`,arrayBuffer)
  console.log('saved')
}
