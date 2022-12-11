import SaveNewWorkBook from '../service/SaveNewWorkBook';
import AlteredDate from '../functions/alteredRows/alteredDate';
import AlteredDateInHeaderAndText from '../functions/alteredRows/AlteredDateInHeaderAndText';
import DateFormat from '../utils/DateFormat';
const XLSX = require('sheetjs-style')


async function AlteredRows(sheetArray: Array<string[]>, month: number, dir:string, file:string) {

    const date =  DateFormat(month)

    const sheet:Array<string[]> = [] 

    sheetArray.forEach(async (row: Array<string>)=> {
      await AlteredDateInHeaderAndText(row)

      // await AlteredDate(row, date)
   
      
    })
    return sheetArray
    // if(sheetArray.length){
    //   
    // }


}


export default AlteredRows;
