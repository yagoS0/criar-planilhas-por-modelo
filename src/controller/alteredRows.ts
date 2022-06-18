import SaveNewWorkBook from '../service/SaveNewWorkBook';
import AlteredDate from '../functions/alteredRows/alteredDate';
import AlteredDateInHeader from '../functions/alteredRows/alteredDateInHeader';
import AlteredDateTexts from '../functions/alteredRows/alteredDateTexts';


const AlteredRows = async (sheetArray: Array<string[]>) => {

  for (let month = 0; month = 11; month++) {
    var dateYear = 2022

    const date = new Date(dateYear,month,1)
    console.log(month)
    sheetArray.forEach(async (row: Array<string>)=> {

      await AlteredDate(row, date)

      await AlteredDateInHeader(row, date)

      await AlteredDateTexts(row,date)

      await SaveNewWorkBook(sheetArray,month)

    })

  }

}


export default AlteredRows;
