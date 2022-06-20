import SaveNewWorkBook from '../service/SaveNewWorkBook';
import AlteredDate from '../functions/alteredRows/alteredDate';
import AlteredDateInHeader from '../functions/alteredRows/alteredDateInHeader';
import AlteredDateTexts from '../functions/alteredRows/alteredDateTexts';


const AlteredRows = async (sheetArray: Array<string[]>) => {
  for (let month = 0; month <= 11; month++) {

    const fullYear  = new Date()
    const date = new Date(fullYear.getFullYear(),month,1)


    sheetArray.forEach(async (row: Array<string>)=> {

      await AlteredDateInHeader(row, date)

      await AlteredDate(row, date)


      await AlteredDateTexts(row,date)


    })

    await SaveNewWorkBook(sheetArray,month)
  }

}


export default AlteredRows;
