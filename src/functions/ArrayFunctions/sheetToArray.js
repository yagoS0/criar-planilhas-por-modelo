import * as XLSX from 'xlsx'

export  default async function sheetToArray(sheet){

  var range = XLSX.utils.decode_range(sheet['!ref']);
  range.s.c = 0; // 0 == XLSX.utils.decode_col("A")
  range.e.c = 5; // 6 == XLSX.utils.decode_col("G")
  var new_range = XLSX.utils.encode_range(range);

  let sheetArray  =  XLSX.utils.sheet_to_json(
    sheet, {defval: "" , range: new_range}
    );
  sheetArray =  sheetArray.map((obj) =>Object.values(obj));

  return sheetArray
}
