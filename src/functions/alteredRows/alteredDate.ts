

export default async function AlteredDate
(row: Array<string>, date: {
  year:number,
  dateMonth: number
  lastDay:number
}){


  const dataDeTroca = row[0].split("/")
  const searchYear = /(\d{4})/.exec(dataDeTroca[2]);

  if(searchYear?.index === 0){
    const dateText =`${date.lastDay}/${date.dateMonth}/${date.year + 1}`

    row[0] = dateText
    return row
  }
  // const row1 = parseInt(row[1])
  // const row2 = parseInt(row[2])
}
