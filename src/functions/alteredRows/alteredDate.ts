import {format} from 'date-fns'


export default async function AlteredDate(row: Array<string>, date: Date){

  var ultimoDia = format( new Date(date.getFullYear(), date.getMonth() + 1, 0), 'dd/MM/yyyy');

  const row1 = parseInt(row[1])
  const row2 = parseInt(row[2])

  if(row1 || row2) {
    row[0] =  ultimoDia
  }
}
