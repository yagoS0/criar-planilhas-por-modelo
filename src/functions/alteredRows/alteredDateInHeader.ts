

export default async function AlteredDateInHeader(row: Array<string>, date: Date){

  var month = date.toLocaleString('pt-BR', { month: 'long'});
  var year = date.toLocaleString('pt-BR', { year: 'numeric'});

  if (row[0] === '') {

    const dataDeTroca = row[3].split("/")
    const searchYear = /(\d{4})/.exec(row[3]);

    if(dataDeTroca.length === 2) {

      row[3] = row[3].replace(dataDeTroca[0], month.toUpperCase())
      console.log(row[3])

      row[3] = row[3].replace(searchYear? searchYear[0] : '' , year)
      console.log(row[3])

    }
  }
}
