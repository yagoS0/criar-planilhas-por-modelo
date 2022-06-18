

export default async function AlteredDateInHeader(row: Array<string>, date: Date){
  var month = date.toLocaleString('pt-BR', { month: 'long'});

  if (row[0] === '') {
    const dataDeTroca = row[3].split("/")
    if(dataDeTroca.length === 2) {
      row[3] = row[3].replace(dataDeTroca[0], month.toUpperCase())
    }
  }
}
