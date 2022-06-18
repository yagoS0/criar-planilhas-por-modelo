
export default async function AlteredDateTexts(row: Array<string>, date: Date){

  var monthNumber = date.toLocaleString('pt-BR', { month: '2-digit'});
  var year = date.toLocaleString('pt-BR', { year: 'numeric'});
  var searchDate = /(\d{2})[-.\/](\d{4})/.exec(row[3]);

  if(searchDate){
    row[3] = row[3].replace(searchDate[0], `${monthNumber}/${year}`)
  }

}
