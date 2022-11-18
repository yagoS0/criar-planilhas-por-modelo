

export default async function 
AlteredDateInHeaderAndText(row: Array<string>){


    const separatedBySlash = row[3].split("/")
    if(separatedBySlash.length === 2){

    const dateDeTroca = separatedBySlash[1].split(' ')
    
    const searchYear = /(\d{4})/.exec(dateDeTroca[0]);

    const newYear = parseInt(dateDeTroca[0]) + 1

    if( searchYear ){
      if(separatedBySlash.length === 2) {
        row[3] = row[3].replace(dateDeTroca[0], newYear.toString())        
        // row[3] = row[3].replace(searchYear? searchYear[0] : '', date.year.toString())
    }
   }
  }

  // console.log(row)
}
