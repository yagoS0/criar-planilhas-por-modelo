export default async function DateFormat(month: number){
    const newDate = new Date()
    
    const year = newDate.getFullYear()

    let dateMonth = new Date(year,month,1).getMonth()
    dateMonth += 1
    
    const getLastDay = new Date(year,dateMonth, 0)
    const lastDay = getLastDay.getDate()

    const date = {year, dateMonth, lastDay}

    return date
}