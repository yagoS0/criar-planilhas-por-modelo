export default async function StyleSheet(sheet) {
  sheet["A1"].s = {
      font: { sz: 50, color: { rgb: '00f' } },
      border: { top: { style: 'bold' }, bottom: { style: 'bold' } }
    }

    console.log('Entrando no style')
    return sheet
}