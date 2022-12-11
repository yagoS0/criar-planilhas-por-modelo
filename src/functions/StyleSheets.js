export default async function StyleSheets(sheet) {

  for (let i = 1; i <= 20; i++) {
   
  sheet['!cols'] = [
    { wpx: 62 },
    { wpx: 53 },
    { wpx: 53 },
    { wpx: 320 },
    { wpx: 60 },
  ]
  sheet['!rows'] = [
    { hpx: 35 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
    { hpx: 18 },
  ]

  var A = `A${i}`
  var B = `B${i}`
  var C = `C${i}`
  var D = `D${i}`
  var E = `E${i}`
  // Titulo
  sheet[A] = {
    s:{
      font: {
        sz: 16,
        color: {
          rgb: '000',
        },
      },
      border: {
        top: {
          style: 'thick',
        },
        bottom: {
          style: 'thick',
        },
      },
    }
  }
  sheet[B] = {
    s:{
      font: {
        sz: 16,
        color: {
          rgb: '000',
        },
      },
      border: {
        top: {
          style: 'thick',
        },
        bottom: {
          style: 'thick',
        },
      },
    }
  }
  sheet[C] = {
    s:{
      font: {
        sz: 16,
        color: {
          rgb: '000',
        },
      },
      border: {
        top: {
          style: 'thick',
        },
        bottom: {
          style: 'thick',
        },
      },
    }
  }
  sheet[D] = {
    s:{
      font: {
        sz: 16,
        color: {
          rgb: '000',
        },
      },
      border: {
        top: {
          style: 'thick',
        },
        bottom: {
          style: 'thick',
        },
      },
    }
  }
  sheet[E] = {
    s:{
      font: {
        sz: 16,
        color: {
          rgb: '000',
        },
      },
      border: {
        top: {
          style: 'thick',
        },
        bottom: {
          style: 'thick',
        },
      },
    }
  }
  
}



  return sheet
}
