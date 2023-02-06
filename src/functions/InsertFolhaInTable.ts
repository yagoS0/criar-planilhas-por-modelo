import XLSX from 'xlsx-js-style'


export default async function InsertFolhaInTable(
  sheet:XLSX.WorkSheet,celulas:string[][],indice: number[],coluns: {
    dataValue: string;
    debitoValue: number;
    creditoValue: number;
    textValue: string;
    value: string;
   }[] ){


    const data = coluns[3].dataValue

    const splitDateText = data.split('/')


    console.log(splitDateText)
    const dateText = `${splitDateText[1]}/${splitDateText[2]}`
    const year = splitDateText[2]

    const indiceMes =  parseInt(splitDateText[1])-1

    const arrayMonthText  = [
      'JANEIRO',
      'FEVEREIRO',
      'MARCO',
      'ABRIL',
      'MAIO',
      'JUNHO',
      'JULHO',
      'AGOSTO',
      'SETEMBRO',
      'OUTUBRO',
      'NOVEMBRO',
      'DEZEMBRO',

    ]


    const letras: number[] = [0,1,2,3,4]

    let numero = indice[0] + 5
    console.log(numero)

    // cabeçalho
    sheet[celulas[numero][0]] = {
      v: " ",
      s:{
        font: {
          sz: 12,
          bold: false,
          color: {
            rgb: '000',
          },
        },
      alignment: {
          horizontal: 'center',
        },
      border: {
          top: {
            style: 'thick',
          },
          bottom: {
            style: 'thick',
          },
          right: {
            style: 'thick',
          },
          left: {
            style: 'thick',
          },
      },
    }}
    sheet[celulas[numero][1]] = {
      v: " ",
      s:{
        font: {
          sz: 12,
          bold: false,
          color: {
            rgb: '000',
          },
        },
      alignment: {
          horizontal: 'center',
        },
      border: {
          top: {
            style: 'thick',
          },
          bottom: {
            style: 'thick',
          },
          right: {
            style: 'thick',
          },
          left: {
            style: 'thick',
          },
      },
    }}
    sheet[celulas[numero][2]] = {
      v: " ",
      s:{
        font: {
          sz: 12,
          bold: false,
          color: {
            rgb: '000',
          },
        },
      alignment: {
          horizontal: 'left',
        },
      border: {
          top: {
            style: 'thick',
          },
          bottom: {
            style: 'thick',
          },
          right: {
            style: 'thick',
          },
          left: {
            style: 'thick',
          },
      },
    }}
    sheet[celulas[numero][3]] = {
      v: `${arrayMonthText[indiceMes]}/${year} FOLHA`,
      s:{
        font: {
          sz: 12,
          bold: true,
          color: {
            rgb: '000',
          },
        },
      alignment: {
          horizontal: 'center',
        },
      border: {
          top: {
            style: 'thick',
          },
          bottom: {
            style: 'thick',
          },
          right: {
            style: 'thick',
          },
          left: {
            style: 'thick',
          },
      },
    }}
    sheet[celulas[numero][4]] = {
      v:" ",
      s:{
        font: {
          sz: 12,
          bold: true,
          color: {
            rgb: '000',
          },
        },
      alignment: {
          horizontal: 'center',
        },
      border: {
          top: {
            style: 'thick',
          },
          bottom: {
            style: 'thick',
          },
          right: {
            style: 'thick',
          },
          left: {
            style: 'thick',
          },
      },
    }}

    

 // cabeçalho 2
 sheet[celulas[numero+1][0]] = {
  v: "DATA",
  s:{
    font: {
      sz: 12,
      bold: true,
      color: {
        rgb: '000',
      },
    },
  alignment: {
      horizontal: 'center',
    },
  border: {
      top: {
        style: 'thick',
      },
      bottom: {
        style: 'thick',
      },
      right: {
        style: 'thick',
      },
      left: {
        style: 'thick',
      },
  },
}}
sheet[celulas[numero+1][1]] = {
  v: "DEBITO",
  s:{
    font: {
      sz: 12,
      bold: true,
      color: {
        rgb: '000',
      },
    },
  alignment: {
      horizontal: 'center',
    },
  border: {
      top: {
        style: 'thick',
      },
      bottom: {
        style: 'thick',
      },
      right: {
        style: 'thick',
      },
      left: {
        style: 'thick',
      },
  },
}}
sheet[celulas[numero+1][2]] = {
  v: "CREDITO",
  s:{
    font: {
      sz: 12,
      bold: true,
      color: {
        rgb: '000',
      },
    },
  alignment: {
      horizontal: 'center',
    },
  border: {
      top: {
        style: 'thick',
      },
      bottom: {
        style: 'thick',
      },
      right: {
        style: 'thick',
      },
      left: {
        style: 'thick',
      },
  },
}}
sheet[celulas[numero+1][3]] = {
  v: `HISTORICO`,
  s:{
    font: {
      sz: 12,
      bold: true,
      color: {
        rgb: '000',
      },
    },
  alignment: {
      horizontal: 'center',
    },
  border: {
      top: {
        style: 'thick',
      },
      bottom: {
        style: 'thick',
      },
      right: {
        style: 'thick',
      },
      left: {
        style: 'thick',
      },
  },
}}
sheet[celulas[numero+1][4]] = {
  v:"VALOR",
  s:{
    font: {
      sz: 12,
      bold: true,
      color: {
        rgb: '000',
      },
    },
  alignment: {
      horizontal: 'center',
    },
  border: {
      top: {
        style: 'thick',
      },
      bottom: {
        style: 'thick',
      },
      right: {
        style: 'thick',
      },
      left: {
        style: 'thick',
      },
  },
}}

    // Salario pro labore
    sheet[celulas[numero+2][0]] = {
      v: data,
      s:{
        font: {
          sz: 12,
          bold: false,
          color: {
            rgb: '000',
          },
        },
      alignment: {
          horizontal: 'center',
        },
      border: {
          top: {
            style: 'medium',
          },
          bottom: {
            style: 'medium',
          },
          right: {
            style: 'medium',
          },
          left: {
            style: 'medium',
          },
      },
    }}
    sheet[celulas[numero+2][1]] = {
      v: "428",
      s:{
        font: {
          sz: 12,
          bold: false,
          color: {
            rgb: '000',
          },
        },
      alignment: {
          horizontal: 'center',
        },
      border: {
          top: {
            style: 'medium',
          },
          bottom: {
            style: 'medium',
          },
          right: {
            style: 'medium',
          },
          left: {
            style: 'medium',
          },
      },
    }}
    sheet[celulas[numero+2][2]] = {
      v: " ",
      s:{
        font: {
          sz: 12,
          bold: false,
          color: {
            rgb: '000',
          },
        },
      alignment: {
          horizontal: 'left',
        },
      border: {
          top: {
            style: 'medium',
          },
          bottom: {
            style: 'medium',
          },
          right: {
            style: 'medium',
          },
          left: {
            style: 'medium',
          },
      },
    }}
    sheet[celulas[numero+2][3]] = {
      v: `VR REF SAL PRO-LAB FP ${dateText}`,
      s:{
        font: {
          sz: 12,
          bold: false,
          color: {
            rgb: '000',
          },
        },
      alignment: {
          horizontal: 'left',
        },
      border: {
          top: {
            style: 'medium',
          },
          bottom: {
            style: 'medium',
          },
          right: {
            style: 'medium',
          },
          left: {
            style: 'medium',
          },
      },
    }}
    sheet[celulas[numero+2][4]] = {
      v:" ",
      s:{
        font: {
          sz: 12,
          bold: true,
          color: {
            rgb: '000',
          },
        },
      alignment: {
          horizontal: 'center',
        },
      border: {
          top: {
            style: 'medium',
          },
          bottom: {
            style: 'medium',
          },
          right: {
            style: 'thick',
          },
          left: {
            style: 'medium',
          },
      },
    }}
    ///////
  
  return sheet
} 


