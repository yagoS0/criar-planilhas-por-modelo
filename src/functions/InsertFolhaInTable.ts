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

 

    // Salario PRO-LABORE
    let c = 2
    const inicioProLab = celulas[numero + 2][4]
    sheet[celulas[numero + c][0]] = {
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
    sheet[celulas[numero + c ][1]] = {
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
    sheet[celulas[numero+ c ][2]] = {
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
    sheet[celulas[numero + c][3]] = {
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
    sheet[celulas[numero+c][4]] = {
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

      // INSS PRO-LABORE
       c = 3 
      sheet[celulas[numero + c][0]] = {
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
      sheet[celulas[numero + c ][1]] = {
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
      sheet[celulas[numero+ c ][2]] = {
        v: "240",
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
      sheet[celulas[numero + c][3]] = {
        v: `VR REF INSS S/PRO LAB  ${dateText}`,
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
      sheet[celulas[numero+c][4]] = {
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

         // IRPF PRO-LABORE
     c = 4
     const finalProLab = celulas[numero + 4][4]
    sheet[celulas[numero + c][0]] = {
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
    sheet[celulas[numero + c ][1]] = {
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
    sheet[celulas[numero+ c ][2]] = {
      v: "257",
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
    sheet[celulas[numero + c][3]] = {
      v: `VR REF IRFF S/PRO LAB FP ${dateText}`,
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
    sheet[celulas[numero+c][4]] = {
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

     //  PRO-LABORE LIQUIDO
     c = 5
    sheet[celulas[numero + c][0]] = {
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
            style: 'thick',
          },
          bottom: {
            style: 'thick',
          },
          right: {
            style: 'medium',
          },
          left: {
            style: 'medium',
          },
      },
    }}
    sheet[celulas[numero + c ][1]] = {
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
            style: 'medium',
          },
          left: {
            style: 'medium',
          },
      },
    }}
    sheet[celulas[numero+ c ][2]] = {
      v: "233",
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
            style: 'medium',
          },
          left: {
            style: 'medium',
          },
      },
    }}
    sheet[celulas[numero + c][3]] = {
      v: `VR REF PRO LAB LIQUIDO FP ${dateText}`,
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
            style: 'medium',
          },
          left: {
            style: 'medium',
          },
      },
    }}
    sheet[celulas[numero+c][4]] = {
      f: `=SOMA(${inicioProLab}:${finalProLab})`,
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
            style: 'medium',
          },
      },
    }}
    ///////

    // salario folha
    c = 6
    const inicioFolha = celulas[numero + 6][4]
    sheet[celulas[numero + c][0]] = {
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
    sheet[celulas[numero + c ][1]] = {
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
    sheet[celulas[numero+ c ][2]] = {
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
    sheet[celulas[numero + c][3]] = {
      v: `VR REF SAL FP ${dateText}`,
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
    sheet[celulas[numero+c][4]] = {
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

      // INSS FOLHA
      c = 7
      sheet[celulas[numero + c][0]] = {
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
      sheet[celulas[numero + c ][1]] = {
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
      sheet[celulas[numero+ c ][2]] = {
        v: "240",
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
      sheet[celulas[numero + c][3]] = {
        v: `VR REF INSS FP ${dateText}`,
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
      sheet[celulas[numero+c][4]] = {
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
      // SAL FAMILIA FOLHA
      c = 8
      sheet[celulas[numero + c][0]] = {
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
      sheet[celulas[numero + c ][1]] = {
        v: "240",
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
      sheet[celulas[numero+ c ][2]] = {
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
      sheet[celulas[numero + c][3]] = {
        v: `VR REF SAL FAMILIA FP ${dateText}`,
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
      sheet[celulas[numero+c][4]] = {
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

      // FERIAS + 1/3 FOLHA
      c = 9
      sheet[celulas[numero + c][0]] = {
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
      sheet[celulas[numero + c ][1]] = {
        v: "444",
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
      sheet[celulas[numero+ c ][2]] = {
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
      sheet[celulas[numero + c][3]] = {
        v: `VR REF FERIAS + 1/3 FP ${dateText}`,
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
      sheet[celulas[numero+c][4]] = {
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
     
      
       // AVISO PREVIO INDENIZADO FOLHA
       c = 10
       sheet[celulas[numero + c][0]] = {
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
       sheet[celulas[numero + c ][1]] = {
         v: "451",
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
       sheet[celulas[numero+ c ][2]] = {
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
       sheet[celulas[numero + c][3]] = {
         v: `VR REF AVISO PREVIO INDENIZADO FP ${dateText}`,
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
       sheet[celulas[numero+c][4]] = {
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

         // 13 PROP RESCISAO FOLHA
         c = 11
         sheet[celulas[numero + c][0]] = {
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
         sheet[celulas[numero + c ][1]] = {
           v: "445",
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
         sheet[celulas[numero+ c ][2]] = {
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
         sheet[celulas[numero + c][3]] = {
           v: `VR REF 13 PROP RESCISAO FP ${dateText}`,
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
         sheet[celulas[numero+c][4]] = {
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


              // FERIAS + 1/3 FOLHA
              c = 12
              sheet[celulas[numero + c][0]] = {
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
              sheet[celulas[numero + c ][1]] = {
                v: "444",
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
              sheet[celulas[numero+ c ][2]] = {
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
              sheet[celulas[numero + c][3]] = {
                v: `VR REF FERIAS + 1/3 FP ${dateText}`,
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
              sheet[celulas[numero+c][4]] = {
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

              
              // FERIAS LIQUIDAS FOLHA
              c = 13
              sheet[celulas[numero + c][0]] = {
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
              sheet[celulas[numero + c ][1]] = {
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
              sheet[celulas[numero+ c ][2]] = {
                v: "234",
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
              sheet[celulas[numero + c][3]] = {
                v: `VR REF FERIAS LIQUIDAS FP ${dateText}`,
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
              sheet[celulas[numero+c][4]] = {
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
              // IRRF  FOLHA FOLHA
              c = 14
              sheet[celulas[numero + c][0]] = {
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
              sheet[celulas[numero + c ][1]] = {
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
              sheet[celulas[numero+ c ][2]] = {
                v: "257",
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
              sheet[celulas[numero + c][3]] = {
                v: `VR REF IRRF S/FOLHA FP ${dateText}`,
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
              sheet[celulas[numero+c][4]] = {
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
      // CONST ASSISTENCIAL  FOLHA FOLHA
      c = 15
      sheet[celulas[numero + c][0]] = {
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
      sheet[celulas[numero + c ][1]] = {
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
      sheet[celulas[numero+ c ][2]] = {
        v: "245",
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
      sheet[celulas[numero + c][3]] = {
        v: `VR REF CONT ASSISTENCIAL FP ${dateText}`,
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
      sheet[celulas[numero+c][4]] = {
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
      // RESCISAO LIQUIDA  FOLHA FOLHA
      c = 16
      sheet[celulas[numero + c][0]] = {
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
      sheet[celulas[numero + c ][1]] = {
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
      sheet[celulas[numero+ c ][2]] = {
        v: "243",
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
      sheet[celulas[numero + c][3]] = {
        v: `VR REF RESCISAO LIQ FP ${dateText}`,
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
      sheet[celulas[numero+c][4]] = {
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
    // VALE ALIMENTACAO  FOLHA FOLHA
    c = 17
    sheet[celulas[numero + c][0]] = {
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
    sheet[celulas[numero + c ][1]] = {
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
    sheet[celulas[numero+ c ][2]] = {
      v: "437",
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
    sheet[celulas[numero + c][3]] = {
      v: `VR REF VALE ALIMENTACAO FP ${dateText}`,
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
    sheet[celulas[numero+c][4]] = {
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
      // VALE TRASNPORT  FOLHA FOLHA
      let finalFolha = celulas[numero + 18][0]
      
      c = 18
      
      sheet[celulas[numero + c][0]] = {
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
      sheet[celulas[numero + c ][1]] = {
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
      sheet[celulas[numero+ c ][2]] = {
        v: "436",
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
      sheet[celulas[numero + c][3]] = {
        v: `VR REF VALE TRANSPORTE FP ${dateText}`,
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
      sheet[celulas[numero+c][4]] = {
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
      // sal liquido LIQUIDO
     c = 19
     sheet[celulas[numero + c][0]] = {
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
             style: 'thick',
           },
           bottom: {
             style: 'thick',
           },
           right: {
             style: 'medium',
           },
           left: {
             style: 'medium',
           },
       },
     }}
     sheet[celulas[numero + c ][1]] = {
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
             style: 'medium',
           },
           left: {
             style: 'medium',
           },
       },
     }}
     sheet[celulas[numero+ c ][2]] = {
       v: "232",
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
             style: 'medium',
           },
           left: {
             style: 'medium',
           },
       },
     }}
     sheet[celulas[numero + c][3]] = {
       v: `VR REF SAL LIQ FP ${dateText}`,
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
             style: 'medium',
           },
           left: {
             style: 'medium',
           },
       },
     }}
     sheet[celulas[numero+c][4]] = {
       f: `=SOMA(${inicioFolha}}:${finalFolha})`,
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
             style: 'medium',
           },
       },
     }}
     ///////
  

  
  return sheet
} 


