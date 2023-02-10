import XLSX from 'xlsx-js-style'


export default async function InsertFolhaInTable(
  sheet:XLSX.WorkSheet,celulas:string[][],indice: number[],coluns: {
    dataValue: string;
    debitoValue: number;
    creditoValue: number;
    textValue: string;
    value: string;
   }[] ){

    let campo 
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

      let numero = indice[0] + 5
      const A = 0
      const B = 1
      const C = 2 
      const D = 3
      const E = 4


    const styleCabecalho = {
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
  } 
    const styleCabecalhoCanto ={
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
    },
  } 
    const cabeçalho = [
      {v: " ", s: styleCabecalho},
      {v: " ", s: styleCabecalho},
      {v: " ", s: styleCabecalho},
      {v: `${arrayMonthText[indiceMes]}/${year} FOLHA `, s: styleCabecalho},
      {v: " ", s: styleCabecalhoCanto},
    ];
    campo = cabeçalho
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]



    const styleTablea = {
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
  }
    const table = [
      {v: "DATA", s: styleTablea},
      {v: "DEBITO", s: styleTablea},
      {v: "CREDITO", s: styleTablea},
      {v: "HISTORICO", s: styleTablea},
      {v: "VALOR", s: styleTablea},
    ];
    campo = table
    numero += 1
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]


    const stylePadrao = {
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
    }
    const stylePadraoCanto = {
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
          style: 'thick',
        },
    },
    }
    const stylePadraoFinal = {
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
        style: 'thick',
      },
      right: {
        style: 'medium',
      },
      left: {
        style: 'medium',
      },
  },
    }
    const stylePadraoFinalCanto = {
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
      style: 'thick',
    },
    right: {
      style: 'medium',
    },
    left: {
      style: 'thick',
    },
},
    }


    const salarioProLab = [
      {v: data, s: stylePadrao},
      {v: "428", s: stylePadrao},
      {v: " ", s: stylePadrao},
      {v: `VR REF PRO LAB FP ${dateText} `, s: stylePadrao},
      {v: " ", s: stylePadraoCanto},
    ];
    campo = salarioProLab
    numero += 1
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]


    const inssProLab = [
      {v: data, s: stylePadrao},
      {v: " ", s: stylePadrao},
      {v: "240", s: stylePadrao},
      {v: `VR REF INSS S/PRO LAB FP ${dateText} `, s: stylePadrao},
      {v: " ", s: stylePadraoCanto},
    ];
    campo = inssProLab
    numero += 1
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]

  
  
    
/*
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
      const inssFolha = celulas[numero + 7][0]
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
             style: 'medium',
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
             style: 'medium',
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
             style: 'medium',
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
             style: 'medium',
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
             style: 'medium',
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

     //PULA LINHA
      c = 20

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
    
      border: {
          right: {
            style: 'thick',
          },
          
      },
    }}
    ///////

    // FGTS FOLHA 
    c = 21
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
      v: "447",
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
      v: "242",
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
      v: `VR REF FGTS ${dateText}`,
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
            style: 'thick',
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

    // sal materniadade
    const salMaternidade = celulas[numero + 22][0]
    c = 22
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
      v: "83",
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
      v: `VR REF SAL MATERNIDADE FP ${dateText}`,
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
  

      // INSS parte empresa
      const inssEmpresa = celulas[numero + 23][0]
      
      c = 23
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
        v: "446",
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
        v: `VR REF INSS EMP ${dateText}`,
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

     // compensação INSS ref sal maternidade
      const compInssMaternidade = celulas[numero + 24][0]
     c = 24
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
       v: "83",
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
       v: `VR REF COMPENSACAO INSS REF SAL MTERNIDADE ${dateText}`,
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
   

     // INSS A RECOLHER
     c = 25
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
             style: 'thick',
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
             style: 'thick',
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
             style: 'thick',
           },
          
       },
     }}
     sheet[celulas[numero + c][3]] = {
       v: `INSS A RECOLHER`,
       s:{
         font: {
           sz: 12,
           bold: true,
           color: {
             rgb: '000',
           },
         },
       alignment: {
           horizontal: 'right',
         },
       border: {
           top: {
             style: 'medium',
           },
           bottom: {
             style: 'thick',
           },
           right: {
             style: 'medium',
           },
          
       },
     }}

    
        sheet[celulas[numero+c][4]] = {
          f: `=-${inssFolha}-${compInssMaternidade}+${inssEmpresa}`, 
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
       }
    } */

     
    
     ///////
    

  
  return sheet
} 


