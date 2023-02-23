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
    numero ++
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]


    const stylePadrao = {
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
          style: 'thick',
        },
        left: {
          style: 'medium',
        },
    },
    }
    const stylePadraoFinal = {
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
    }
    const stylePadraoFinalCanto = {
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
            style: 'thick',
          },
      },
    }

    const stylePadraoDebitoCreditoFinal = {
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
    }
    const stylePadraoDebitoCredito = {
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
    }


    const salarioProLab = [
      {v: data, s: stylePadrao},
      {v: "428", s: stylePadraoDebitoCredito},
      {v: " ", s: stylePadraoDebitoCredito},
      {v: `VR REF PRO LAB FP ${dateText} `, s: stylePadrao},
      {v: " ", s: stylePadraoCanto},
    ];
    campo = salarioProLab
    numero ++
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]


    const inssProLab = [
      {v: data, s: stylePadrao},
      {v: " ", s: stylePadraoDebitoCredito},
      {v: "240", s: stylePadraoDebitoCredito},
      {v: `VR REF INSS S/PRO LAB FP ${dateText} `, s: stylePadrao},
      {v: " ", s: stylePadraoCanto},
    ];
    campo = inssProLab
    numero ++
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]


    const irrfProLab = [
      {v: data, s: stylePadrao},
      {v: " ", s: stylePadraoDebitoCredito},
      {v: "257", s: stylePadraoDebitoCredito},
      {v: `VR REF IRRF S/PRO LAB FP ${dateText} `, s: stylePadrao},
      {v: " ", s: stylePadraoCanto},
    ];
    campo = irrfProLab
    numero ++
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]

    const proLabLiq = [
      {v: data, s: stylePadraoFinal},
      {v: " ", s: stylePadraoDebitoCreditoFinal},
      {v: "233", s: stylePadraoDebitoCreditoFinal},
      {v: `VR REF PRO LAB LIQ FP ${dateText} `, s: stylePadraoFinal},
      {v: " ", s: stylePadraoFinalCanto},
    ];
    campo = proLabLiq
    numero ++
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]

    const salFolha = [
      {v: data, s: stylePadrao},
      {v: "428", s: stylePadraoDebitoCredito},
      {v: " ", s: stylePadraoDebitoCredito},
      {v: `VR REF SAL FP ${dateText} `, s: stylePadrao},
      {v: " ", s: stylePadraoCanto},
    ];
    campo = salFolha
    numero ++
    const primeriaCelFolha = numero
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]

    const inssFolha = [
      {v: data, s: stylePadrao},
      {v: " ", s: stylePadraoDebitoCredito},
      {v: "240", s: stylePadraoDebitoCredito},
      {v: `VR REF INSS FP ${dateText} `, s: stylePadrao},
      {v: " ", s: stylePadraoCanto},
    ];
    campo = inssFolha
    numero ++
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]

    const salFamiliaFolha = [
      {v: data, s: stylePadrao},
      {v: "240", s: stylePadraoDebitoCredito},
      {v: " ", s: stylePadraoDebitoCredito},
      {v: `VR REF SAL FAMILIA FP ${dateText} `, s: stylePadrao},
      {v: " ", s: stylePadraoCanto},
    ];
    campo = salFamiliaFolha
    numero ++
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]

    const feriaUmTerco = [
      {v: data, s: stylePadrao},
      {v: "444", s: stylePadraoDebitoCredito},
      {v: " ", s: stylePadraoDebitoCredito},
      {v: `VR REF FERIAS + 1/3 FP ${dateText} `, s: stylePadrao},
      {v: " ", s: stylePadraoCanto},
    ];
    campo = feriaUmTerco
    numero ++
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]


    const avisoPrevio = [
      {v: data, s: stylePadrao},
      {v: "451", s: stylePadraoDebitoCredito},
      {v: " ", s: stylePadraoDebitoCredito},
      {v: `VR REF AVISO PREVIO INDENIZADO FP ${dateText} `, s: stylePadrao},
      {v: " ", s: stylePadraoCanto},
    ];
    campo = avisoPrevio
    numero ++
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]

    const decimoTerceiroRescisao = [
      {v: data, s: stylePadrao},
      {v: "445", s: stylePadraoDebitoCredito},
      {v: " ", s: stylePadraoDebitoCredito},
      {v: `VR REF 13 PROP RESCISAO  FP ${dateText} `, s: stylePadrao},
      {v: " ", s: stylePadraoCanto},
    ];
    campo = decimoTerceiroRescisao
    numero ++
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]


    const feriasMaisUmTerco = [
      {v: data, s: stylePadrao},
      {v: "444", s: stylePadraoDebitoCredito},
      {v: " ", s: stylePadraoDebitoCredito},
      {v: `VR REF FERIAS + 1/3 FP PROP ${dateText} `, s: stylePadrao},
      {v: " ", s: stylePadraoCanto},
    ];
    campo = feriasMaisUmTerco
    numero ++
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]

    const feriasLiqFolha = [
      {v: data, s: stylePadrao},
      {v: " ", s: stylePadraoDebitoCredito},
      {v: "234", s: stylePadraoDebitoCredito},
      {v: `VR REF FERIAS LIQ FP ${dateText} `, s: stylePadrao},
      {v: " ", s: stylePadraoCanto},
    ];
    campo = feriasLiqFolha
    numero ++
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]

    const irrfFolha = [
      {v: data, s: stylePadrao},
      {v: " ", s: stylePadraoDebitoCredito},
      {v: "257", s: stylePadraoDebitoCredito},
      {v: `VR REF IRRF S/FOLHA FP ${dateText} `, s: stylePadrao},
      {v: " ", s: stylePadraoCanto},
    ];
    campo = irrfFolha
    numero ++
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]


    const contAssistencial = [
      {v: data, s: stylePadrao},
      {v: " ", s: stylePadraoDebitoCredito},
      {v: "245", s: stylePadraoDebitoCredito},
      {v: `VR REF CONT ASSISTENCIAL FP ${dateText} `, s: stylePadrao},
      {v: " ", s: stylePadraoCanto},
    ];
    campo = contAssistencial
    numero ++
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]

    const rescisaoLiq = [
      {v: data, s: stylePadrao},
      {v: " ", s: stylePadraoDebitoCredito},
      {v: "243", s: stylePadraoDebitoCredito},
      {v: `VR REF RESCISAO LIQ FP ${dateText} `, s: stylePadrao},
      {v: " ", s: stylePadraoCanto},
    ];
    campo = rescisaoLiq
    numero ++
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]


    const valeAlimentacao = [
      {v: data, s: stylePadrao},
      {v: " ", s: stylePadraoDebitoCredito},
      {v: "437", s: stylePadraoDebitoCredito},
      {v: `VR REF VALE ALIMENTACAO FP ${dateText} `, s: stylePadrao},
      {v: " ", s: stylePadraoCanto},
    ];
    campo = valeAlimentacao
    numero ++
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]

    const valeTransporte = [
      {v: data, s: stylePadrao},
      {v: " ", s: stylePadraoDebitoCredito},
      {v: "436", s: stylePadraoDebitoCredito},
      {v: `VR REF VALE TRANSPORTE FP ${dateText} `, s: stylePadrao},
      {v: " ", s: stylePadraoCanto},
    ];
    campo = valeTransporte
    numero ++
    const UltimaCelFolha = numero
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]


    const salLiqFolha = [
      {v: data, s: stylePadraoFinal},
      {v: " ", s: stylePadraoDebitoCreditoFinal},
      {v: "232", s: stylePadraoDebitoCreditoFinal},
      {v: `VR REF SAL LIQ FP ${dateText} `, s: stylePadraoFinal},
      {v: " ", s: stylePadraoFinalCanto, f: `=SOMA(${primeriaCelFolha}:${UltimaCelFolha})`},
    ];
    campo = salLiqFolha
    numero ++
    sheet[celulas[numero][A]] = campo[A]
    sheet[celulas[numero][B]] = campo[B]
    sheet[celulas[numero][C]] = campo[C]
    sheet[celulas[numero][D]] = campo[D]
    sheet[celulas[numero][E]] = campo[E]

  
  return sheet
} 


