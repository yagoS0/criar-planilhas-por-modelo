import XLSX from 'xlsx-js-style'


export default async function StyleCompras(
  sheet:XLSX.WorkSheet,celulas:string[][],indiceCompra: number[]){

  const letras: number[] = [0,1,2,3,4]
  
  const rowInicio = indiceCompra[0]
  letras.map(colun => {

    if(rowInicio === indiceCompra[0]){

      // Primeira row
      sheet[celulas[rowInicio][colun]].s = {
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
            style: 'thick',
          },
          left: {
            style: 'medium',
          },
        },
      }
        
      //Cabeçalhos
      sheet[celulas[rowInicio-2][colun]].s = {
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
      sheet[celulas[rowInicio-1][colun]].s = {
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

      
    }

    // Borda direita das celulas
    if(colun === 4){

        sheet[celulas[rowInicio-2][colun]].s = {
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
          right: {
            style: 'thick',
          },
          top: {
            style: 'thick',
          },
        
        },
      }

      sheet[celulas[rowInicio+1][colun]].s = {
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
        
          bottom: {
            style: 'thick',
          },
          right: {
            style: 'thick',
          },
        },
      }
      sheet[celulas[rowInicio][colun]].s = {
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
      }
    }
  })
  


  if(indiceCompra[1]){
    const rowFinal = indiceCompra[1]
    letras.map(colun => {
      // Ultima row
      sheet[celulas[rowFinal][colun]].s = {
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
      }
      sheet[celulas[rowFinal+1][colun]].s = {
        font: {
          sz: 12,
          bold: false,
          color: {
            rgb: '000',
          },
        },
        alignment: {
          horizontal: 'right',
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
      

      if(colun === 4){
        sheet[celulas[rowFinal+1][colun]].s = {
          font: {
            sz: 12,
            bold: false,
            color: {
              rgb: '000',
            },
          },
          alignment: {
            horizontal: 'right',
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
        sheet[celulas[rowFinal][colun]].s = {
          font: {
            sz: 12,
            bold: false,
            color: {
              rgb: '000',
            },
          },
          alignment: {
            horizontal: 'right',
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
      }
    })
  }

  return sheet
} 