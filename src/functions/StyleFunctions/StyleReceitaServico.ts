import XLSX from 'xlsx-js-style'


export default async function StyleReceitaServico(
  sheet:XLSX.WorkSheet,celulas:string[][],indiceReceitaServiço: number[]){

  const letras: number[] = [0,1,2,3,4]

  // Cabeçalhos   
  for (let row = indiceReceitaServiço[0]-2; row <= indiceReceitaServiço[1]; row++) {
  
    letras.map(async (colun)=>{

    if(row === indiceReceitaServiço[0]-2){
      sheet[celulas[row][colun]].s = {
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
    if(row === indiceReceitaServiço[0]-1){
      sheet[celulas[row][colun]].s = {
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
    ////
    // Bold Canto direito
    if(colun === 4){
      sheet[celulas[row][colun]].s = {
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
  }
)}

  for (let row = indiceReceitaServiço[0]; row <= indiceReceitaServiço[1]; row++) {

    letras.map(async (colun)=>{

    // Meio da tabela
        sheet[celulas[row][colun]].s = {
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
        }
        // Ultima linha da tabela
        if(row === indiceReceitaServiço[1]){
          sheet[celulas[row][colun]].s = {
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
                style: 'thick',
              },
              left: {
                style: 'medium',
              },
            },
          }
        }
        // Penultima linha da tabela
        if(row === indiceReceitaServiço[1]-1){
          sheet[celulas[row][colun]].s = {
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
        }
        if(colun === 4){
          sheet[celulas[row][colun]].s = {
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
    })
  }

  return sheet
} 