import XLSX from 'xlsx-js-style'


export default async function StyleReceitaPresumido(
  sheet:XLSX.WorkSheet,celulas:string[][],indiceReceitaPresumido: number[]){

  const letras: number[] = [0,1,2,3,4]

  // Cabeçalhos   
  for (let row = indiceReceitaPresumido[0]-2; row <= indiceReceitaPresumido[1]; row++) {
    letras.map(async (colun)=>{
      if(row === indiceReceitaPresumido[0]-2){
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
    if(row === indiceReceitaPresumido[0]-2){
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
    if(row === indiceReceitaPresumido[0]-1){
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
    }
  }
)}

  for (let row = indiceReceitaPresumido[0]; row <= indiceReceitaPresumido[1]; row++) {

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
        // penultima linha linha da tabela
        if(row === indiceReceitaPresumido[1]){
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
        // Ultima linha
        if(row === indiceReceitaPresumido[1]){
          sheet[celulas[row+1][colun]].s = {
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
        }
        // pis a pagar
        if(row === indiceReceitaPresumido[1]-2){
          sheet[celulas[row][colun]].s = {
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
        }
          // iss a pagar
          if(row === indiceReceitaPresumido[1]-5){
            sheet[celulas[row][colun]].s = {
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
          }
            // Receita n/mes
        if(row === indiceReceitaPresumido[1]-8){
          sheet[celulas[row][colun]].s = {
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
        }
    })
  }

  return sheet
} 