import XLSX from 'xlsx-js-style'


export default async function StyleReceitaVenda(
  sheet:XLSX.WorkSheet,celulas:string[][],indiceReceitaVenda: number[]){
    const letras: number[] = [0,1,2,3,4]

    const rowInicio = indiceReceitaVenda[0]
    const rowFinal = indiceReceitaVenda[1]
  
    // Cabeçalhos   
    letras.map(async (colun)=>{
      if(rowInicio === indiceReceitaVenda[0]){
  
        // Rows do meio 
        for (let index = 0; index <= rowFinal; index++) {
          
          sheet[celulas[index][colun]].s = {
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
        }
        sheet[celulas[rowFinal-1][colun]].s = {
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
            rigth: {
              style: 'thick',
            },
            left: {
              style: 'thick',
            },
          },
        }
          
        //Cabeçalhos
        sheet[celulas[rowInicio-3][colun]].s = {
          font: {
            sz: 16,
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
  
        if(colun === 4){
          sheet[celulas[rowInicio-3][colun]].s = {
            font: {
              sz: 16,
              bold: true,
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
           
            },
          }
  
          sheet[celulas[rowFinal-1][colun]].s = {
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
              style: 'thick',
            },
            left: {
              style: 'thick',
            },
          },
        }
      }
        
      }
  
    })
   

  return sheet
} 