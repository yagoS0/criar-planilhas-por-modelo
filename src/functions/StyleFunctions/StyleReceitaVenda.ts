import XLSX from 'xlsx-js-style'


export default async function StyleReceitaVenda(
  sheet:XLSX.WorkSheet,celulas:string[][],indiceReceitaVenda: number[]){

  const letras: number[] = [0,1,2,3,4]

  // Cabeçalhos   
  
    letras.map(async (colun)=>{
      sheet[celulas[indiceReceitaVenda[0]-2][colun]].s = {
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
      sheet[celulas[indiceReceitaVenda[0]-1][colun]].s = {
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
        
      // Meio da tabela
      sheet[celulas[indiceReceitaVenda[0]][colun]].s = {
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
        sheet[celulas[indiceReceitaVenda[1]][colun]].s = {
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

      // Penultima linha da tabela
        sheet[celulas[indiceReceitaVenda[1]-1][colun]].s = {
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
            left: {
              style: 'medium',
            },
          },
        }
     

      ////
      // Bold Canto direito
      if(colun === 4){
        sheet[celulas[indiceReceitaVenda[0]][colun]].s = {
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

        sheet[celulas[indiceReceitaVenda[0]-2][colun]].s = {
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
            right: {
              style: 'thick',
            },
            top: {
              style: 'thick',
            },
            bottom: {
              style: 'thick',
            },
           
           
          },
        }
      }

    })

   

  return sheet
} 