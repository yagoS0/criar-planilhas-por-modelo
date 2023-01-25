import XLSX from 'xlsx-js-style'
import StyleCompras from './StyleFunctions/StyleCompras';
import StyleReceitaPresumido from './StyleFunctions/StyleReceitaPresumido';

import StyleReceitaServico from './StyleFunctions/StyleReceitaServico';
import StyleReceitaVenda from './StyleFunctions/StyleReceitaVenda';

export default async function styleSheet(sheet:XLSX.WorkSheet, 
  coluns: {
    dataValue: string;
    debitoValue: number;
    creditoValue: number;
    textValue: string;
    value: string;
}[] ){


    const celulas = [
      [ 'A1', 'B1', 'C1', 'D1', 'E1' ],
      [ 'A2', 'B2', 'C2', 'D2', 'E2' ],
      [ 'A3', 'B3', 'C3', 'D3', 'E3' ],
      [ 'A4', 'B4', 'C4', 'D4', 'E4' ],
      [ 'A5', 'B5', 'C5', 'D5', 'E5' ],
      [ 'A6', 'B6', 'C6', 'D6', 'E6' ],
      [ 'A7', 'B7', 'C7', 'D7', 'E7' ],
      [ 'A8', 'B8', 'C8', 'D8', 'E8' ],
      [ 'A9', 'B9', 'C9', 'D9', 'E9' ],
      [ 'A10', 'B10', 'C10', 'D10', 'E10' ],
      [ 'A11', 'B11', 'C11', 'D11', 'E11' ],
      [ 'A12', 'B12', 'C12', 'D12', 'E12' ],
      [ 'A13', 'B13', 'C13', 'D13', 'E13' ],
      [ 'A14', 'B14', 'C14', 'D14', 'E14' ],
      [ 'A15', 'B15', 'C15', 'D15', 'E15' ],
      [ 'A16', 'B16', 'C16', 'D16', 'E16' ],
      [ 'A17', 'B17', 'C17', 'D17', 'E17' ],
      [ 'A18', 'B18', 'C18', 'D18', 'E18' ],
      [ 'A19', 'B19', 'C19', 'D19', 'E19' ],
      [ 'A20', 'B20', 'C20', 'D20', 'E20' ],
      [ 'A21', 'B21', 'C21', 'D21', 'E21' ],
      [ 'A22', 'B22', 'C22', 'D22', 'E22' ],
      [ 'A23', 'B23', 'C23', 'D23', 'E23' ],
      [ 'A24', 'B24', 'C24', 'D24', 'E24' ],
      [ 'A25', 'B25', 'C25', 'D25', 'E25' ],
      [ 'A26', 'B26', 'C26', 'D26', 'E26' ],
      [ 'A27', 'B27', 'C27', 'D27', 'E27' ],
      [ 'A28', 'B28', 'C28', 'D28', 'E28' ],
      [ 'A29', 'B29', 'C29', 'D29', 'E29' ],
      [ 'A30', 'B30', 'C30', 'D30', 'E30' ],
      [ 'A31', 'B31', 'C31', 'D31', 'E31' ],
      [ 'A32', 'B32', 'C32', 'D32', 'E32' ],
      [ 'A33', 'B33', 'C33', 'D33', 'E33' ],
      [ 'A34', 'B34', 'C34', 'D34', 'E34' ],
      [ 'A35', 'B35', 'C35', 'D35', 'E35' ],
      [ 'A36', 'B36', 'C36', 'D36', 'E36' ],
      [ 'A37', 'B37', 'C37', 'D37', 'E37' ],
      [ 'A38', 'B38', 'C38', 'D38', 'E38' ],
      [ 'A39', 'B39', 'C39', 'D39', 'E39' ],
      [ 'A40', 'B40', 'C40', 'D40', 'E40' ],
      [ 'A41', 'B41', 'C41', 'D41', 'E41' ],
      [ 'A42', 'B42', 'C42', 'D42', 'E42' ],
      [ 'A43', 'B43', 'C43', 'D43', 'E43' ],
      [ 'A44', 'B44', 'C44', 'D44', 'E44' ],
      [ 'A45', 'B45', 'C45', 'D45', 'E45' ],
      [ 'A46', 'B46', 'C46', 'D46', 'E46' ],
      [ 'A47', 'B47', 'C47', 'D47', 'E47' ],
      [ 'A48', 'B48', 'C48', 'D48', 'E48' ],
      [ 'A49', 'B49', 'C49', 'D49', 'E49' ],
      [ 'A50', 'B50', 'C50', 'D50', 'E50' ],
      [ 'A51', 'B51', 'C51', 'D51', 'E51' ],
      [ 'A52', 'B52', 'C52', 'D52', 'E52' ],
      [ 'A53', 'B53', 'C53', 'D53', 'E53' ],
      [ 'A54', 'B54', 'C54', 'D54', 'E54' ],
      [ 'A55', 'B55', 'C55', 'D55', 'E55' ],
      [ 'A56', 'B56', 'C56', 'D56', 'E56' ],
      [ 'A57', 'B57', 'C57', 'D57', 'E57' ],
      [ 'A58', 'B58', 'C58', 'D58', 'E58' ],
      [ 'A59', 'B59', 'C59', 'D59', 'E59' ],
      [ 'A60', 'B60', 'C60', 'D60', 'E60' ],
      [ 'A61', 'B61', 'C61', 'D61', 'E61' ],
      [ 'A62', 'B62', 'C62', 'D62', 'E62' ],
      [ 'A63', 'B63', 'C63', 'D63', 'E63' ],
      [ 'A64', 'B64', 'C64', 'D64', 'E64' ],
      [ 'A65', 'B65', 'C65', 'D65', 'E65' ],
      [ 'A66', 'B66', 'C66', 'D66', 'E66' ],
      [ 'A67', 'B67', 'C67', 'D67', 'E67' ],
      [ 'A68', 'B68', 'C68', 'D68', 'E68' ],
      [ 'A69', 'B69', 'C69', 'D69', 'E69' ],
      [ 'A70', 'B70', 'C70', 'D70', 'E70' ],
      [ 'A71', 'B71', 'C71', 'D71', 'E71' ],
      [ 'A72', 'B72', 'C72', 'D72', 'E72' ],
      [ 'A73', 'B73', 'C73', 'D73', 'E73' ],
      [ 'A74', 'B74', 'C74', 'D74', 'E74' ],
      [ 'A75', 'B75', 'C75', 'D75', 'E75' ],
      [ 'A76', 'B76', 'C76', 'D76', 'E76' ],
      [ 'A77', 'B77', 'C77', 'D77', 'E77' ],
      [ 'A78', 'B78', 'C78', 'D78', 'E78' ],
      [ 'A79', 'B79', 'C79', 'D79', 'E79' ],
      [ 'A80', 'B80', 'C80', 'D80', 'E80' ],
      [ 'A81', 'B81', 'C81', 'D81', 'E81' ],
      [ 'A82', 'B82', 'C82', 'D82', 'E82' ],
      [ 'A83', 'B83', 'C83', 'D83', 'E83' ],
      [ 'A84', 'B84', 'C84', 'D84', 'E84' ],
      [ 'A85', 'B85', 'C85', 'D85', 'E85' ],
      [ 'A86', 'B86', 'C86', 'D86', 'E86' ],
      [ 'A87', 'B87', 'C87', 'D87', 'E87' ],
      [ 'A88', 'B88', 'C88', 'D88', 'E88' ],
      [ 'A89', 'B89', 'C89', 'D89', 'E89' ],
      [ 'A90', 'B90', 'C90', 'D90', 'E90' ],
      [ 'A91', 'B91', 'C91', 'D91', 'E91' ],
      [ 'A92', 'B92', 'C92', 'D92', 'E92' ],
      [ 'A93', 'B93', 'C93', 'D93', 'E93' ],
      [ 'A94', 'B94', 'C94', 'D94', 'E94' ],
      [ 'A95', 'B95', 'C95', 'D95', 'E95' ],
      [ 'A96', 'B96', 'C96', 'D96', 'E96' ],
      [ 'A97', 'B97', 'C97', 'D97', 'E97' ],
      [ 'A98', 'B98', 'C98', 'D98', 'E98' ],
      [ 'A99', 'B99', 'C99', 'D99', 'E99' ],
      [ 'A100', 'B100', 'C100', 'D100', 'E100' ], 
      [ 'A101', 'B101', 'C101', 'D101', 'E101' ]  

    ]

    // cabeçalho
    const letras: number[] = [0,1,2,3,4]
    letras.map((colun)=>{
      sheet[celulas[0][colun]].s = {
        font: {
          sz: 18,
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
    })
      

    var indiceReceitaServiço: number[] = []
    var indiceReceitaVenda: number[] = []
    var indiceReceitaPresumido: number[] = []
    var indiceCompra: number[] = []
 // Compras
 
    coluns.map(async (colun, index) => {
      if(colun.debitoValue === 524) {
        indiceCompra.push(index)
      }
      if(colun.debitoValue === 526) {
        indiceCompra.push(index)
      }

    })
  
    coluns.map(async (colun, index) => {
      // Receita Serviço
      if(colun.creditoValue === 372) {
        indiceReceitaServiço.push(index)
      }
      if(colun.creditoValue === 265) {
        indiceReceitaServiço.push(index)
      }
      // Receita Venda
      if(colun.creditoValue === 362) {
        indiceReceitaVenda.push(index)
      }
      if(colun.creditoValue === 265) {
        indiceReceitaVenda.push(index)
      }
      // Receita Presumido
      if(colun.creditoValue === 372) {
        indiceReceitaPresumido.push(index)
      }
      if(colun.creditoValue === 75) {
        indiceReceitaPresumido.push(index)
      }

    })

    // Receita Serviço
    if(indiceReceitaServiço.length === 2){
      await StyleReceitaServico(sheet,celulas,indiceReceitaServiço)
    }

    // Receita Venda
    if(indiceReceitaVenda.length === 2){
      await StyleReceitaVenda(sheet,celulas,indiceReceitaVenda)
    }
    //Receita Presumido
    if(indiceReceitaPresumido.length === 2){
      await StyleReceitaPresumido(sheet,celulas,indiceReceitaPresumido)
    }
    //Compras
    if(indiceCompra.length > 0){
      await StyleCompras(sheet,celulas,indiceCompra)
    }

    
    return sheet
}
 
