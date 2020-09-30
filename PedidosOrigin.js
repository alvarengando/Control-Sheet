/* ********************************************  Inicio Novo Pedido ******************************************* */

//Formular Pedido

function formularPedido(){

  var spreadsheet = SpreadsheetApp.getActive(); 
  
  spreadsheet.getRange('D5').setFormula('=IF(AND(C16="";G7="");"";IF(AP4="";COUNTA(\'Pedidos Dados\'!C2:C)+1;AP4))');
  spreadsheet.getRange('K15').setFormula('=IF(K14="";"";K13*K14)');
  spreadsheet.getRange('G7').setFormula('=IF(AND(C13="";C16="");"";AQ4)');
  spreadsheet.getRange('G9').setFormula('=IF(AND(C13="";C16="");"";AR4)');
  spreadsheet.getRange('H10').setFormula('=IF(AND(C13="";C16="");"";AS4)');
  spreadsheet.getRange('H11').setFormula('=IF(AND(C13="";C16="");"";AT4)');
  spreadsheet.getRange('H12').setFormula('=IF(AND(C13="";C16="");"";AU4)');
  spreadsheet.getRange('H13').setFormula('=IF(AND(C13="";C16="");"";AV4)');
  spreadsheet.getRange('G15').setFormula('=IF(AND(C13="";C16="");"";AW4)');          
  spreadsheet.getRange('M7').setFormula('=IF(D4="";"";"Pendente")');
  spreadsheet.getRange('J7').setFormula('=IF(C16="";"";C16)');

};

//Modo Salvar Pedido
function modoNovoPedido() {
    var spreadsheet = SpreadsheetApp.getActive(); 
    
    spreadsheet.getRange('AL3').setValue(1);    
    spreadsheet.getRange('D1').setValue("Novo");
    spreadsheet.getRange('AN3').setFormula('=IF(AND(C13="";C16="");"";IF(C16<>"";QUERY(\'Pedidos Dados\'!A:K;"SELECT * WHERE "&C16&" = I ORDER BY A DESC LIMIT 1");QUERY(\'Pedidos Dados\'!A:K;"SELECT * WHERE \'"&C13&"\' = D ORDER BY A DESC LIMIT 1")))');

    spreadsheet.getRangeList(['C13','C16','J10','K12:K15']).clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('D4').setBackground('#134f5c').setFontColor('#ffffff').clearDataValidations().setFormula('=IF(G7="";"";MAX(\'Pedidos Dados\'!A2:A)+1)');
    spreadsheet.getRange('D5').setFormula('=IF(AND(C16="";G7="");"";IF(AP4="";COUNTA(\'Pedidos Dados\'!C2:C)+1;AP4))');
    spreadsheet.getRange('D6').setFormula('=IF(K15="";"";NOW())');

    formularPedido();

    spreadsheet.getRange('C16').activate();
    
  };

  /* ************************ Salvar Pedido ******************** */

function SalvarPedido() {
  
  var spreadsheet = SpreadsheetApp.getActive();
  var PedidosDados = spreadsheet.getSheetByName('Pedidos Dados');
  
  if (spreadsheet.getRange('AK3').getValue() > 0 ) 
  {
    Browser.msgBox("Erro", "Necessário preencher todos os campos essenciais!", Browser.Buttons.OK);
  }
  
  else{
    
    // Sal  var na Página Pedidos Dados
                                           
        var values = [[spreadsheet.getRange('D4').getValue(),    // ID Pedido
                       spreadsheet.getRange('D6').getValue(),    // Data Pedido
                       spreadsheet.getRange('D5').getValue(),    // ID Cliente
                       spreadsheet.getRange('G7').getValue(),    // Cliente
                       spreadsheet.getRange('G9').getValue(),    // Logradouro
                       spreadsheet.getRange('H10').getValue(),   // Complemento
                       spreadsheet.getRange('H11').getValue(),   // Município
                       spreadsheet.getRange('H12').getValue(),   // Bairro
                       spreadsheet.getRange('H13').getValue(),   // Telefone Cadastrado
                       spreadsheet.getRange('G15').getValue(),   // Referência
                       spreadsheet.getRange('J7').getValue(),    // Telefone Utilizado
                       spreadsheet.getRange('J10').getValue(),   // Motorista
                       spreadsheet.getRange('K12').getValue(),   // Produto
                       spreadsheet.getRange('K13').getValue(),   // Quantidade
                       spreadsheet.getRange('K14').getValue(),   // Preço
                       spreadsheet.getRange('K15').getValue(),   // Total
                       spreadsheet.getRange('M7').getValue(),    // Status
                     ]];
       
          PedidosDados.getRange(PedidosDados.getLastRow()+1,1,1,17).setValues(values);
          spreadsheet.getRangeList(['C13','C16','J10','K12:K15']).clear({contentsOnly: true, skipFilteredRows: true});

          formularPedido();
          Browser.msgBox("Informativo", "Registro salvo com sucesso!", Browser.Buttons.OK);
          spreadsheet.getRange('C16').activate(); 
                   
       } 
         
};

//Formular editar Pedido
function formularEditarPedido(){

  var spreadsheet = SpreadsheetApp.getActive();

  spreadsheet.getRange('D5').setFormula('=IF(AND(C13="";C16="");"";IF(AP4="";COUNTA(\'Pedidos Dados\'!C2:C)+1;AP4))'); 
  spreadsheet.getRange('D6').setFormula('=IF(D4="";"";AO4)');
  spreadsheet.getRange('G7').setFormula('=IF(AND(C13="";C16="");"";AQ4)');
  spreadsheet.getRange('G9').setFormula('=IF(AND(C13="";C16="");"";AR4)');
  spreadsheet.getRange('H10').setFormula('=IF(AND(C13="";C16="");"";AS4)');
  spreadsheet.getRange('H11').setFormula('=IF(AND(C13="";C16="");"";AT4)');
  spreadsheet.getRange('H12').setFormula('=IF(AND(C13="";C16="");"";AU4)');
  spreadsheet.getRange('H13').setFormula('=IF(AND(C13="";C16="");"";AV4)');
  spreadsheet.getRange('G15').setFormula('=IF(AND(C13="";C16="");"";AW4)');
  spreadsheet.getRange('J7').setFormula('=IF(D4="";"";AX4)');
  spreadsheet.getRange('J10').setFormula('=IF(D4="";"";AY4)');
  spreadsheet.getRange('K12').setFormula('=IF(D4="";"";AZ4)');
  spreadsheet.getRange('K13').setFormula('=IF(D4="";"";BA4)');
  spreadsheet.getRange('K14').setFormula('=IF(D4="";"";BB4)');
  spreadsheet.getRange('K15').setFormula('=IF(K14="";"";K13*K14)');
  spreadsheet.getRange('M7').setFormula('=IF(D4="";"";BD4)');

};

//Modo Editar Pedido
function modoEditarPedido(){

 var spreadsheet = SpreadsheetApp.getActive();

 spreadsheet.getRange('AL3').setValue(2);
 spreadsheet.getRange('AN3').setFormula('=IF(AM3="";""; IF(AM3="16";QUERY(\'Pedidos Dados\'!A:Q;"SELECT * WHERE "&C16&" = I");IF(AM3="164";QUERY(\'Pedidos Dados\'!A:Q;"SELECT * WHERE "&C16&" = I AND "&D4&" = A");IF(AM3="13";QUERY(\'Pedidos Dados\'!A:Q;"SELECT * WHERE \'"&C13&"\' = D");IF(AM3="134";QUERY(\'Pedidos Dados\'!A:Q;"SELECT * WHERE \'"&C13&"\' = D AND "&D4&" = A"))))))');
 
 spreadsheet.getRangeList(['D4','C13','C16']).clear({contentsOnly: true, skipFilteredRows: true});
 
 spreadsheet.getRange('D1').setValue("Editar");
 //ID Pedido
 spreadsheet.getRange('D4').setBackground('#ffffff').setFontColor('#000000').setDataValidation(SpreadsheetApp.newDataValidation() .setAllowInvalid(false).requireValueInRange(spreadsheet.getRange('\'Pedidos\'!$BG$4:$BG'), true).build()); 
 
 formularEditarPedido();

 spreadsheet.getRange('C16').activate();


};

//Salvar alteração
function editarPedido(){

  var spreadsheet = SpreadsheetApp.getActive();
  var PedidosDados = spreadsheet.getSheetByName('Pedidos Dados');
  var linhaPedido = spreadsheet.getRange('AJ3').getValue(); //linha correspondente em Pedidos dados

   
  if (spreadsheet.getRange('AK3').getValue() > 0 ) 
  {
    Browser.msgBox("Erro", "Necessário preencher todos os campos essenciais!", Browser.Buttons.OK);
  }
  
  else{
    
    // Salvar na Página Pedidos Dados
                                           
    var values = [[spreadsheet.getRange('D6').getValue(),    // Data Pedido
                   spreadsheet.getRange('D5').getValue(),    // ID Cliente
                   spreadsheet.getRange('G7').getValue(),    // Cliente
                   spreadsheet.getRange('G9').getValue(),    // Logradouro
                   spreadsheet.getRange('H10').getValue(),    // Complemento
                   spreadsheet.getRange('H11').getValue(),   // Município
                   spreadsheet.getRange('H12').getValue(),   // Bairro
                   spreadsheet.getRange('H13').getValue(),   // Telefone Cadastrado
                   spreadsheet.getRange('G15').getValue(),   // Referência
                   spreadsheet.getRange('J7').getValue(),    // Telefone Utilizado
                   spreadsheet.getRange('J10').getValue(),   // Motorista
                   spreadsheet.getRange('K12').getValue(),   // Produto
                   spreadsheet.getRange('K13').getValue(),   // Quantidade
                   spreadsheet.getRange('K14').getValue(),   // Preço
                   spreadsheet.getRange('K15').getValue(),   // Total
                   spreadsheet.getRange('M7').getValue(),    // Status
                 ]];
       
          PedidosDados.getRange(linhaPedido, 2, 1, 16).setValues(values);

          Browser.msgBox("Informativo", "Registro alterado com sucesso!", Browser.Buttons.OK);
          spreadsheet.getRangeList(['D4','C13','C16']).clear({contentsOnly: true, skipFilteredRows: true});
          formularEditarPedido();
          spreadsheet.getRange('C16').activate();

    }
};

/**    * ********************************* */
//Modo Excluir Pedido
function modoDeletarPedido(){

  var spreadsheet = SpreadsheetApp.getActive();

  spreadsheet.getRange('AL3').setValue(3);
  spreadsheet.getRange('AN3').setFormula('=IF(AM3="";""; IF(AM3="16";QUERY(\'Pedidos Dados\'!A:Q;"SELECT * WHERE "&C16&" = I");IF(AM3="164";QUERY(\'Pedidos Dados\'!A:Q;"SELECT * WHERE "&C16&" = I AND "&D4&" = A");IF(AM3="13";QUERY(\'Pedidos Dados\'!A:Q;"SELECT * WHERE \'"&C13&"\' = D");IF(AM3="134";QUERY(\'Pedidos Dados\'!A:Q;"SELECT * WHERE \'"&C13&"\' = D AND "&D4&" = A"))))))');
  
  spreadsheet.getRangeList(['D4','C13','C16']).clear({contentsOnly: true, skipFilteredRows: true});
  
  spreadsheet.getRange('D1').setValue("Deletar");
  //ID Pedido
  spreadsheet.getRange('D4').setBackground('#ffffff').setFontColor('#000000').setDataValidation(SpreadsheetApp.newDataValidation() .setAllowInvalid(false).requireValueInRange(spreadsheet.getRange('\'Pedidos\'!$BG$4:$BG'), true).build()); 
  
  formularEditarPedido();
 
  spreadsheet.getRange('C16').activate();
  
};

function deletarPedido(){

  var spreadsheet = SpreadsheetApp.getActive();
  var PedidosDados = spreadsheet.getSheetByName('Pedidos Dados');
  var linhaPedido = spreadsheet.getRange('AJ3').getValue(); //linha correspondente em Pedidos dados

   
  if (spreadsheet.getRange('AK3').getValue() > 0 ) 
  {
    Browser.msgBox("Erro", "Necessário preencher todos os campos essenciais!", Browser.Buttons.OK);
  }
  
  else{
           
        PedidosDados.deleteRow(linhaPedido);
        spreadsheet.getRangeList(['D4','C13','C16']).clear({contentsOnly: true, skipFilteredRows: true});
        Browser.msgBox("Informativo", "Registro Deletado com sucesso!", Browser.Buttons.OK);
        
       // formularEditarPedido();
        spreadsheet.getRange('C16').activate();

    }

};


//******************    Finalizador   ******************************************************************


function FinalizadorPedido(){

  var spreadsheet = SpreadsheetApp.getActive();

  if(spreadsheet.getRange('AL3').getValue() == 1)
  {
      SalvarPedido();
  }
                      
  else if(spreadsheet.getRange('AL3').getValue() == 2)
   {
         editarPedido();
   }
                      
   else
   {
    deletarPedido(); 
   } 


};