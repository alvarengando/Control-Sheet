/* ********************************************  Inicio Novo Pedido ******************************************* */

//Formular Pedido

function formularPedido2() {

  var spreadsheet = SpreadsheetApp.getActive();
  var pedidos2 = spreadsheet.getSheetByName('Pedidos 2');

  pedidos2.getRange('D5').setFormula('=IF(AND(C16="";G7="");"";IF(AP4="";COUNTA(\'Pedidos Dados\'!C2:C)+1;AP4))');
  pedidos2.getRange('D6').setFormula('=IF(K15="";"";NOW())');
  pedidos2.getRange('K15').setFormula('=IF(K14="";"";K13*K14)');
  pedidos2.getRange('G7').setFormula('=IF(AND(C13="";C16="");"";AQ4)');
  pedidos2.getRange('G9').setFormula('=IF(AND(C13="";C16="");"";AR4)');
  pedidos2.getRange('H10').setFormula('=IF(AND(C13="";C16="");"";AS4)');
  pedidos2.getRange('H11').setFormula('=IF(AND(C13="";C16="");"";AT4)');
  pedidos2.getRange('H12').setFormula('=IF(AND(C13="";C16="");"";AU4)');
  pedidos2.getRange('H13').setFormula('=IF(AND(C13="";C16="");"";AV4)');
  pedidos2.getRange('G15').setFormula('=IF(AND(C13="";C16="");"";AW4)');
  pedidos2.getRange('M7').setFormula('=IF(D4="";"";"Pendente")');
  pedidos2.getRange('J7').setFormula('=IF(C16="";H13;C16)');
  pedidos2.getRange('J9').setFormula('=IF(J7="";"";"Telegás")');

};

//Modo Salvar Pedido
function modoNovoPedido2() {
  var spreadsheet = SpreadsheetApp.getActive();
  var pedidos2 = spreadsheet.getSheetByName('Pedidos 2');

  pedidos2.getRange('AL3').setValue(1);
  pedidos2.getRange('D1').setValue("Novo");
  pedidos2.getRange('AN3').setFormula('=IF(AND(C13="";C16="");"";IF(C16<>"";QUERY(\'Pedidos Dados\'!A:K;"SELECT * WHERE "&C16&" = I ORDER BY A DESC LIMIT 1");QUERY(\'Pedidos Dados\'!A:K;"SELECT * WHERE \'"&C13&"\' = D ORDER BY A DESC LIMIT 1")))');

  pedidos2.getRangeList(['C13', 'C16', 'J11', 'K12:K16', 'M10']).clear({ contentsOnly: true, skipFilteredRows: true });
  pedidos2.getRange('D4').setBackground('#b45f06').setFontColor('#ffffff').clearDataValidations().setFormula('=IF(G7="";"";MAX(\'Pedidos Dados\'!A2:A)+1)');
  pedidos2.getRange('D5').setFormula('=IF(AND(C16="";G7="");"";IF(AP4="";COUNTA(\'Pedidos Dados\'!C2:C)+1;AP4))');

  formularPedido2();

  pedidos2.getRange('C16').activate();

};

/* ************************ Salvar Pedido ******************** */

function SalvarPedido2() {

  var spreadsheet = SpreadsheetApp.getActive();
  var pedidos2 = spreadsheet.getSheetByName('Pedidos 2');
  var PedidosDados = spreadsheet.getSheetByName('Pedidos Dados');

  if (pedidos2.getRange('AK3').getValue() > 0) {
    Browser.msgBox("Erro", "Necessário preencher todos os campos essenciais!", Browser.Buttons.OK);
  }

  else {

    // Sal  var na Página Pedidos Dados

    var values = [[pedidos2.getRange('D4').getValue(),    // ID Pedido
    pedidos2.getRange('D6').getValue(),    // Data Pedido
    pedidos2.getRange('D5').getValue(),    // ID Cliente
    pedidos2.getRange('G7').getValue(),    // Cliente
    pedidos2.getRange('G9').getValue(),    // Logradouro
    pedidos2.getRange('H10').getValue(),   // Complemento
    pedidos2.getRange('H11').getValue(),   // Município
    pedidos2.getRange('H12').getValue(),   // Bairro
    pedidos2.getRange('H13').getValue(),   // Telefone Cadastrado
    pedidos2.getRange('G15').getValue(),   // Referência
    pedidos2.getRange('J7').getValue(),    // Telefone Utilizado
    pedidos2.getRange('J11').getValue(),   // Motorista
    pedidos2.getRange('K12').getValue(),   // Produto
    pedidos2.getRange('K13').getValue(),   // Quantidade
    pedidos2.getRange('K14').getValue(),   // Preço
    pedidos2.getRange('K15').getValue(),   // Total
    pedidos2.getRange('M7').getValue(),    // Status
    pedidos2.getRange('J9').getValue(),    // Canal de Venda
      "",                                    // Justificativa de cancelamento  
    pedidos2.getRange('K16').getValue()     // Forma de Pagamento
    ]];

    PedidosDados.getRange(PedidosDados.getLastRow() + 1, 1, 1, 20).setValues(values);
    pedidos2.getRangeList(['C13', 'C16', 'J11', 'K12:K16']).clear({ contentsOnly: true, skipFilteredRows: true });

    formularPedido2();
    Browser.msgBox("Informativo", "Registro salvo com sucesso!", Browser.Buttons.OK);
    pedidos2.getRange('C16').activate();

  }

};

//Formular editar Pedido
function formularEditarPedido2() {

  var spreadsheet = SpreadsheetApp.getActive();
  var pedidos2 = spreadsheet.getSheetByName('Pedidos 2');

  pedidos2.getRange('D5').setFormula('=IF(AND(C13="";C16="");"";IF(AP4="";COUNTA(\'Pedidos Dados\'!C2:C)+1;AP4))');
  pedidos2.getRange('D6').setFormula('=IF(D4="";"";AO4)');
  pedidos2.getRange('G7').setFormula('=IF(AND(C13="";C16="");"";AQ4)');
  pedidos2.getRange('G9').setFormula('=IF(AND(C13="";C16="");"";AR4)');
  pedidos2.getRange('H10').setFormula('=IF(AND(C13="";C16="");"";AS4)');
  pedidos2.getRange('H11').setFormula('=IF(AND(C13="";C16="");"";AT4)');
  pedidos2.getRange('H12').setFormula('=IF(AND(C13="";C16="");"";AU4)');
  pedidos2.getRange('H13').setFormula('=IF(AND(C13="";C16="");"";AV4)');
  pedidos2.getRange('G15').setFormula('=IF(AND(C13="";C16="");"";AW4)');
  pedidos2.getRange('J7').setFormula('=IF(D4="";"";AX4)');
  pedidos2.getRange('J11').setFormula('=IF(D4="";"";AY4)');
  pedidos2.getRange('J9').setFormula('=IF(D4="";"";BE4)');
  pedidos2.getRange('K12').setFormula('=IF(D4="";"";AZ4)');
  pedidos2.getRange('K13').setFormula('=IF(D4="";"";BA4)');
  pedidos2.getRange('K14').setFormula('=IF(D4="";"";BB4)');
  pedidos2.getRange('K15').setFormula('=IF(K14="";"";K13*K14)');
  pedidos2.getRange('K16').setFormula('=IF(D4="";"";BG4)');
  pedidos2.getRange('M7').setFormula('=IF(D4="";"";BD4)');

};

//Modo Editar Pedido
function modoEditarPedido2() {

  var spreadsheet = SpreadsheetApp.getActive();
  var pedidos2 = spreadsheet.getSheetByName('Pedidos 2');

  pedidos2.getRange('AL3').setValue(2);
  pedidos2.getRange('AN3').setFormula('=IF(AM3="";""; IF(AM3="16";QUERY(\'Pedidos Dados\'!A:T;"SELECT * WHERE "&C16&" = I");IF(AM3="164";QUERY(\'Pedidos Dados\'!A:T;"SELECT * WHERE "&C16&" = I AND "&D4&" = A");IF(AM3="13";QUERY(\'Pedidos Dados\'!A:T;"SELECT * WHERE \'"&C13&"\' = D");IF(AM3="134";QUERY(\'Pedidos Dados\'!A:T;"SELECT * WHERE \'"&C13&"\' = D AND "&D4&" = A"))))))');

  pedidos2.getRangeList(['D4', 'C13', 'C16']).clear({ contentsOnly: true, skipFilteredRows: true });

  pedidos2.getRange('D1').setValue("Editar");
  //ID Pedido
  pedidos2.getRange('D4').setBackground('#ffffff').setFontColor('#000000').setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInRange(spreadsheet.getRange('\'Pedidos 2\'!$BH$4:$BH'), true).build());

  formularEditarPedido2();

  pedidos2.getRange('C16').activate();


};

//Salvar alteração
function editarPedido2() {

  var spreadsheet = SpreadsheetApp.getActive();
  var pedidos2 = spreadsheet.getSheetByName('Pedidos 2');
  var PedidosDados = spreadsheet.getSheetByName('Pedidos Dados');
  var linhaPedido = pedidos2.getRange('AJ3').getValue(); //linha correspondente em Pedidos dados
  var contVazioPedi = pedidos2.getRange('AK3').getValue();
  var status = pedidos2.getRange('AI3').getValue();

  if (contVazioPedi > 0 || status == 1) {
    if (contVazioPedi > 0) {
      Browser.msgBox("Erro", "Necessário preencher todos os campos essenciais!", Browser.Buttons.OK);
    } else {
      Browser.msgBox("Erro", "Necessário preencher a justificativa do Cancelamento!", Browser.Buttons.OK);
    }


  } else {

    // Salvar na Página Pedidos Dados

    var values = [[pedidos2.getRange('D6').getValue(),    // Data Pedido
    pedidos2.getRange('D5').getValue(),    // ID Cliente
    pedidos2.getRange('G7').getValue(),    // Cliente
    pedidos2.getRange('G9').getValue(),    // Logradouro
    pedidos2.getRange('H10').getValue(),    // Complemento
    pedidos2.getRange('H11').getValue(),   // Município
    pedidos2.getRange('H12').getValue(),   // Bairro
    pedidos2.getRange('H13').getValue(),   // Telefone Cadastrado
    pedidos2.getRange('G15').getValue(),   // Referência
    pedidos2.getRange('J7').getValue(),    // Telefone Utilizado
    pedidos2.getRange('J11').getValue(),   // Motorista
    pedidos2.getRange('K12').getValue(),   // Produto
    pedidos2.getRange('K13').getValue(),   // Quantidade
    pedidos2.getRange('K14').getValue(),   // Preço
    pedidos2.getRange('K15').getValue(),   // Total
    pedidos2.getRange('M7').getValue(),    // Status
    pedidos2.getRange('J9').getValue(),    // Canal de Venda
    pedidos2.getRange('M10').getValue(),   // Justificativa
    pedidos2.getRange('K16').getValue()    // Forma de Pagamento 
    ]];

    PedidosDados.getRange(linhaPedido, 2, 1, 19).setValues(values);

    Browser.msgBox("Informativo", "Registro alterado com sucesso!", Browser.Buttons.OK);
    pedidos2.getRangeList(['D4', 'C13', 'C16', 'M10']).clear({ contentsOnly: true, skipFilteredRows: true });
    formularEditarPedido2();
    pedidos2.getRange('C16').activate();

  }
};

/**    * ********************************* */
//Modo Excluir Pedido
function modoDeletarPedido2() {

  var spreadsheet = SpreadsheetApp.getActive();
  var pedidos2 = spreadsheet.getSheetByName('Pedidos 2');

  pedidos2.getRange('AL3').setValue(3);
  pedidos2.getRange('AN3').setFormula('=IF(AM3="";""; IF(AM3="16";QUERY(\'Pedidos Dados\'!A:T;"SELECT * WHERE "&C16&" = I");IF(AM3="164";QUERY(\'Pedidos Dados\'!A:T;"SELECT * WHERE "&C16&" = I AND "&D4&" = A");IF(AM3="13";QUERY(\'Pedidos Dados\'!A:T;"SELECT * WHERE \'"&C13&"\' = D");IF(AM3="134";QUERY(\'Pedidos Dados\'!A:T;"SELECT * WHERE \'"&C13&"\' = D AND "&D4&" = A"))))))');

  pedidos2.getRangeList(['D4', 'C13', 'C16']).clear({ contentsOnly: true, skipFilteredRows: true });

  pedidos2.getRange('D1').setValue("Deletar");
  //ID Pedido
  pedidos2.getRange('D4').setBackground('#ffffff').setFontColor('#000000').setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInRange(spreadsheet.getRange('\'Pedidos 2\'!$BH$4:$BH'), true).build());

  formularEditarPedido2();

  pedidos2.getRange('C16').activate();

};

function deletarPedido2() {

  var spreadsheet = SpreadsheetApp.getActive();
  var pedidos2 = spreadsheet.getSheetByName('Pedidos 2');
  var PedidosDados = spreadsheet.getSheetByName('Pedidos Dados');
  var linhaPedido = pedidos2.getRange('AJ3').getValue(); //linha correspondente em Pedidos dados


  if (pedidos2.getRange('AK3').getValue() > 0) {
    Browser.msgBox("Erro", "Necessário preencher todos os campos essenciais!", Browser.Buttons.OK);
  }

  else {

    PedidosDados.deleteRow(linhaPedido);
    pedidos2.getRangeList(['D4', 'C13', 'C16', 'J11', 'M10']).clear({ contentsOnly: true, skipFilteredRows: true });
    Browser.msgBox("Informativo", "Registro Deletado com sucesso!", Browser.Buttons.OK);

    pedidos2.getRange('C16').activate();

  }

};


//******************    Finalizador   ******************************************************************


function finalizadorPedido2() {

  var spreadsheet = SpreadsheetApp.getActive();
  var pedidos2 = spreadsheet.getSheetByName('Pedidos 2');

  if (pedidos2.getRange('AL3').getValue() == 1) {
    SalvarPedido2();
  }

  else if (pedidos2.getRange('AL3').getValue() == 2) {
    editarPedido2();
  }

  else {
    deletarPedido2();
  }


};