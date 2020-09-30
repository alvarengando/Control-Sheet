/* ********************************************  Inicio Novo Pedido ******************************************* */

//Formular Pedido

function formularPedido1() {

  var spreadsheet = SpreadsheetApp.getActive();
  var vendas = spreadsheet.getSheetByName('Vendas');

  vendas.getRange('D5').setFormula('=IF(AND(C16="";G7="");"";IF(AP4="";COUNTA(\'Vendas Dados\'!C2:C)+1;AP4))');
  vendas.getRange('D6').setFormula('=IF(K15="";"";NOW())');
  vendas.getRange('K15').setFormula('=IF(K14="";"";K13*K14)');
  vendas.getRange('G7').setFormula('=IF(AND(C13="";C16="");"";AQ4)');
  vendas.getRange('G9').setFormula('=IF(AND(C13="";C16="");"";AR4)');
  vendas.getRange('H10').setFormula('=IF(AND(C13="";C16="");"";AS4)');
  vendas.getRange('H11').setFormula('=IF(AND(C13="";C16="");"";AT4)');
  vendas.getRange('H12').setFormula('=IF(AND(C13="";C16="");"";AU4)');
  vendas.getRange('H13').setFormula('=IF(AND(C13="";C16="");"";AV4)');
  vendas.getRange('G15').setFormula('=IF(AND(C13="";C16="");"";AW4)');
  vendas.getRange('M7').setFormula('=IF(D4="";"";"Pendente")');
  vendas.getRange('J7').setFormula('=IF(C16="";H13;C16)');
  vendas.getRange('J9').setFormula('=IF(J7="";"";"Telegás")');

};

//Modo Salvar Pedido
function modoNovoPedido1() {
  var spreadsheet = SpreadsheetApp.getActive();
  var vendas = spreadsheet.getSheetByName('Vendas');

  vendas.getRange('AL3').setValue(1);
  vendas.getRange('D1').setValue("Novo");
  vendas.getRange('AN3').setFormula('=IF(AND(C13="";C16="");"";IF(C16<>"";QUERY(\'Vendas Dados\'!A:K;"SELECT * WHERE "&C16&" = I ORDER BY A DESC LIMIT 1");QUERY(\'Vendas Dados\'!A:K;"SELECT * WHERE \'"&C13&"\' = D ORDER BY A DESC LIMIT 1")))');

  vendas.getRangeList(['C13', 'C16', 'J11', 'K12:K16', 'M10']).clear({ contentsOnly: true, skipFilteredRows: true });
  vendas.getRange('D4').setBackground('#134f5c').setFontColor('#ffffff').clearDataValidations().setFormula('=IF(G7="";"";MAX(\'Vendas Dados\'!A2:A)+1)');
  vendas.getRange('D5').setFormula('=IF(AND(C16="";G7="");"";IF(AP4="";COUNTA(\'Vendas Dados\'!C2:C)+1;AP4))');


  formularPedido1();

  vendas.getRange('C16').activate();

};

/* ************************ Salvar Pedido ******************** */

function SalvarPedido1() {

  var spreadsheet = SpreadsheetApp.getActive();
  var vendas = spreadsheet.getSheetByName('Vendas');
  var vendasDados = spreadsheet.getSheetByName('Vendas Dados');

  if (vendas.getRange('AK3').getValue() > 0) {
    Browser.msgBox("Erro", "Necessário preencher todos os campos essenciais!", Browser.Buttons.OK);
  }

  else {

    // Sal  var na Página Vendas Dados

    var values = [[vendas.getRange('D4').getValue(),    // ID Pedido
    vendas.getRange('D6').getValue(),    // Data Pedido
    vendas.getRange('D5').getValue(),    // ID Cliente
    vendas.getRange('G7').getValue(),    // Cliente
    vendas.getRange('G9').getValue(),    // Logradouro
    vendas.getRange('H10').getValue(),   // Complemento
    vendas.getRange('H11').getValue(),   // Município
    vendas.getRange('H12').getValue(),   // Bairro
    vendas.getRange('H13').getValue(),   // Telefone Cadastrado
    vendas.getRange('G15').getValue(),   // Referência
    vendas.getRange('J7').getValue(),    // Telefone Utilizado
    vendas.getRange('J11').getValue(),   // Motorista
    vendas.getRange('K12').getValue(),   // Produto
    vendas.getRange('K13').getValue(),   // Quantidade
    vendas.getRange('K14').getValue(),   // Preço
    vendas.getRange('K15').getValue(),   // Total
    vendas.getRange('M7').getValue(),    // Status
    vendas.getRange('J9').getValue(),    // Canal de Venda
      "",                                  // Justificativa de cancelamento  
    vendas.getRange('K16').getValue()    // Forma de Pagamento
    ]];

    vendasDados.getRange(vendasDados.getLastRow() + 1, 1, 1, 20).setValues(values);
    vendas.getRangeList(['C13', 'C16', 'J11', 'K12:K16']).clear({ contentsOnly: true, skipFilteredRows: true });

    formularPedido();
    Browser.msgBox("Informativo", "Registro salvo com sucesso!", Browser.Buttons.OK);
    vendas.getRange('C16').activate();

  }

};

//Formular editar Pedido
function formularEditarPedido1() {

  var spreadsheet = SpreadsheetApp.getActive();
  var vendas = spreadsheet.getSheetByName('Vendas');

  vendas.getRange('D5').setFormula('=IF(AND(C13="";C16="");"";IF(AP4="";COUNTA(\'Vendas Dados\'!C2:C)+1;AP4))');
  vendas.getRange('D6').setFormula('=IF(D4="";"";AO4)');
  vendas.getRange('G7').setFormula('=IF(AND(C13="";C16="");"";AQ4)');
  vendas.getRange('G9').setFormula('=IF(AND(C13="";C16="");"";AR4)');
  vendas.getRange('H10').setFormula('=IF(AND(C13="";C16="");"";AS4)');
  vendas.getRange('H11').setFormula('=IF(AND(C13="";C16="");"";AT4)');
  vendas.getRange('H12').setFormula('=IF(AND(C13="";C16="");"";AU4)');
  vendas.getRange('H13').setFormula('=IF(AND(C13="";C16="");"";AV4)');
  vendas.getRange('G15').setFormula('=IF(AND(C13="";C16="");"";AW4)');
  vendas.getRange('J7').setFormula('=IF(D4="";"";AX4)');
  vendas.getRange('J11').setFormula('=IF(D4="";"";AY4)');
  vendas.getRange('J9').setFormula('=IF(D4="";"";BE4)');
  vendas.getRange('K12').setFormula('=IF(D4="";"";AZ4)');
  vendas.getRange('K13').setFormula('=IF(D4="";"";BA4)');
  vendas.getRange('K14').setFormula('=IF(D4="";"";BB4)');
  vendas.getRange('K15').setFormula('=IF(K14="";"";K13*K14)');
  vendas.getRange('K16').setFormula('=IF(D4="";"";BG4)');
  vendas.getRange('M7').setFormula('=IF(D4="";"";BD4)');

};

//Modo Editar Pedido
function modoEditarPedido1() {

  var spreadsheet = SpreadsheetApp.getActive();
  var vendas = spreadsheet.getSheetByName('Vendas');

  vendas.getRange('AL3').setValue(2);
  vendas.getRange('AN3').setFormula('=IF(AM3="";""; IF(AM3="16";QUERY(\'Vendas Dados\'!A:T;"SELECT * WHERE "&C16&" = I");IF(AM3="164";QUERY(\'Vendas Dados\'!A:T;"SELECT * WHERE "&C16&" = I AND "&D4&" = A");IF(AM3="13";QUERY(\'Vendas Dados\'!A:T;"SELECT * WHERE \'"&C13&"\' = D");IF(AM3="134";QUERY(\'Vendas Dados\'!A:T;"SELECT * WHERE \'"&C13&"\' = D AND "&D4&" = A"))))))');

  vendas.getRangeList(['D4', 'C13', 'C16']).clear({ contentsOnly: true, skipFilteredRows: true });

  vendas.getRange('D1').setValue("Editar");
  //ID Pedido
  vendas.getRange('D4').setBackground('#ffffff').setFontColor('#000000').setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInRange(spreadsheet.getRange('\'Vendas\'!$BH$4:$BH'), true).build());

  formularEditarPedido1();

  vendas.getRange('C16').activate();


};

//Salvar alteração
function editarPedido1() {

  var spreadsheet = SpreadsheetApp.getActive();
  var vendas = spreadsheet.getSheetByName('Vendas');
  var vendasDados = spreadsheet.getSheetByName('Vendas Dados');
  var linhaPedido = vendas.getRange('AJ3').getValue(); //linha correspondente em Vendas Dados
  var contVazioPedi = vendas.getRange('AK3').getValue();
  var status = vendas.getRange('AI3').getValue();

  if (contVazioPedi > 0 || status == 1) {
    if (contVazioPedi > 0) {
      Browser.msgBox("Erro", "Necessário preencher todos os campos essenciais!", Browser.Buttons.OK);
    } else {
      Browser.msgBox("Erro", "Necessário preencher a justificativa do Cancelamento!", Browser.Buttons.OK);
    }


  } else {

    // Salvar na Página Vendas Dados

    var values = [[vendas.getRange('D6').getValue(),    // Data Pedido
    vendas.getRange('D5').getValue(),    // ID Cliente
    vendas.getRange('G7').getValue(),    // Cliente
    vendas.getRange('G9').getValue(),    // Logradouro
    vendas.getRange('H10').getValue(),    // Complemento
    vendas.getRange('H11').getValue(),   // Município
    vendas.getRange('H12').getValue(),   // Bairro
    vendas.getRange('H13').getValue(),   // Telefone Cadastrado
    vendas.getRange('G15').getValue(),   // Referência
    vendas.getRange('J7').getValue(),    // Telefone Utilizado
    vendas.getRange('J11').getValue(),   // Motorista
    vendas.getRange('K12').getValue(),   // Produto
    vendas.getRange('K13').getValue(),   // Quantidade
    vendas.getRange('K14').getValue(),   // Preço
    vendas.getRange('K15').getValue(),   // Total
    vendas.getRange('M7').getValue(),    // Status
    vendas.getRange('J9').getValue(),   // Canal de Venda
    vendas.getRange('M10').getValue(),   // Justificativa
    vendas.getRange('K16').getValue()    // Forma de Pagamento
    ]];

    vendasDados.getRange(linhaPedido, 2, 1, 19).setValues(values);

    Browser.msgBox("Informativo", "Registro alterado com sucesso!", Browser.Buttons.OK);
    vendas.getRangeList(['D4', 'C13', 'C16', 'M10']).clear({ contentsOnly: true, skipFilteredRows: true });
    formularEditarPedido1();
    vendas.getRange('C16').activate();

  }
};

/**    * ********************************* */
//Modo Excluir Pedido
function modoDeletarPedido1() {

  var spreadsheet = SpreadsheetApp.getActive();
  var vendas = spreadsheet.getSheetByName('Vendas');

  vendas.getRange('AL3').setValue(3);
  vendas.getRange('AN3').setFormula('=IF(AM3="";""; IF(AM3="16";QUERY(\'Vendas Dados\'!A:T;"SELECT * WHERE "&C16&" = I");IF(AM3="164";QUERY(\'Vendas Dados\'!A:T;"SELECT * WHERE "&C16&" = I AND "&D4&" = A");IF(AM3="13";QUERY(\'Vendas Dados\'!A:T;"SELECT * WHERE \'"&C13&"\' = D");IF(AM3="134";QUERY(\'Vendas Dados\'!A:T;"SELECT * WHERE \'"&C13&"\' = D AND "&D4&" = A"))))))');

  vendas.getRangeList(['D4', 'C13', 'C16']).clear({ contentsOnly: true, skipFilteredRows: true });

  vendas.getRange('D1').setValue("Deletar");
  //ID Pedido
  vendas.getRange('D4').setBackground('#ffffff').setFontColor('#000000').setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInRange(spreadsheet.getRange('\'Vendas\'!$BH$4:$BH'), true).build());

  formularEditarPedido1();

  vendas.getRange('C16').activate();

};

function deletarPedido1() {

  var spreadsheet = SpreadsheetApp.getActive();
  var vendas = spreadsheet.getSheetByName('Vendas');
  var vendasDados = spreadsheet.getSheetByName('Vendas Dados');
  var linhaPedido = vendas.getRange('AJ3').getValue(); //linha correspondente em Vendas Dados


  if (vendas.getRange('AK3').getValue() > 0) {
    Browser.msgBox("Erro", "Necessário preencher todos os campos essenciais!", Browser.Buttons.OK);
  }

  else {

    vendasDados.deleteRow(linhaPedido);
    vendas.getRangeList(['D4', 'C13', 'C16', 'J11', 'M10']).clear({ contentsOnly: true, skipFilteredRows: true });
    Browser.msgBox("Informativo", "Registro Deletado com sucesso!", Browser.Buttons.OK);

    vendas.getRange('C16').activate();

  }

};


//******************    Finalizador   ******************************************************************


function finalizadorPedido1() {

  var spreadsheet = SpreadsheetApp.getActive();
  var vendas = spreadsheet.getSheetByName('Vendas');

  if (vendas.getRange('AL3').getValue() == 1) {
    SalvarPedido1();
  }

  else if (vendas.getRange('AL3').getValue() == 2) {
    editarPedido1();
  }

  else {
    deletarPedido1();
  }


};