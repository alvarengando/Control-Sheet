/* ********************************************  Inicio Nova Venda ******************************************* */
//Reformulação
function ReformularVendaNova() {
  var spreadsheet = SpreadsheetApp.getActive();

  //Limpar
  spreadsheet
    .getRangeList(["G5", "H6:H8", "K4:K8", "K25", "G11:M15"])
    .clear({ contentsOnly: true, skipFilteredRows: true });

  spreadsheet
    .getRange("D4")
    .setBackground("#5da68a")
    .clearDataValidations()
    .setFormula('=IF(G5="";"";MAX(\'Vendas Dados\'!A2:A)+1)');
  spreadsheet
    .getRange("D5")
    .setFormula(
      "=IF(G5=\"\";\"\";IF(COUNTIF('Vendas Dados'!C2:C;G5) >= 1;LOOKUP(G5;'Vendas Dados'!C2:C;'Vendas Dados'!B2:B);MAX('Vendas Dados'!B2:B)+1))"
    );

  spreadsheet.getRange("H6").setFormula('=IF(G5="";"";Today())');

  spreadsheet.getRange("K7").setFormula('=IF(K6="";"";L7)');
  spreadsheet.getRange("K8").setFormula('=IF(K7="";""; K6*K7)');
}

function ModoNovaVenda() {
  var spreadsheet = SpreadsheetApp.getActive();
  //Formulação
  spreadsheet.getRange("AL3").setValue(1);
  spreadsheet.getRange("D1").setValue("Novo");

  // spreadsheet.getRange('AN3').setFormula('=IF(G5="";"";QUERY(\'Clientes Dados\'!A:M; "SELECT * WHERE \'"&G5&"\' = C "))');

  ReformularVendaNova();

  spreadsheet.getRange("G18").setFormula('=IF(G11="";"";COUNTA(G11:G16))');
  spreadsheet.getRange("I18").setFormula('=IF(G11="";"";SUM(J11:J16))');
  spreadsheet
    .getRange("K18")
    .setFormula('=IF(G11="";"";SUMPRODUCT(J11:J16;K11:K16))');
  //spreadsheet.getRange('L18').setFormula('=IF(G11="";"";SUM(M11:M16))');

  spreadsheet.getRange("G25").setFormula('=IF(G5="";"";H6)');
  spreadsheet.getRange("I25").setFormula('=IF(K18="";"";K18)');
  spreadsheet.getRange("M25").setFormula('=IF(I25="";"";K18-I25)');
  spreadsheet
    .getRange("H22")
    .setFormula('=IF(G18="";"";MAX(\'Vendas Dados\'!M2:M)+1)'); //ID Pagamento

  spreadsheet.getRange("G5").activate();
}

/* ************************ Salvar Venda ******************** */

// Função auxiliar de SalvarVenda
function salvarVenda2(x, idCliente, idCompra, idRecebimento) {
  var spreadsheet = SpreadsheetApp.getActive();
  var dadosVendas = spreadsheet.getSheetByName("Vendas Dados");
  var Form = spreadsheet.getSheetByName("Vendas");

  var values = [
    [
      idCompra,                       // ID Venda
      idCliente,                       // ID Cliente
      Form.getRange("G5").getValue(), // Cliente
      Form.getRange("H6").getValue(), // Data Venda
      Form.getRange("H7").getValue(), // Motorista
      Form.getRange("H8").getValue(), // Placa
      Form.getRange(x, 7).getValue(), // ID Produto
      Form.getRange(x, 8).getValue(), // Marca
      Form.getRange(x, 9).getValue(), // Produto
      Form.getRange(x, 10).getValue(), // Quantidade
      Form.getRange(x, 11).getValue(), // Custo de Venda
      Form.getRange(x, 12).getValue(), // Total de Venda
      idRecebimento, // ID Recebimento
      Form.getRange("G25").getValue(), // Data Recebimento
      Form.getRange("K25").getValue(), // Forma de Pagamento
      Form.getRange("I25").getValue(), // Valor recebido
      Form.getRange("M25").getValue(),
    ],
  ]; // Restante

  return dadosVendas
    .getRange(dadosVendas.getLastRow() + 1, 1, 1, 17)
    .setValues(values);
}

// Função condutora para salvar Venda

function SalvarVenda() {
  var spreadsheet = SpreadsheetApp.getActive();
  var Form = spreadsheet.getSheetByName("Vendas");
  var idCompra = Form.getRange("D4").getValue();
  var idCliente = Form.getRange("D5").getValue();
  var idRecebimento = spreadsheet.getRange("H22").getValue();
  var quantLinhasProd = spreadsheet.getRange("AJ3").getValue();

  if (spreadsheet.getRange("AK3").getValue() > 0) {
    Browser.msgBox(
      "Erro",
      "Necessário preencher todos os campos referente ao cliente, produto e recebimentos! ",
      Browser.Buttons.OK
    );
  } else {
    let aux = 11;
    for (let index = 1; index <= quantLinhasProd; index++) {
      salvarVenda2(aux, idCliente, idCompra, idRecebimento);
      aux++;
    }
    ReformularVendaNova();
    LimparProdVendas();
    Browser.msgBox(
      "Informativo",
      "Registro salvo com sucesso!",
      Browser.Buttons.OK
    );
    
    spreadsheet.getRange("G5").activate();
  }
}

/* ******************************************** Término Nova Venda ******************************************* */

/* ******************************************* Início Deletar Vendas ********************************************** */

//Modo Deletar

function ReformularDeletarVenda() {
  var spreadsheet = SpreadsheetApp.getActive();

  spreadsheet.getRange("H6").setFormula('=IF(D4="";"";BI5)'); //Data Venda
  spreadsheet.getRange("H7").setFormula('=IF(D4="";"";BJ5)'); //Entregador
  spreadsheet.getRange("H8").setFormula('=IF(D4="";"";BK5)'); //Veículo

  spreadsheet.getRange("K25").setFormula('=IF(D4="";"";BT5)'); //Forma de Pagamento

  //spreadsheet.getRange('K7').setFormula('=IF(K6="";"";L7)');  //Preço de Venda
  //spreadsheet.getRange('K8').setFormula('=IF(K7="";""; K6*K7)');// Total de Venda

  //Formular aŕea de produto

  var values = [
    [
      '=IF($D$4="";"";BL5)',
      '=IF($D$4="";"";BM5)',
      '=IF($D$4="";"";BN5)',
      '=IF($D$4="";"";BO5)',
      '=IF($D$4="";"";BP5)',
      '=IF($D$4="";"";BQ5)',
    ],
    [
      '=IF($D$4="";"";BL6)',
      '=IF($D$4="";"";BM6)',
      '=IF($D$4="";"";BN6)',
      '=IF($D$4="";"";BO6)',
      '=IF($D$4="";"";BP6)',
      '=IF($D$4="";"";BQ6)',
    ],
    [
      '=IF($D$4="";"";BL7)',
      '=IF($D$4="";"";BM7)',
      '=IF($D$4="";"";BN7)',
      '=IF($D$4="";"";BO7)',
      '=IF($D$4="";"";BP7)',
      '=IF($D$4="";"";BQ7)',
    ],
    [
      '=IF($D$4="";"";BL8)',
      '=IF($D$4="";"";BM8)',
      '=IF($D$4="";"";BM8)',
      '=IF($D$4="";"";BO8)',
      '=IF($D$4="";"";BP8)',
      '=IF($D$4="";"";BQ8)',
    ],
    [
      '=IF($D$4="";"";BL9)',
      '=IF($D$4="";"";BM9)',
      '=IF($D$4="";"";BN9)',
      '=IF($D$4="";"";BO9)',
      '=IF($D$4="";"";BP9)',
      '=IF($D$4="";"";BQ9)',
    ],
    [
      '=IF($D$4="";"";BL10)',
      '=IF($D$4="";"";BM10)',
      '=IF($D$4="";"";BN10)',
      '=IF($D$4="";"";BO10)',
      '=IF($D$4="";"";BP10)',
      '=IF($D$4="";"";BQ10)',
    ],
  ];

  spreadsheet.getRange("G11:L16").setValues(values);
}

function ModoDeletarVenda() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange("AL3").setValue(3);
  spreadsheet.getRange("D1").setValue("Deletar");
  //Query consulta Venda
  spreadsheet
    .getRange("BF4")
    .setFormula(
      '=IF(G5="";QUERY(\'Vendas Dados\'!A:R;"SELECT *");IF(D4="";QUERY(\'Vendas Dados\'!A:R;"SELECT * WHERE "&D5&" = B");QUERY(\'Vendas Dados\'!A:R;"SELECT * WHERE "&D5&" = B AND "&D4&" = A")))'
    );
  //Query ID Cliente
  spreadsheet
    .getRange("AN3")
    .setFormula(
      '=IF(G5="";"";QUERY(\'Vendas Dados\'!A:C; "SELECT * WHERE \'"&G5&"\' = C "))'
    );
  //Limpar
  spreadsheet
    .getRangeList(["D4", "G5", "H6:H8", "K4:K8", "K25", "G11:M15"])
    .clear({ contentsOnly: true, skipFilteredRows: true });

  spreadsheet
    .getRange("D4")
    .setBackground("#ffffff")
    .activate()
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInRange(spreadsheet.getRange("'Vendas'!$BZ$5:$BZ"), true)
        .build()
    ); //ID Venda

  spreadsheet.getRange("D5").setFormula('=IF(G5="";"";AO4)'); //ID Cliente
  //spreadsheet.getRange('D6').setFormula('=IF(D4="";"";BH5)'); //Canal de Venda

  spreadsheet
    .getRange("G5")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .setAllowInvalid(true)
        .requireValueInRange(
          spreadsheet.getRange("'Vendas Dados'!$C$2:$C"),
          true
        )
        .build()
    );

  ReformularDeletarVenda();

  spreadsheet.getRange("H22").setFormula('=IF(D4="";"";BR5)'); //ID Pagamento

  spreadsheet.getRange("G25").setFormula('=IF(D4="";"";BS5)'); //Data Pagamento
  spreadsheet.getRange("I25").setFormula('=IF(D4="";"";BU5)'); //Valor Pago
  spreadsheet.getRange("K25").setFormula('=IF(D4="";"";BT5)'); //Forma de Pagamento
  spreadsheet.getRange("M25").setFormula('=IF(D4="";"";BV5)'); //Restante

  spreadsheet.getRange("G5").activate();
}

//**** Consolidar Deletar Vendas

function DeletarVenda() {
  var spreadsheet = SpreadsheetApp.getActive();

  if (spreadsheet.getRange("AK3").getValue() > 0) {
    Browser.msgBox(
      "Erro",
      "Necessário preencher os campos com ' * ' ",
      Browser.Buttons.OK
    );
  } else {
    var vendasDados = spreadsheet.getSheetByName("Vendas Dados");
    var LINHA = spreadsheet.getRange("AI3").getValue();
    var quantLinhas = spreadsheet.getRange("AJ3").getValue();

    vendasDados.deleteRows(LINHA, quantLinhas);

    //Limpar
    spreadsheet
      .getRangeList(["D4", "G5"])
      .clear({ contentsOnly: true, skipFilteredRows: true });

    Browser.msgBox("Informativo", "Registro deletado!", Browser.Buttons.OK);

    //Reformular
    ReformularDeletarVenda();

    spreadsheet.getRange("G5").activate();
  }
}

/* *********************************************  Término Vendas ********************************************** */

/* ******************************************** Função Inserir produto ***************************************** */
function InserirProdutoVenda2(x) {
  var spreadsheet = SpreadsheetApp.getActive();
  var Form = spreadsheet.getSheetByName("Vendas");

  var values = [
    [
      Form.getRange("L5").getValue(), // ID Produto
      Form.getRange("K4").getValue(), // Marca
      Form.getRange("K5").getValue(), // Produto
      Form.getRange("K6").getValue(), // Quantidade
      Form.getRange("K7").getValue(), // Preço de Venda
      //Form.getRange('K7').getValue(),    // Preço de Venda
      Form.getRange("K8").getValue(),
    ],
  ]; // Total

  return (
    Form.getRange(x).setValues(values),
    spreadsheet
      .getRangeList(["K4:K6"])
      .clear({ contentsOnly: true, skipFilteredRows: true }),
    spreadsheet.getRange("K7").setFormula('=IF(K6="";"";L7)'),
    spreadsheet.getRange("K25").activate()
  );
}

function InserirProdutoVenda() {
  var spreadsheet = SpreadsheetApp.getActive();

  if (spreadsheet.getRange("AM3").getValue() > 0) {
    Browser.msgBox(
      "Erro:",
      "Necessário preencher os campos de produto!",
      Browser.Buttons
    );
    spreadsheet.getRange("K4").activate();
  } else {
    if (spreadsheet.getRange("G11").getValue() == "") {
      InserirProdutoVenda2("G11:L11");
    } else if (spreadsheet.getRange("G12").getValue() == "") {
      InserirProdutoVenda2("G12:L12");
    } else if (spreadsheet.getRange("G13").getValue() == "") {
      InserirProdutoVenda2("G13:L13");
    } else if (spreadsheet.getRange("G14").getValue() == "") {
      InserirProdutoVenda2("G14:L14");
    } else if (spreadsheet.getRange("G15").getValue() == "") {
      InserirProdutoVenda2("G15:L15");
    } else if (spreadsheet.getRange("G16").getValue() == "") {
      InserirProdutoVenda2("G16:L16");
    } else {
      Browser.msgBox(
        "Erro:",
        "Todas as linhas foram preenchidas, finalize a Venda!",
        Browser.Buttons
      );
    }
    //spreadsheet.getRange('K8').setFormula('=IF(K7="";""; K6*K7)');
    //spreadsheet.getRange('K7').setFormula('=IF(K6="";"";L7)');
  }
}

//******************    Finalizador   ******************************************************************

function FinalizadorVenda() {
  var spreadsheet = SpreadsheetApp.getActive();

  if (spreadsheet.getRange("AL3").getValue() == 1) {
    SalvarVenda();
  } else if (spreadsheet.getRange("AL3").getValue() == 2) {
    EditarCompra();
  } else {
    DeletarVenda();
  }
}

//*******************************************************************************************************

/*  ---------  Auxilio de Finalizadores ----------    */

// Declaração
Array.prototype.findIndex = function (Procura) {
  if (Procura == "") return false;
  for (var i = 0; i < this.length; i++) if (this[i] == Procura) return i;
  return -i;
};

/* ******************************************** Limpar  ******************************************* */
function LimparProdVendas() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet
    .getRangeList(["G11:M16", "K25", "K4:K7"])
    .clear({ contentsOnly: true, skipFilteredRows: true });
}

/* ************************************************************************************************** */

function pagReltorioVendas() {
  var spreadsheet = SpreadsheetApp.getActive();
  //var Vendas = spreadsheet.getSheetByName('Vendas');

  spreadsheet.setActiveSheet(
    spreadsheet.getSheetByName("Relatório Vendas"),
    true
  );
}

function RelatorioVendas() {
  var spreadsheet = SpreadsheetApp.getActive();
  var url =
    "https://datastudio.google.com/reporting/a4d2055e-8a6f-47b4-9f28-650ff423fa5b/page/feyJB";
  var html =
    "<script> window.open('" + url + "');google.script.host.close();</script>";
  var userInterface = HtmlService.createHtmlOutput(html);

  SpreadsheetApp.getUi().showModalDialog(
    userInterface,
    "Relatório de Vendas..."
  );
}
// Exemplo de criar abrir relotario por uma caixa de dialogo ou direto
  
  function RelatoriosVendasDialog() {
    
    var url = 'https://datastudio.google.com/reporting/e2471175-38dc-4eb1-a28f-62c850a9d178/page/feyJB';
    var name = 'Vendas Consolidado';

    var url2 = 'https://datastudio.google.com/reporting/8bc163d7-8140-4abb-b6d3-f8d708d64d7c/page/feyJB';
    var name2 = 'Vendas Analítico';  

    var html = '<html><body><a href="'+url+'" target="blank" onclick="google.script.host.close()">'+name+'</a> <br><br/><a href="'+url2+'" target="blank" onclick="google.script.host.close()">'+name2+'</a></body></html>';
    var ui = HtmlService.createHtmlOutput(html)
    SpreadsheetApp.getUi().showModelessDialog(ui,"Relatórios de Vendas");
  }
  
  // function RelatorioVendas(){
  
  //   var spreadsheet = SpreadsheetApp.getActive();
  //   var url = "https://datastudio.google.com/embed/reporting/a4d2055e-8a6f-47b4-9f28-650ff423fa5b/page/feyJB"
  //   var html = "<script> window.open('"+ url + "');google.script.host.close();</script>";
  //   var userInterface = HtmlService.createHtmlOutput(html);
    
  //   SpreadsheetApp.getUi().showModalDialog(userInterface, "Relatório de Vendas...");
  // }
  
  
