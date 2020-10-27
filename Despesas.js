/* ********************************************  Inicio Nova Despesa ******************************************* */
//Reformulação
function ReformularDespesaNova() {
  var spreadsheet = SpreadsheetApp.getActive();

  //Limpar
  spreadsheet
    .getRangeList(["G5", "G12", "H6:H8", "K5", "L6:L8", "L14:L17", "K25"])
    .clear({ contentsOnly: true, skipFilteredRows: true });

  spreadsheet
    .getRange("D4")
    .setBackground("#ff0000")
    .clearDataValidations()
    .setFormula('=IF(G5="";"";MAX(\'Despesas Dados\'!A2:A)+1)');
  spreadsheet
    .getRange("D5")
    .setFormula(
      "=IF(G5=\"\";\"\";IF(COUNTIF('Despesas Dados'!C2:C;G5) >= 1;LOOKUP(G5;'Despesas Dados'!C2:C;'Despesas Dados'!B2:B);MAX('Despesas Dados'!B2:B)+1))"
    );

  spreadsheet.getRange("H6").setFormula('=IF(G5="";"";Today())');
  spreadsheet.getRange("H7").setFormula('=IF(G5="";"";Today())');

  spreadsheet.getRangeList(["K13", "L13"]).setValue("False");

  spreadsheet.getRange("L15").setFormula('=IF(L14="";"";AE2)');
  spreadsheet.getRange("L16").setFormula('=IF(L14="";"";AF2)');
  spreadsheet.getRange("L17").setFormula('=IF(L14="";"";AG2)');
}

function ModoNovaDespesa() {
  var spreadsheet = SpreadsheetApp.getActive();
  //Formulação
  spreadsheet.getRange("AL3").setValue(1);
  spreadsheet.getRange("D1").setValue("Novo");

  // spreadsheet.getRange('AN3').setFormula('=IF(G5="";"";QUERY(\'Clientes Dados\'!A:M; "SELECT * WHERE \'"&G5&"\' = C "))');

  ReformularDespesaNova();

  spreadsheet.getRange("G25").setFormula('=IF(L8="";"";H6)');
  spreadsheet.getRange("I25").setFormula('=IF(L8="";"";L8)');
  spreadsheet.getRange("M25").setFormula('=IF(L8="";"";L8-I25)');
  spreadsheet
    .getRange("H22")
    .setFormula('=IF(L8="";"";MAX(\'Despesas Dados\'!Q2:Q)+1)'); //ID Pagamento

  spreadsheet.getRange("G5").activate();
}
/*  -------------------------------------------------------------------------------------------- */

/* ************************ Salvar Despesa ******************** */

// Função auxiliar de SalvarDespesa
function SalvarDespesa2() {
  let spreadsheet = SpreadsheetApp.getActive();
  let proprietario = spreadsheet.getRange("AH3").getValue();
  let tipo = spreadsheet.getRange("L6").getValue();

  if (proprietario == "" && tipo == "Frota") {
    Browser.msgBox(
      "Erro",
      "Necessário marcar uma das opções: Próprio ou Terceiros! ",
      Browser.Buttons.OK
    );
  } else {
    let frotaCampos = spreadsheet.getRange("AM3").getValue(); //testar se os campos de frotas estão vazios
    if (frotaCampos > 0 && tipo == "Frota") {
      Browser.msgBox(
        "Erro",
        "Necessário preencher os campos do Veículo! ",
        Browser.Buttons.OK
      );
    } else {
      let dadosDespesas = spreadsheet.getSheetByName("Despesas Dados");
      let Form = spreadsheet.getSheetByName("Despesas");

      let values = [
        [
          Form.getRange("D4").getValue(), // ID Despesa
          Form.getRange("D5").getValue(), // ID Fornecedor
          Form.getRange("G5").getValue(), // Fornecedor
          Form.getRange("H6").getValue(), // Data Despesa
          Form.getRange("H7").getValue(), // Data Emissão
          Form.getRange("H8").getValue(), // Nota Fiscal
          Form.getRange("G12").getValue(), // Descrição
          Form.getRange("K5").getValue(), // Plano de Conta
          tipo, // Tipo
          Form.getRange("L7").getValue(), // Parcela
          Form.getRange("L8").getValue(), // Valor
          proprietario, // Proprietário
          Form.getRange("L14").getValue(), // Placa
          Form.getRange("L15").getValue(), // Categoria
          Form.getRange("L16").getValue(), // Fabricante
          Form.getRange("L17").getValue(), // Modelo
          Form.getRange("H22").getValue(), // ID Pagamento
          Form.getRange("G25").getValue(), // Data Pagamento
          Form.getRange("I25").getValue(), // Valor Pagamento
          Form.getRange("K25").getValue(), // Forma Pagamento
          Form.getRange("M25").getValue(), // Restante
        ],
      ];

      return (
        dadosDespesas
          .getRange(dadosDespesas.getLastRow() + 1, 1, 1, 21)
          .setValues(values),
        Browser.msgBox(
          "Informativo",
          "Registro salvo com sucesso!",
          Browser.Buttons.OK
        ),
        ReformularDespesaNova(), // Constatar finalidade
        spreadsheet.getRange("G5").activate()
      );
    }
  }
}

// Função condutora para salvar Despesa

function SalvarDespesa() {
  var spreadsheet = SpreadsheetApp.getActive();

  if (spreadsheet.getRange("AK3").getValue() > 0) {
    Browser.msgBox(
      "Erro",
      "Necessário preencher todos os campos: Datas, Plano de Contas, Tipo, Parcela e Valor! ",
      Browser.Buttons.OK
    );
  } else {
    SalvarDespesa2();
  }
}

/* ************************ Término Nova Despesa ******************************************* */
/* ************************    Deletar Despesa ******************************************** */
function ReformularDeletarDespesa() {
  var spreadsheet = SpreadsheetApp.getActive();

  spreadsheet.getRange("H6").setFormula('=IF(D4="";"";BI5)'); //Data Lançamento
  spreadsheet.getRange("H7").setFormula('=IF(D4="";"";BJ5)'); //Data Emissão
  spreadsheet.getRange("H8").setFormula('=IF(D4="";"";BK5)'); //Nota Fiscal
  spreadsheet.getRange("G12").setFormula('=IF(D4="";"";BL5)'); //Descrição
  spreadsheet.getRange("K5").setFormula('=IF(D4="";"";BM5)'); //Plano de Conta
  spreadsheet.getRange("L6").setFormula('=IF(D4="";"";BN5)'); //Tipo
  spreadsheet.getRange("L7").setFormula('=IF(D4="";"";BO5)'); //Parcela
  spreadsheet.getRange("L8").setFormula('=IF(D4="";"";BP5)'); //Valor
  spreadsheet
    .getRange("K13")
    .setFormula('=IF(AND(BQ5="Próprio";L6="Frota");TRUE())'); //Proprietário
  spreadsheet
    .getRange("L13")
    .setFormula('=IF(AND(BQ5="Terceiros";L6="Frota");TRUE())'); //Proprietário
  spreadsheet.getRange("L14").setFormula('=IF(D4="";"";BR5)'); //Placa
  spreadsheet.getRange("L15").setFormula('=IF(D4="";"";BS5)'); //Categoria
  spreadsheet.getRange("L16").setFormula('=IF(D4="";"";BT5)'); //Fabricante
  spreadsheet.getRange("L17").setFormula('=IF(D4="";"";BU5)'); //Modelo
  spreadsheet.getRange("H22").setFormula('=IF(D4="";"";BV5)'); //ID Pagamento
  spreadsheet.getRange("G25").setFormula('=IF(D4="";"";BW5)'); //Data Pagamento
  spreadsheet.getRange("I25").setFormula('=IF(D4="";"";BX5)'); //Valor Pagamento
  spreadsheet.getRange("K25").setFormula('=IF(D4="";"";BY5)'); //Forma Pagto
  spreadsheet.getRange("M25").setFormula('=IF(D4="";"";BZ5)'); //Restante
}

function ModoDeletarDespesa() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange("AL3").setValue(3);
  spreadsheet.getRange("D1").setValue("Deletar");
  //Query consulta Despesa
  spreadsheet
    .getRange("BF4")
    .setFormula(
      '=IF(G5="";QUERY(\'Despesas Dados\'!A:U;"SELECT *");IF(D4="";QUERY(\'Despesas Dados\'!A:U;"SELECT * WHERE "&D5&" = B");QUERY(\'Despesas Dados\'!A:U;"SELECT * WHERE "&D5&" = B AND "&D4&" = A")))'
    );
  //Query ID Fornecedor
  spreadsheet
    .getRange("AN3")
    .setFormula(
      '=IF(G5="";"";QUERY(\'Despesas Dados\'!A:C; "SELECT * WHERE \'"&G5&"\' = C "))'
    );
  //Limpar
  spreadsheet
    .getRangeList(["D4", "G5", "G12", "H6:H8", "K5", "L6:L8", "L14:L17", "K25"])
    .clear({ contentsOnly: true, skipFilteredRows: true });

  spreadsheet
    .getRange("D4")
    .setBackground("#ffffff")
    .activate()
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInRange(spreadsheet.getRange("'Despesas'!$CB$5:$CB"), true)
        .build()
    ); //ID Despesa

  spreadsheet.getRange("D5").setFormula('=IF(G5="";"";AO4)'); //ID Fornecedor
  //spreadsheet.getRange('D6').setFormula('=IF(D4="";"";BH5)'); //Canal de Despesa

  spreadsheet
    .getRange("G5")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .setAllowInvalid(true)
        .requireValueInRange(
          spreadsheet.getRange("'Despesas Dados'!$C$2:$C"),
          true
        )
        .build()
    );

  ReformularDeletarDespesa();

  spreadsheet.getRange("G5").activate();
}

//*** ***** Consolidar Deletar Despesas

function DeletarDespesa() {
  var spreadsheet = SpreadsheetApp.getActive();

  if (spreadsheet.getRange("AK3").getValue() > 0) {
    Browser.msgBox(
      "Erro",
      "Necessário selecionar o registro a ser deletado! ",
      Browser.Buttons.OK
    );
  } else {
    var despesasDados = spreadsheet.getSheetByName("Despesas Dados");
    var linha = spreadsheet.getRange("AI3").getValue();

    despesasDados.deleteRow(linha);

    //Limpar
    spreadsheet
      .getRangeList(["D4", "G5"])
      .clear({ contentsOnly: true, skipFilteredRows: true });

    Browser.msgBox("Informativo", "Registro deletado!", Browser.Buttons.OK);

    //Reformular
    //ReformularDeletarDespesa();

    spreadsheet.getRange("G5").activate();
  }
}

/* *********************************************  Término Despesas ********************************************** */

//******************    Finalizador   ******************************************************************

function FinalizadorDespesa() {
  var spreadsheet = SpreadsheetApp.getActive();

  if (spreadsheet.getRange("AL3").getValue() == 1) {
    SalvarDespesa();
  } else if (spreadsheet.getRange("AL3").getValue() == 2) {
    EditarDespesa();
  } else {
    DeletarDespesa();
  }
}
