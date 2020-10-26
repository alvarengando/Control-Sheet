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
          Form.getRange("D4").getValue(),  // ID Despesa
          Form.getRange("D5").getValue(),  // ID Fornecedor
          Form.getRange("G5").getValue(),  // Fornecedor
          Form.getRange("H6").getValue(),  // Data Despesa
          Form.getRange("H7").getValue(),  // Data Emissão
          Form.getRange("H8").getValue(),  // Nota Fiscal
          Form.getRange("G12").getValue(), // Descrição
          Form.getRange("K5").getValue(),  // Plano de Conta
          tipo,                            // Tipo
          Form.getRange("L7").getValue(),  // Parcela
          Form.getRange("L8").getValue(),  // Valor
          proprietario,                    // Proprietário
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
        ReformularDespesaNova(),
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

/* ******************************************** Término Nova Despesa ******************************************* */

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
