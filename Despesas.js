/* ********************************************  Inicio Nova Venda ******************************************* */
//Reformulação
function ReformularDespesaNova() {
  var spreadsheet = SpreadsheetApp.getActive();

  //Limpar
  spreadsheet
    .getRangeList(["G5","G12", "H6:H8", "K5", "M6:M8"])
    .clear({ contentsOnly: true, skipFilteredRows: true });

  spreadsheet
    .getRange("D4")
    .setBackground("#cc0000")
    .clearDataValidations()
    .setFormula('=IF(G5="";"";MAX(\'Despesas Dados\'!A2:A)+1)');
  spreadsheet
    .getRange("D5")
    .setFormula(
      "=IF(G5=\"\";\"\";IF(COUNTIF('Despesas Dados'!C2:C;G5) >= 1;LOOKUP(G5;'Despesas Dados'!C2:C;'Despesas Dados'!B2:B);MAX('Despesas Dados'!B2:B)+1))"
    );

  spreadsheet.getRange("H6").setFormula('=IF(G5="";"";Today())');
  spreadsheet.getRange("H7").setFormula('=IF(G5="";"";Today())');

}

function ModoNovaDespesa() {
  var spreadsheet = SpreadsheetApp.getActive();
  //Formulação
  spreadsheet.getRange("AL3").setValue(1);
  spreadsheet.getRange("D1").setValue("Novo");

  // spreadsheet.getRange('AN3').setFormula('=IF(G5="";"";QUERY(\'Clientes Dados\'!A:M; "SELECT * WHERE \'"&G5&"\' = C "))');

  ReformularDespesaNova();

  spreadsheet.getRange("G25").setFormula('=IF(G5="";"";H6)');
  spreadsheet.getRange("I25").setFormula('=IF(K18="";"";K18)');
  spreadsheet.getRange("M25").setFormula('=IF(I25="";"";K18-I25)');
  spreadsheet
    .getRange("H22")
    .setFormula('=IF(G18="";"";MAX(\'Despesas Dados\'!Q2:Q)+1)'); //ID Pagamento

  spreadsheet.getRange("G5").activate();
}
