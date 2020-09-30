function ggff() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C32').activate()
  .setFormula('=IF(AND(C16="";G7="");"";IF(AP4="";COUNTA(\'Pedidos Dados\'!C2:C)+1;AP4))');
  spreadsheet.getRange('D32').activate()
  .setFormula('=IF(AND(C16="";G7="");"";IF(AP4="";COUNTA(\'Pedidos Dados\'!C2:C)+1;AP4))');
};

function frmmmss() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G31').activate()
  .setFormula('=IF(AND(C13="";C16="");"";IF(C16<>"";QUERY(\'Pedidos Dados\'!A:Q;"SELECT * WHERE "&C16&" = I");QUERY(\'Pedidos Dados\'!A:Q;"SELECT * WHERE \'"&C13&"\' = D")))');
};

function FRM2() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G31').activate()
  .setFormula('=IF(D4="";"";IF(C16<>"";QUERY(\'Pedidos Dados\'!A:Q;"SELECT * WHERE "&C16&" = I AND "&D4&" = A");QUERY(\'Pedidos Dados\'!A:Q;"SELECT * WHERE \'"&C13&"\' = D AND "&D4&" = A")))');
};

function frmsassffffffffffff() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AG12').activate()
  .setFormula('=IF(AM3="";""; IF(AM3="16";QUERY(\'Pedidos Dados\'!A:Q;"SELECT * WHERE "&C16&" = I");IF(AM3="164";QUERY(\'Pedidos Dados\'!A:Q;"SELECT * WHERE "&C16&" = I AND "&D4&" = A");IF(AM3="13";QUERY(\'Pedidos Dados\'!A:Q;"SELECT * WHERE \'"&C13&"\' = D");IF(AM3="134";QUERY(\'Pedidos Dados\'!A:Q;"SELECT * WHERE \'"&C13&"\' = D AND "&D4&" = A"))))))');
};