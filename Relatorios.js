function relatoriosPedidosDialog() {
  
    var url = 'https://datastudio.google.com/embed/reporting/9619e806-d091-4b3c-8d35-555ce5b7d744/page/Xn1JB';
    var name = 'Pedidos por Status';
    var url2 = 'https://datastudio.google.com/embed/reporting/a4d2055e-8a6f-47b4-9f28-650ff423fa5b/page/feyJB';
    var name2 = 'Cadastro de Clientes';  
    var url3 = 'https://datastudio.google.com/embed/reporting/f47b572f-c980-4c2b-aaaf-0f7de1410ca0/page/jw1OB';
    var name3 = 'Vendas Por Motoristas'; 
    var url4 = 'https://datastudio.google.com/embed/reporting/c4e43e3f-f2d2-4cd3-b93b-d3e70ecc4c30/page/p7ZQB';
    var name4 = 'Estoque'; 

    var html = '<html><body><a href="'+url+'" target="blank" onclick="google.script.host.close()">'+name+'</a> <br><br/><a href="'+url2+'" target="blank" onclick="google.script.host.close()">'+name2+'</a><br><br/><a href="'+url3+'" target="blank" onclick="google.script.host.close()">'+name3+'</a><br><br/><a href="'+url4+'" target="blank" onclick="google.script.host.close()">'+name4+'</a></body></html>';
    var ui = HtmlService.createHtmlOutput(html)
    SpreadsheetApp.getUi().showModelessDialog(ui,"Relatórios de Pedidos");
  }