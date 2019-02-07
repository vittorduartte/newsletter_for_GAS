function newsletterSend() {
  
  /* Chamamos o nosso objeto global SpreadsheetApp e invocamos o método getActive() para selecionarmos a planilha atual. */
  var sps = SpreadsheetApp.getActive();
  
  /* O método getSheets() do objeto global na variável sps retorna outro objeto do tipo Sheet.
  Enquanto a nossa variável data é um array de de arrays (em C é algo como um vetor), onde vamos percorrer as informações dela índice a índice.*/  
  var sheet = sps.getSheets()[0];
  var data = sps.getDataRange().getValues();
  
  /* O método getLastRow retorna o último índice de leitura, no entanto a contagem dos índice do array i
  iniciam com zero por isso subtraímos 1, enquanto nesse caso em particular não necessitamos da nossa linha de cabeçalho,
  por isso definimos nosso ponto de partida na variável firstRow como 1. */
  var firstRow = 1;
  var lastRow = sheet.getLastRow() - 1; 
  
  /* Definimos aqui o escopo do nosso email na qual vamos extrair de um documento de texto as informações.
  Para isso chamamos o nosso objeto global DocumentApp seguido dos métodos openById, onde inserimos o id do documento em questão,
  em seguida o método getBody onde nos retorna o corpo do nosso documento e finalizando com getText onde nos retorna o texto do corpo
  do nosso documento. */
  var id = 'ID do seu documento'; // >>> ou se prefirir uso o método openByName e insira o nome do Documento Google.'
  var document_body = DocumentApp.openById(id).getBody().getText();
  
  
  /* Como citado acima o nosso vetor de informações é na realidade uma matriz, vamos percorrer as linhas dessa matriz e a partir
     da nossa linha vamos acessar as colunas que correspondem aos índices da variável row. */    
  for(var i = firstRow; i <= lastRow; i++){
    
    var row = data[i]; 
    var email_address = row[2];
    var name = row[1];
   
  /* Aqui outra vez chamamos mais um objeto global dessa vez para acessar o nosso email, o objeto GmailApp e o método sendEmail que recebe 3 parâmetros:
     1. Endereço de Email
     2. Assunto do Email
     3. Escopo do Email
  */
    GmailApp.sendEmail(email_address, 'Olá, '+name+', tudo bem?', document_body);

  }
  
}


 /* A função onOpen é uma função que é acionada automaticamente sempre que abrirmos a planilha atual. */
function onOpen(){
 
 /* Aqui vamos acessar mais uma vez o método do nosso objeto global SpreadsheetApp, o método getActiveSpreadsheet ou acessamos o documento atual. */
 var sheet = SpreadsheetApp.getActiveSpreadsheet();
 
 /* A partir daqui vamos acrescentar um menu à nossa planilha em que irá nos dar a opção de acionar a nossa função principal a newsletterSend.
    O método do nosso objeto global sheet, addMenu, se encarrega disso onde recebe como parâmetro um nome para o menu e um objeto de submenus. */
 var subMenus = [{
    name : 'Newsletter Run',
    functionName : 'newsletterSend'
  }];
 
 sheet.addMenu('Newsletter', subMenus)
}
