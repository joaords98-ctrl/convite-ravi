/**
 * Recebe as confirmações do convite do Ravi e grava numa aba da planilha.
 *
 * COMO INSTALAR
 * 1. Crie a planilha no Google Sheets (nome livre).
 * 2. Menu Extensões > Apps Script. Apague o que estiver lá e cole este arquivo.
 * 3. Salve (disquete). Depois: Implantar > Nova implantação.
 *    - Tipo: "App da Web"
 *    - Executar como: "Eu"
 *    - Quem tem acesso: "Qualquer pessoa"
 * 4. Autorize quando o Google pedir (vai avisar que o app não é verificado —
 *    clique em "Avançado" > "Acessar <nome do projeto>"). É seu próprio script.
 * 5. Copie a URL que termina em /exec e me mande, ou cole você mesmo
 *    no index.html, na linha:  var PLANILHA = "";
 */

var ABA = 'Confirmações';

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var aba = planilha.getSheetByName(ABA) || criaAba(planilha);

    var d = (e && e.parameter) ? e.parameter : {};
    var vai = String(d.vai) === 'true';

    aba.appendRow([
      new Date(),
      d.nome || '',
      d.fone || '',
      vai ? 'Vai' : 'Não vai',
      Number(d.adultos || 0),
      Number(d.criancas || 0),
      Number(d.bebes || 0),
      d.recado || ''
    ]);

    return ok({ status: 'ok' });
  } catch (erro) {
    return ok({ status: 'erro', mensagem: String(erro) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ok({ status: 'no ar' });
}

function criaAba(planilha) {
  var aba = planilha.insertSheet(ABA);
  aba.appendRow([
    'Quando respondeu', 'Nome', 'WhatsApp', 'Presença',
    'Adultos', 'Crianças 3-10', 'Até 2 anos', 'Recado'
  ]);
  aba.getRange('A1:H1').setFontWeight('bold').setBackground('#e8e0c8');
  aba.setFrozenRows(1);
  aba.setColumnWidth(1, 160);
  aba.setColumnWidth(2, 200);
  aba.setColumnWidth(8, 320);
  return aba;
}

function ok(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
