/**
 * Recebe as confirmações do site da festa e grava na planilha
 * "Confirmações de Presença — Festa Francisco".
 *
 * COMO CONFIGURAR (só precisa fazer uma vez):
 * 1. Abra a planilha no Google Drive
 * 2. Menu Extensões → Apps Script
 * 3. Apague o conteúdo do editor e cole este arquivo inteiro
 * 4. Clique em "Implantar" → "Nova implantação"
 * 5. Tipo: "App da Web"
 *    - Executar como: Eu
 *    - Quem pode acessar: Qualquer pessoa
 * 6. Clique em "Implantar" e autorize com sua conta Google
 * 7. Copie a "URL do app da Web" (termina em /exec)
 * 8. Cole essa URL no campo SHEETS_ENDPOINT do arquivo
 *    "Festa Francisco.dc.html" (ou envie a URL para o Claude fazer isso)
 */

var SHEET_ID = '1XfQP-t3YzBkUKiR4nDlDdpxMUpNwyXfqrxDwgULqWEM';

function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];

  var gn = d.guestNames || {};
  var ga = d.guestAges || {};

  // Acompanhantes: todos os nomes menos o "Adulto 1" (que é a própria pessoa)
  var acompanhantes = [];
  var idades = [];
  Object.keys(gn).sort().forEach(function (k) {
    if (k === 'a0' || !gn[k]) return;
    acompanhantes.push(gn[k]);
    if (k.charAt(0) === 'c' && ga[k]) idades.push(gn[k] + ': ' + ga[k] + ' anos');
  });

  var vai = !!d.vai;
  sheet.appendRow([
    new Date(d.enviadoEm || Date.now()),
    d.nome || '',
    vai ? 'Sim' : 'Não',
    vai ? (d.adultos || 0) : 0,
    vai ? (d.criancas || 0) : 0,
    vai ? ((d.adultos || 0) + (d.criancas || 0)) : 0,
    acompanhantes.join(', '),
    idades.join(', '),
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
