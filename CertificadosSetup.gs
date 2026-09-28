/**
 * CertificadosSetup.gs
 * Adiciona as colunas ID_Certificado e Situacao_PJ na aba Controle_Certificados
 * e preenche ID_Certificado com IDs unicos de 8 caracteres (formato compativel AppSheet UNIQUEID).
 * NAO altera nenhuma outra aba.
 */

function setupControleCertificados() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Controle_Certificados');
  if (!sheet) { throw new Error('Aba Controle_Certificados nao encontrada'); }

  var lastRow = sheet.getLastRow();      // inclui cabecalho
  var lastCol = sheet.getLastColumn();
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Localiza (ou cria) as colunas ID_Certificado e Situacao_PJ
  var idCol = header.indexOf('ID_Certificado');
  var sitCol = header.indexOf('Situação_PJ');

  if (idCol === -1) {
    lastCol += 1;
    sheet.getRange(1, lastCol).setValue('ID_Certificado');
    idCol = lastCol - 1;
  }
  if (sitCol === -1) {
    lastCol += 1;
    sheet.getRange(1, lastCol).setValue('Situação_PJ');
    sitCol = lastCol - 1;
  }

  var dataRows = lastRow - 1;
  if (dataRows < 1) { return 'Sem linhas de dados.'; }

  // Gera IDs unicos apenas onde ainda estiver vazio
  var idRange = sheet.getRange(2, idCol + 1, dataRows, 1);
  var existing = idRange.getValues();
  var seen = {};
  for (var i = 0; i < existing.length; i++) {
    var v = existing[i][0];
    if (v) { seen[v] = true; }
  }
  var out = [];
  var gerados = 0;
  for (var r = 0; r < dataRows; r++) {
    if (existing[r][0]) { out.push([existing[r][0]]); continue; }
    var id;
    do { id = gerarId8(); } while (seen[id]);
    seen[id] = true;
    out.push([id]);
    gerados++;
  }
  idRange.setValues(out);

  return 'OK. Linhas de dados: ' + dataRows + ' | IDs gerados: ' + gerados +
         ' | ID_Certificado col ' + (idCol + 1) + ' | Situacao_PJ col ' + (sitCol + 1);
}

function gerarId8() {
  var chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  var s = '';
  for (var i = 0; i < 8; i++) {
    s += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return s;
}

/**
 * Restaura os cabecalhos Porte_Real e Solicitar_PDF_Faturamento na aba Empresas,
 * adicionando-os como novas colunas ao final, SEM alterar nenhuma coluna existente.
 */


function restaurarCabecalhosEmpresas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Empresas');
  if (!sheet) { throw new Error('Aba Empresas nao encontrada'); }
  var desired = ['Porte_Real', 'Solicitar_PDF_Faturamento', 'Fat_Total_Periodo'];
  var log = [];
  desired.forEach(function(name) {
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (headers.indexOf(name) !== -1) { log.push(name + ' JA EXISTE col ' + (headers.indexOf(name)+1)); return; }
    var maxCols = sheet.getMaxColumns();
    if (maxCols <= lastCol) { sheet.insertColumnsAfter(maxCols, 1); }
    var target = lastCol + 1;
    sheet.getRange(1, target).setValue(name);
    SpreadsheetApp.flush();
    log.push(name + ' ADICIONADO col ' + target);
  });
  var finalLastCol = sheet.getLastColumn();
  var finalHeaders = sheet.getRange(1, 1, 1, finalLastCol).getValues()[0];
  var result = 'ACOES: ' + log.join(' | ') + ' || TOTAL COLS: ' + finalLastCol + ' || ULTIMOS: ' + finalHeaders.slice(-4).join(', ');
  Logger.log(result);
  return result;
}