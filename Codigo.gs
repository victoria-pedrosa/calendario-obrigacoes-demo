const SHEET_EVENTOS = 'Calendario_Eventos';
const SHEET_REGRAS = 'Regras_Obrigacoes_Tributarias';

function resetTudo() {
  const ss = getSpreadsheet();
  const antigoEventos = ss.getSheetByName(SHEET_EVENTOS);
  if (antigoEventos) ss.deleteSheet(antigoEventos);
  const antigoRegras = ss.getSheetByName(SHEET_REGRAS);
  if (antigoRegras) ss.deleteSheet(antigoRegras);
  inicializarPlanilhas();
  gerarCalendarioCompleto();
}

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet(nome, headers) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(nome);
  if (!sheet) {
    sheet = ss.insertSheet(nome);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function inicializarPlanilhas() {
  getOrCreateSheet(SHEET_EVENTOS, ['Data', 'Descricao', 'Tipo', 'Abrangencia', 'Responsavel', 'Observacao']);
  const regrasSheet = getOrCreateSheet(SHEET_REGRAS, ['Obrigacao', 'Periodicidade', 'Mes_Offset', 'Dia_Vencimento', 'Enesimo_Dia_Util', 'Mes_Fixo', 'Ultimo_Dia_Util', 'Ajustar_Dia_Util', 'Responsavel', 'Observacao']);
  if (regrasSheet.getLastRow() === 1) {
    const defaults = [
      ['DAS - Simples Nacional', 'Mensal', 1, 20, '', '', false, true, 'Fiscal', 'Recolhimento mensal do Simples Nacional (PGDAS-D). Verificar vigencia da regra na legislacao atual.'],
      ['FGTS', 'Mensal', 1, 7, '', '', false, true, 'Folha', 'Recolhimento mensal do FGTS. Verificar vigencia da regra na legislacao atual.'],
      ['GPS / INSS', 'Mensal', 1, 20, '', '', false, true, 'Folha', 'Guia da Previdencia Social para empresas com folha de pagamento. Verificar vigencia da regra na legislacao atual.'],
      ['DCTFWeb', 'Mensal', 2, '', 15, '', false, false, 'Fiscal', 'Declaracao de Debitos e Creditos Tributarios Federais. Vencimento no 15o dia util. Verificar vigencia da regra na legislacao atual.'],
      ['DEFIS', 'Anual', 0, '', '', 3, false, false, 'Contabil', 'Declaracao de Informacoes Socioeconomicas e Fiscais - vencimento anual em 31/03. Verificar vigencia da regra na legislacao atual.'],
      ['ECF', 'Anual', 0, '', '', 7, true, false, 'Contabil', 'Escrituracao Contabil Fiscal - vencimento no ultimo dia util de julho. Verificar vigencia da regra na legislacao atual.']
    ];
    defaults.forEach(function(row){ regrasSheet.appendRow(row); });
  }
}

function pascoa(ano) {
  const a = ano % 19;
  const b = Math.floor(ano / 100);
  const c = ano % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const mes = Math.floor((h + l - 7 * m + 114) / 31);
  const dia = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(ano, mes - 1, dia);
}

function addDias(data, dias) {
  const d = new Date(data.getTime());
  d.setDate(d.getDate() + dias);
  return d;
}

function feriadosDoAno(ano) {
  const pasc = pascoa(ano);
  const carnaval = addDias(pasc, -47);
  const sextaSanta = addDias(pasc, -2);
  const corpusChristi = addDias(pasc, 60);
  return [
    { data: new Date(ano, 0, 1), descricao: 'Confraternizacao Universal', abrangencia: 'Nacional' },
    { data: carnaval, descricao: 'Carnaval', abrangencia: 'Nacional' },
    { data: sextaSanta, descricao: 'Sexta-feira Santa', abrangencia: 'Nacional' },
    { data: corpusChristi, descricao: 'Corpus Christi', abrangencia: 'Nacional' },
    { data: new Date(ano, 3, 21), descricao: 'Tiradentes', abrangencia: 'Nacional' },
    { data: new Date(ano, 4, 1), descricao: 'Dia do Trabalho', abrangencia: 'Nacional' },
    { data: new Date(ano, 8, 7), descricao: 'Independencia do Brasil', abrangencia: 'Nacional' },
    { data: new Date(ano, 9, 12), descricao: 'Nossa Senhora Aparecida', abrangencia: 'Nacional' },
    { data: new Date(ano, 10, 2), descricao: 'Finados', abrangencia: 'Nacional' },
    { data: new Date(ano, 10, 15), descricao: 'Proclamacao da Republica', abrangencia: 'Nacional' },
    { data: new Date(ano, 10, 20), descricao: 'Dia Nacional de Zumbi e da Consciencia Negra', abrangencia: 'Nacional' },
    { data: new Date(ano, 11, 25), descricao: 'Natal', abrangencia: 'Nacional' },
    { data: new Date(ano, 6, 2), descricao: 'Independencia da Bahia (2 de Julho)', abrangencia: 'Estadual (BA)' },
    { data: new Date(ano, 11, 8), descricao: 'Nossa Senhora da Conceicao da Praia', abrangencia: 'Municipal (Salvador)' }
  ];
}

function chaveDataStr(data) {
  return Utilities.formatDate(data, Session.getScriptTimeZone() || 'GMT-3', 'yyyy-MM-dd');
}

function feriadosComoSet(ano) {
  const set = {};
  [ano - 1, ano, ano + 1].forEach(function(y) {
    feriadosDoAno(y).forEach(function(f) { set[chaveDataStr(f.data)] = true; });
  });
  return set;
}

function ehDiaUtil(data, feriadosSet) {
  const dia = data.getDay();
  if (dia === 0 || dia === 6) return false;
  if (feriadosSet[chaveDataStr(data)]) return false;
  return true;
}

function proximoDiaUtil(data, feriadosSet) {
  let d = new Date(data.getTime());
  while (!ehDiaUtil(d, feriadosSet)) {
    d = addDias(d, 1);
  }
  return d;
}

function diaUtilAnteriorOuIgual(data, feriadosSet) {
  let d = new Date(data.getTime());
  while (!ehDiaUtil(d, feriadosSet)) {
    d = addDias(d, -1);
  }
  return d;
}

function enesimoDiaUtilDoMes(ano, mes, n, feriadosSet) {
  let contador = 0;
  let d = new Date(ano, mes, 1);
  while (true) {
    if (ehDiaUtil(d, feriadosSet)) {
      contador++;
      if (contador === n) return d;
    }
    d = addDias(d, 1);
  }
}

function ultimoDiaUtilDoMes(ano, mes, feriadosSet) {
  let d = new Date(ano, mes + 1, 0);
  while (!ehDiaUtil(d, feriadosSet)) {
    d = addDias(d, -1);
  }
  return d;
}

function lerRegras() {
  const sheet = getOrCreateSheet(SHEET_REGRAS, ['Obrigacao', 'Periodicidade', 'Mes_Offset', 'Dia_Vencimento', 'Enesimo_Dia_Util', 'Mes_Fixo', 'Ultimo_Dia_Util', 'Ajustar_Dia_Util', 'Responsavel', 'Observacao']);
  const dados = sheet.getDataRange().getValues();
  const linhas = dados.slice(1);
  return linhas.filter(function(l){ return l[0]; }).map(function(l) {
    return {
      obrigacao: l[0], periodicidade: l[1], mesOffset: Number(l[2] || 0), diaVencimento: l[3], enesimoDiaUtil: l[4],
      mesFixo: l[5], ultimoDiaUtilFlag: l[6] === true, ajustarDiaUtil: l[7] === true, responsavel: l[8] || '', observacao: l[9] || ''
    };
  });
}

function eventosObrigacoesDoAno(ano, regras, feriadosSet) {
  const eventos = [];
  regras.forEach(function(r) {
    if (r.periodicidade === 'Mensal') {
      for (let mes = 0; mes < 12; mes++) {
        let mesVenc = mes + r.mesOffset;
        let anoVenc = ano;
        if (mesVenc > 11) { mesVenc -= 12; anoVenc += 1; }
        let data;
        if (r.enesimoDiaUtil) {
          data = enesimoDiaUtilDoMes(anoVenc, mesVenc, Number(r.enesimoDiaUtil), feriadosSet);
        } else {
          data = new Date(anoVenc, mesVenc, Number(r.diaVencimento));
          if (r.ajustarDiaUtil) data = proximoDiaUtil(data, feriadosSet);
        }
        eventos.push({ data: data, descricao: r.obrigacao, tipo: 'Obrigacao Tributaria', abrangencia: 'Nacional', responsavel: r.responsavel, observacao: r.observacao });
      }
    } else if (r.periodicidade === 'Anual') {
      let data;
      if (r.ultimoDiaUtilFlag) {
        data = ultimoDiaUtilDoMes(ano, Number(r.mesFixo) - 1, feriadosSet);
      } else {
        data = new Date(ano, Number(r.mesFixo) - 1, Number(r.diaVencimento));
        if (r.ajustarDiaUtil) data = proximoDiaUtil(data, feriadosSet);
      }
      eventos.push({ data: data, descricao: r.obrigacao, tipo: 'Obrigacao Tributaria', abrangencia: 'Nacional', responsavel: r.responsavel, observacao: r.observacao });
    }
  });
  return eventos;
}

function popularAno(ano) {
  const sheet = getOrCreateSheet(SHEET_EVENTOS, ['Data', 'Descricao', 'Tipo', 'Abrangencia', 'Responsavel', 'Observacao']);
  const feriadosSet = feriadosComoSet(ano);
  const existentes = sheet.getDataRange().getValues().slice(1);
  const chaves = {};
  existentes.forEach(function(l) {
    if (!l[0]) return;
    const dstr = (l[0] instanceof Date) ? chaveDataStr(l[0]) : String(l[0]);
    chaves[dstr + '|' + l[1]] = true;
  });

  const eventos = [];
  feriadosDoAno(ano).forEach(function(f) {
    eventos.push({ data: f.data, descricao: f.descricao, tipo: (f.abrangencia === 'Nacional' ? 'Feriado Nacional' : 'Feriado Estadual'), abrangencia: f.abrangencia, responsavel: 'Todos', observacao: '' });
  });
  const regras = lerRegras();
  eventosObrigacoesDoAno(ano, regras, feriadosSet).forEach(function(e) { eventos.push(e); });

  let criados = 0;
  const novasLinhas = [];
  eventos.forEach(function(e) {
    const dstr = chaveDataStr(e.data);
    const chave = dstr + '|' + e.descricao;
    if (!chaves[chave]) {
      novasLinhas.push([e.data, e.descricao, e.tipo, e.abrangencia, e.responsavel, e.observacao]);
      chaves[chave] = true;
      criados++;
    }
  });
  if (novasLinhas.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, novasLinhas.length, 6).setValues(novasLinhas);
  }
  return criados;
}

function ordenarPlanilhaEventos() {
  const sheet = getOrCreateSheet(SHEET_EVENTOS, ['Data', 'Descricao', 'Tipo', 'Abrangencia', 'Responsavel', 'Observacao']);
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet.getRange(2, 1, lastRow - 1, 6).sort({ column: 1, ascending: true });
  }
}

function gerarCalendarioCompleto() {
  inicializarPlanilhas();
  const anoAtual = new Date().getFullYear();
  const criadosAtual = popularAno(anoAtual);
  const criadosProximo = popularAno(anoAtual + 1);
  ordenarPlanilhaEventos();
  Logger.log('Eventos criados: ano atual = ' + criadosAtual + ', ano seguinte = ' + criadosProximo);
}

function atualizarCalendarioMensal() {
  gerarCalendarioCompleto();
}

function criarGatilhoMensal() {
  const triggers = ScriptApp.getProjectTriggers();
  const jaExiste = triggers.some(function(t) { return t.getHandlerFunction() === 'atualizarCalendarioMensal'; });
  if (jaExiste) {
    Logger.log('Gatilho mensal ja existia.');
    return;
  }
  ScriptApp.newTrigger('atualizarCalendarioMensal').timeBased().onMonthDay(1).atHour(0).create();
  Logger.log('Gatilho mensal criado.');
}



// === Criado por assistente: cria as 3 novas abas de estrutura ===
function criarAbasBalanco() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var defs = {
    'Tarefas_Setoriais': ['ID_Tarefa','Setor','Empresa','Obrigacao','Responsavel_Entrega','Prazo_Interno','Vencimento_Legal','Data_Entrega_Efetiva','Forma_Envio','Status'],
    'Controle_Extratos_Balanco': ['ID_Extrato','Empresa','Responsavel_Entrega','Mes_Referencia','Prazo_Interno','Data_Entrega','Observacoes','Status'],
    'Controle_Nibo_Balanco': ['ID_Nibo','Empresa','Socio_Responsavel','Tentativa_Contato','Data_Ultimo_Contato','Adesao_Nibo','Data_Adesao']
  };
  var criadas = [];
  for (var nome in defs) {
    var sh = ss.getSheetByName(nome);
    if (!sh) { sh = ss.insertSheet(nome); criadas.push(nome); }
    var headers = defs[nome];
    sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  Logger.log('Abas processadas. Criadas agora: ' + (criadas.join(', ') || 'nenhuma (ja existiam)'));
  return criadas;
}

function normalizarDatasCalendario() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Calendario_Eventos");
  if (!sh) throw new Error("aba Calendario_Eventos nao encontrada");
  var last = sh.getLastRow();
  if (last < 2) return;
  var rng = sh.getRange(2, 1, last - 1, 1);
  var vals = rng.getValues();
  var out = vals.map(function(r){
    var d = r[0];
    if (!(d instanceof Date)) {
      var s = String(d).trim();
      var parts;
      if (s.indexOf("-") > -1) { parts = s.split(" ")[0].split("-"); return [new Date(Number(parts[0]), Number(parts[1])-1, Number(parts[2]), 12,0,0)]; }
      if (s.indexOf("/") > -1) { parts = s.split(" ")[0].split("/"); return [new Date(Number(parts[2]), Number(parts[1])-1, Number(parts[0]), 12,0,0)]; }
      return [r[0]];
    }
    return [new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0)];
  });
  rng.setNumberFormat("dd/MM/yyyy");
  rng.setValues(out);
  SpreadsheetApp.flush();
}



function diagFaturamento2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var db = ss.getSheetByName('Faturamento_DB');
  var d = db.getDataRange().getValues();
  // sum Total_RS by Codigo
  var sumByCod = {};
  for (var i=1;i<d.length;i++){ var cod=String(d[i][2]); var t=Number(d[i][7])||0; sumByCod[cod]=(sumByCod[cod]||0)+t; }
  // Empresas: find Codigo Dominio col and Empresa col
  var emp = ss.getSheetByName('Empresas');
  var e = emp.getDataRange().getValues();
  var hdr = e[0];
  var ci = hdr.indexOf('Código Dominio'); if(ci<0) ci=hdr.indexOf('Codigo Dominio');
  var ei = hdr.indexOf('Empresa');
  var companies = [];
  for (var j=1;j<e.length;j++){
    var cod=String(e[j][ci]);
    var total = sumByCod[cod]||0;
    companies.push({cod:cod, emp:e[j][ei], total:total});
  }
  companies.sort(function(a,b){return b.total-a.total;});
  var me = companies.filter(function(c){return c.total>0 && c.total<=360000;});
  var epp = companies.filter(function(c){return c.total>360000 && c.total<=4800000;});
  var out = {
    empresasRows: companies.length,
    top10ME: me.slice(0,10).map(function(c){return [c.emp, Math.round(c.total)];}),
    top10EPP: epp.slice(0,10).map(function(c){return [c.emp, Math.round(c.total)];})
  };
  Logger.log(JSON.stringify(out));
  return out;
}

function materializarFaturamento() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var db = ss.getSheetByName('Faturamento_DB');
  var d = db.getDataRange().getValues();
  var sumByCod = {};
  for (var i=1;i<d.length;i++){ var cod=String(d[i][2]).trim(); var t=Number(d[i][7])||0; if(cod!=='') sumByCod[cod]=(sumByCod[cod]||0)+t; }
  var emp = ss.getSheetByName('Empresas');
  var rng = emp.getDataRange();
  var e = rng.getValues();
  var hdr = e[0];
  var ci = hdr.indexOf('Código Dominio'); if(ci<0) ci=hdr.indexOf('Codigo Dominio');
  // find or create helper columns at end
  var colFat = hdr.indexOf('Fat_Total_Periodo');
  var colPorte = hdr.indexOf('Porte_Real');
  var lastCol = hdr.length;
  if (colFat<0){ colFat=lastCol; lastCol++; }
  if (colPorte<0){ colPorte=lastCol; lastCol++; }
  // build output arrays
  var nRows = e.length;
  var fatArr = [], porteArr = [];
  for (var j=0;j<nRows;j++){
    if (j===0){ fatArr.push(['Fat_Total_Periodo']); porteArr.push(['Porte_Real']); continue; }
    var cod=String(e[j][ci]).trim();
    var total = sumByCod[cod]||0;
    var porte = total<=360000 ? 'ME' : (total<=4800000 ? 'EPP' : 'Outro');
    fatArr.push([total]);
    porteArr.push([porte]);
  }
  emp.getRange(1, colFat+1, nRows, 1).setValues(fatArr);
  emp.getRange(1, colPorte+1, nRows, 1).setValues(porteArr);
  Logger.log('OK cols: Fat_Total_Periodo='+(colFat+1)+' Porte_Real='+(colPorte+1)+' rows='+nRows);
  return 'done';
}