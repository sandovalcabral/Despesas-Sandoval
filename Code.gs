/**
 * DESPESAS SANDOVAL - API em Google Apps Script
 * ------------------------------------------------------------
 * Liga a planilha (Google Sheets) ao Dashboard e ao App de Lancamento.
 *
 *  GET  ?action=dados&pin=...   -> todas as linhas da aba "Despesas"
 *  GET  ?action=listas&pin=...  -> listas (pessoas, formas de pagamento...)
 *  POST {action:"adicionar", ...}  -> grava 1 ou N parcelas
 *  POST {action:"marcarPago", ...} -> muda a Situacao de uma linha para "OK"
 *
 * Os PINs ficam nas "Propriedades do script" (nunca no codigo).
 * Use as funcoes criarConfig() e definirPins() uma unica vez.
 */

var ABA_DESPESAS = 'Despesas';
var ABA_CONFIG = 'Config';
var FUSO_PADRAO = 'America/Bahia';

// Nome do cabecalho na planilha (sem acento, minusculo) -> campo interno.
// ATENCAO: na planilha, "Categoria" guarda a PESSOA e "Cartao" a FORMA DE PAGAMENTO.
var CAMPOS = {
  dataCompra: 'data compra',
  descricao: 'descricao',
  pessoa: 'categoria',
  valor: 'valor',
  parcela: 'parcela',
  formaPagamento: 'cartao',
  observacoes: 'observacoes',
  mes: 'mes',
  ano: 'ano',
  situacao: 'situacao'
};

var SITUACAO_OK = 'OK';
var SITUACAO_ABERTO = 'A Vencer';

/* ============================================================
 * CONFIGURACAO INICIAL (rodar uma vez, pelo editor do Apps Script)
 * ============================================================ */

/** Cria a aba "Config" com as listas. Nao mexe em nenhuma aba existente. */
function criarConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(ABA_CONFIG)) {
    Logger.log('A aba Config j\u00e1 existe. Nada foi alterado.');
    return;
  }
  var sh = ss.insertSheet(ABA_CONFIG);
  var pessoas = ['Sandro', 'Carla', 'Lu Lando', 'Lilian', 'Tia Lilia', 'Marisa', 'Gil',
                 'Taislane', 'Gute', 'Lando', 'Henrique', 'Daniel', 'Jaci'];
  var formas = ['Amex', 'Boleto', 'Caixa', 'D\u00e9bito em Conta', 'Mercado Pago', 'Nubank', 'Pix'];
  var situacoes = [SITUACAO_ABERTO, SITUACAO_OK];

  sh.getRange(1, 1, 1, 3).setValues([['Pessoas', 'Formas de Pagamento', 'Situa\u00e7\u00f5es']]).setFontWeight('bold');
  sh.getRange(2, 1, pessoas.length, 1).setValues(pessoas.map(function (v) { return [v]; }));
  sh.getRange(2, 2, formas.length, 1).setValues(formas.map(function (v) { return [v]; }));
  sh.getRange(2, 3, situacoes.length, 1).setValues(situacoes.map(function (v) { return [v]; }));
  sh.setColumnWidths(1, 3, 170);
  Logger.log('Aba Config criada.');
}

/**
 * Define os PINs. Edite os dois valores abaixo, rode a funcao UMA vez
 * e depois APAGUE os valores digitados (volte para 'TROQUE-ME').
 */
function definirPins() {
  var pinLeitura = 'TROQUE-ME';      // PIN do Dashboard (leitura)
  var pinLancamento = 'TROQUE-ME';   // PIN do App de Lan\u00e7amento (leitura + grava\u00e7\u00e3o)

  if (pinLeitura === 'TROQUE-ME' || pinLancamento === 'TROQUE-ME') {
    throw new Error('Edite os PINs dentro da fun\u00e7\u00e3o antes de rodar.');
  }
  if (String(pinLeitura).length < 4 || String(pinLancamento).length < 4) {
    throw new Error('Use PINs com pelo menos 4 caracteres.');
  }
  PropertiesService.getScriptProperties().setProperties({
    PIN_LEITURA: String(pinLeitura),
    PIN_LANCAMENTO: String(pinLancamento)
  });
  Logger.log('PINs gravados. Agora apague os valores do c\u00f3digo.');
}

/* ============================================================
 * ENTRADAS DO WEB APP
 * ============================================================ */

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    var acao = p.action || 'dados';
    if (!pinValido(p.pin, 'leitura')) return resposta({ ok: false, erro: 'Acesso negado' });

    if (acao === 'dados') return resposta(acaoDados());
    if (acao === 'listas') return resposta(acaoListas());
    return resposta({ ok: false, erro: 'A\u00e7\u00e3o desconhecida' });
  } catch (err) {
    return resposta({ ok: false, erro: String(err.message || err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var d = JSON.parse(e.postData.contents);
    if (!pinValido(d.pin, 'escrita')) return resposta({ ok: false, erro: 'Acesso negado' });

    lock.waitLock(20000);
    var saida;
    if (d.action === 'adicionar') saida = acaoAdicionar(d);
    else if (d.action === 'marcarPago') saida = acaoMarcarPago(d);
    else saida = { ok: false, erro: 'A\u00e7\u00e3o desconhecida' };
    return resposta(saida);
  } catch (err) {
    return resposta({ ok: false, erro: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) { /* sem lock */ }
  }
}

/* ============================================================
 * ACOES
 * ============================================================ */

/** Devolve todas as linhas da aba Despesas, exatamente como estao. */
function acaoDados() {
  var sh = abaDespesas();
  var col = mapearColunas(sh);
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || FUSO_PADRAO;
  var linhas = lerLinhas(sh, col, tz);
  return {
    ok: true,
    atualizadoEm: Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss"),
    total: linhas.length,
    linhas: linhas
  };
}

/** Listas para os formularios: vem da aba Config; se faltar, dos proprios dados. */
function acaoListas() {
  var listas = lerListasConfig();
  var sh = abaDespesas();
  var col = mapearColunas(sh);
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || FUSO_PADRAO;
  var linhas = lerLinhas(sh, col, tz);

  var contagemObs = {};
  linhas.forEach(function (l) {
    var o = String(l.observacoes || '').trim();
    if (o) contagemObs[o] = (contagemObs[o] || 0) + 1;
  });
  var sugestoes = Object.keys(contagemObs)
    .sort(function (a, b) { return contagemObs[b] - contagemObs[a]; })
    .slice(0, 40);

  if (!listas.pessoas.length) listas.pessoas = unicos(linhas, 'pessoa');
  if (!listas.formas.length) listas.formas = unicos(linhas, 'formaPagamento');
  return {
    ok: true,
    pessoas: listas.pessoas,
    formasPagamento: listas.formas,
    situacoes: [SITUACAO_ABERTO, SITUACAO_OK],
    sugestoesObservacoes: sugestoes
  };
}

/** Grava 1 ou N parcelas ao final da aba Despesas (somente nas colunas dos 10 campos). */
function acaoAdicionar(d) {
  // Protecao contra envio duplicado (fila offline / duplo toque)
  if (d.clientId && clienteJaProcessado(d.clientId)) {
    return { ok: true, duplicado: true, inseridas: 0 };
  }

  var listas = lerListasConfig();
  var n = parseInt(d.parcelas, 10) || 1;
  var valor = Number(d.valor);
  var mes = parseInt(d.mes, 10);
  var ano = parseInt(d.ano, 10);
  var situacao = normalizarSituacao(d.situacao);

  if (!d.descricao || !String(d.descricao).trim()) throw new Error('Informe a descri\u00e7\u00e3o.');
  if (!isFinite(valor) || valor === 0) throw new Error('Valor inv\u00e1lido.');
  if (n < 1 || n > 24) throw new Error('Parcelas deve ficar entre 1 e 24.');
  if (!(mes >= 1 && mes <= 12) || !(ano >= 2000 && ano <= 2100)) throw new Error('M\u00eas/Ano inv\u00e1lidos.');
  if (!situacao) throw new Error('Situa\u00e7\u00e3o inv\u00e1lida.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d.dataCompra))) throw new Error('Data da compra inv\u00e1lida.');
  if (listas.pessoas.length && listas.pessoas.indexOf(d.pessoa) < 0) throw new Error('Pessoa fora da lista.');
  if (listas.formas.length && listas.formas.indexOf(d.formaPagamento) < 0) throw new Error('Forma de pagamento fora da lista.');

  var sh = abaDespesas();
  var col = mapearColunas(sh);
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || FUSO_PADRAO;
  var dataCompra = Utilities.parseDate(String(d.dataCompra), tz, 'yyyy-MM-dd');

  var colunas = Object.keys(col).map(function (k) { return col[k]; });
  var minCol = Math.min.apply(null, colunas);
  var largura = Math.max.apply(null, colunas) - minCol + 1;
  if (largura !== colunas.length) {
    throw new Error('As colunas da aba Despesas n\u00e3o est\u00e3o lado a lado. Grava\u00e7\u00e3o cancelada por seguran\u00e7a.');
  }

  var ultima = ultimaLinhaDados(sh, col);
  var proxima = ultima + 1;
  var precisa = proxima + n - 1;
  if (precisa > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), precisa - sh.getMaxRows());

  var matriz = [];
  var m = mes, a = ano;
  for (var i = 0; i < n; i++) {
    var linha = new Array(largura);
    for (var j = 0; j < largura; j++) linha[j] = '';
    linha[col.dataCompra - minCol] = dataCompra;
    linha[col.descricao - minCol] = String(d.descricao).trim();
    linha[col.pessoa - minCol] = d.pessoa;
    linha[col.valor - minCol] = valor;
    linha[col.parcela - minCol] = (i + 1) + '/' + n;
    linha[col.formaPagamento - minCol] = d.formaPagamento;
    linha[col.observacoes - minCol] = d.observacoes ? String(d.observacoes).trim() : '';
    linha[col.mes - minCol] = m;
    linha[col.ano - minCol] = a;
    linha[col.situacao - minCol] = (i === 0) ? situacao : SITUACAO_ABERTO;
    matriz.push(linha);
    m++;
    if (m > 12) { m = 1; a++; }
  }

  var destino = sh.getRange(proxima, minCol, n, largura);
  if (ultima >= 2) {
    // Copia a formatacao da ultima linha de dados (moeda, fonte, etc.)
    sh.getRange(ultima, minCol, 1, largura)
      .copyTo(destino, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }
  // "1/3" precisa ser TEXTO, senao o Sheets transforma em data
  sh.getRange(proxima, col.parcela, n, 1).setNumberFormat('@');
  sh.getRange(proxima, col.dataCompra, n, 1).setNumberFormat('dd/MM/yyyy');
  destino.setValues(matriz);
  SpreadsheetApp.flush();

  if (d.clientId) registrarCliente(d.clientId);
  return { ok: true, inseridas: n, primeiraLinha: proxima, ultimaLinha: proxima + n - 1 };
}

/** Marca uma linha como paga ("OK"), conferindo se e mesmo a linha certa. */
function acaoMarcarPago(d) {
  var linha = parseInt(d.linha, 10);
  if (!(linha >= 2)) throw new Error('Linha inv\u00e1lida.');

  var sh = abaDespesas();
  var col = mapearColunas(sh);
  var atual = function (campo) { return sh.getRange(linha, col[campo]).getValue(); };

  var confere =
    String(atual('descricao')).trim() === String(d.descricao || '').trim() &&
    Math.abs(Number(atual('valor')) - Number(d.valor)) < 0.005 &&
    Number(atual('mes')) === Number(d.mes) &&
    Number(atual('ano')) === Number(d.ano);
  if (!confere) {
    throw new Error('A linha mudou na planilha. Atualize a lista e tente de novo.');
  }
  if (normalizarSituacao(atual('situacao')) === SITUACAO_OK) {
    return { ok: true, jaEstavaPago: true, linha: linha };
  }
  sh.getRange(linha, col.situacao).setValue(SITUACAO_OK);
  SpreadsheetApp.flush();
  return { ok: true, linha: linha };
}

/* ============================================================
 * FUNCOES AUXILIARES
 * ============================================================ */

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function pinValido(pin, tipo) {
  var props = PropertiesService.getScriptProperties();
  var leitura = props.getProperty('PIN_LEITURA');
  var lanc = props.getProperty('PIN_LANCAMENTO');
  var p = String(pin || '');
  if (!p) return false;
  if (tipo === 'escrita') return !!lanc && p === lanc;
  return (!!leitura && p === leitura) || (!!lanc && p === lanc);
}

function norm(s) {
  return String(s == null ? '' : s).trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function normalizarSituacao(s) {
  var n = norm(s);
  if (n === 'ok') return SITUACAO_OK;
  if (n === 'a vencer') return SITUACAO_ABERTO;
  return '';
}

function abaDespesas() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_DESPESAS);
  if (!sh) throw new Error('Aba "' + ABA_DESPESAS + '" n\u00e3o encontrada.');
  return sh;
}

/** Descobre a coluna (1, 2, 3...) de cada campo pelo NOME do cabecalho (linha 1). */
function mapearColunas(sh) {
  var cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
  var col = {};
  Object.keys(CAMPOS).forEach(function (campo) {
    var idx = cab.indexOf(CAMPOS[campo]);
    if (idx < 0) throw new Error('Cabe\u00e7alho n\u00e3o encontrado na aba Despesas: "' + CAMPOS[campo] + '"');
    col[campo] = idx + 1;
  });
  return col;
}

/** Ultima linha com dado real (ignora a formatacao vazia ate a linha 965). */
function ultimaLinhaDados(sh, col) {
  var total = sh.getMaxRows();
  var desc = sh.getRange(1, col.descricao, total, 1).getValues();
  var val = sh.getRange(1, col.valor, total, 1).getValues();
  for (var i = total - 1; i >= 1; i--) {
    if (String(desc[i][0]).trim() !== '' || String(val[i][0]).trim() !== '') return i + 1;
  }
  return 1;
}

function lerLinhas(sh, col, tz) {
  var ultima = ultimaLinhaDados(sh, col);
  if (ultima < 2) return [];
  var colunas = Object.keys(col).map(function (k) { return col[k]; });
  var minCol = Math.min.apply(null, colunas);
  var largura = Math.max.apply(null, colunas) - minCol + 1;
  var dados = sh.getRange(2, minCol, ultima - 1, largura).getValues();
  var saida = [];

  dados.forEach(function (r, i) {
    var get = function (campo) { return r[col[campo] - minCol]; };
    var descricao = get('descricao');
    var valor = get('valor');
    if (String(descricao).trim() === '' && String(valor).trim() === '') return; // linha vazia

    var dc = get('dataCompra');
    var dataCompra = (dc instanceof Date)
      ? Utilities.formatDate(dc, tz, 'yyyy-MM-dd')
      : String(dc == null ? '' : dc); // texto fica como veio; o front converte

    saida.push({
      linha: i + 2,
      dataCompra: dataCompra,
      descricao: String(descricao),
      pessoa: String(get('pessoa')).trim(),
      valor: (typeof valor === 'number') ? valor : String(valor),
      parcela: String(get('parcela')),
      formaPagamento: String(get('formaPagamento')).trim(),
      observacoes: String(get('observacoes')),
      mes: numeroOuNulo(get('mes')),
      ano: numeroOuNulo(get('ano')),
      situacao: String(get('situacao')).trim()
    });
  });
  return saida;
}

function numeroOuNulo(v) {
  if (v === '' || v == null) return null;
  var n = Number(v);
  return isFinite(n) ? Math.round(n) : null;
}

function lerListasConfig() {
  var vazio = { pessoas: [], formas: [], situacoes: [] };
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_CONFIG);
  if (!sh || sh.getLastRow() < 2) return vazio;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  var coluna = function (i) {
    return v.map(function (r) { return String(r[i]).trim(); })
            .filter(function (x) { return x !== ''; });
  };
  return { pessoas: coluna(0), formas: coluna(1), situacoes: coluna(2) };
}

function unicos(linhas, campo) {
  var visto = {};
  var lista = [];
  linhas.forEach(function (l) {
    var v = String(l[campo] || '').trim();
    if (v && !visto[norm(v)]) { visto[norm(v)] = true; lista.push(v); }
  });
  return lista.sort();
}

/* ---- Anti-duplicidade (guarda os ultimos 300 clientId enviados) ---- */
function clienteJaProcessado(id) {
  var s = PropertiesService.getScriptProperties().getProperty('IDS_RECENTES');
  var lista = s ? JSON.parse(s) : [];
  return lista.indexOf(String(id)) >= 0;
}

function registrarCliente(id) {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('IDS_RECENTES');
  var lista = s ? JSON.parse(s) : [];
  lista.push(String(id));
  if (lista.length > 300) lista = lista.slice(lista.length - 300);
  props.setProperty('IDS_RECENTES', JSON.stringify(lista));
}

/* ---- Teste rapido pelo editor: confere leitura e cabecalhos ---- */
function testarLeitura() {
  var r = acaoDados();
  var soma = 0, ok = 0;
  r.linhas.forEach(function (l) {
    var v = Number(l.valor) || 0;
    soma += v;
    if (normalizarSituacao(l.situacao) === SITUACAO_OK) ok += v;
  });
  Logger.log('Linhas: ' + r.total);
  Logger.log('Total: ' + soma.toFixed(2) + ' | Pago (OK): ' + ok.toFixed(2) + ' | A Vencer: ' + (soma - ok).toFixed(2));
  Logger.log('Refer\u00eancia do arquivo de 12/10/2026 -> 698 | 57238.00 | 38745.13 | 18492.87');
}
