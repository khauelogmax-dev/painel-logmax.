/**
 * LogMax · Tela do Operador (marcar chegada)
 *
 * Este é um ARQUIVO SEPARADO no mesmo Apps Script da planilha.
 * Ele NÃO substitui nem mexe no código do painel (doGet), que continua igual.
 * Todos os nomes aqui começam com OP_ / op para não brigar com o código que já existe.
 *
 * O que ele faz (só com a senha certa):
 *   - acao "entrar":   confere a senha.
 *   - acao "ler":      devolve a escala (loja, data, horário, código, nome, presença e nº da linha).
 *                      NÃO devolve telefone, CPF, PIX nem valores.
 *   - acao "chegou":   escreve OK na coluna PRESENÇA daquela linha.
 *   - acao "desfazer": volta para N/D.
 *   - acao "codigo":   escreve o CÓDIGO do motoboy na coluna COD (para preencher uma vaga ou trocar quem está).
 *                      O NOME a planilha puxa sozinha pela fórmula. Não mexe no OK.
 *
 * Senha: Configurações do projeto (engrenagem) → Propriedades do script → SENHA_EDICAO.
 * Só escreve na coluna PRESENÇA e só depois de conferir que a linha ainda é a mesma pessoa no mesmo dia.
 */
const OP_ABAS_IGNORADAS = ['RESUMO', 'FOLHA DE PAGAMENTO', 'CADASTRO MOTORISTA', 'ENDEREÇOS', 'PAG FNAL'];
const OP_DIAS_ANTES = 2, OP_DIAS_DEPOIS = 7;
const OP_LINHAS_FINAIS = 700;
const OP_COLUNAS = 8;
const OP_GUARDAR_SEGUNDOS = 20;
const OP_MAX_ERROS_SENHA = 15;   // 15 senhas erradas em 10 minutos → bloqueia 10 minutos

function doPost(e) {
  let pedido;
  try { pedido = JSON.parse(e.postData.contents); } catch (err) { return opResposta({ ok: false, erro: 'Pedido inválido.' }); }

  // 1) senha
  const cache = CacheService.getScriptCache();
  const erros = +(cache.get('op_erros_senha') || 0);
  if (erros >= OP_MAX_ERROS_SENHA) return opResposta({ ok: false, erro: 'Muitas senhas erradas. Espere 10 minutos.', senha: true });
  const senhaCerta = PropertiesService.getScriptProperties().getProperty('SENHA_EDICAO');
  if (!senhaCerta) return opResposta({ ok: false, erro: 'A senha ainda não foi criada nas Propriedades do script (SENHA_EDICAO).', senha: true });
  if (String(pedido.senha || '') !== senhaCerta) {
    cache.put('op_erros_senha', String(erros + 1), 600);
    return opResposta({ ok: false, erro: 'Senha errada.', senha: true });
  }

  // 2) entrar / ler
  if (pedido.acao === 'entrar') {
    return opResposta({ ok: true, planilha: SpreadsheetApp.getActiveSpreadsheet().getName() });
  }
  if (pedido.acao === 'ler') {
    let csv = cache.get('op_escala');
    if (!csv) {
      csv = opMontarCsv();
      try { cache.put('op_escala', csv, OP_GUARDAR_SEGUNDOS); } catch (err) { /* grande demais para guardar: tudo bem */ }
    }
    return opResposta({ ok: true, csv: csv });
  }

  if (['chegou', 'desfazer', 'codigo'].indexOf(pedido.acao) < 0) return opResposta({ ok: false, erro: 'Ação desconhecida.' });

  // 3) uma gravação por vez
  const trava = LockService.getScriptLock();
  if (!trava.tryLock(15000)) return opResposta({ ok: false, erro: 'A planilha está ocupada. Tente de novo.' });
  try {
    return opResposta(pedido.acao === 'codigo' ? opTrocarCodigo(pedido) : opMarcar(pedido));
  } finally {
    trava.releaseLock();
  }
}

function opColunas(l) {
  const achar = function (txt) { return l.findIndex(function (c) { return String(c).toUpperCase().indexOf(txt) >= 0; }); };
  const col = { cod: achar('COD'), nome: achar('NOME'), hora: achar('HORARIO'), pres: achar('PRESEN') };
  // abas sem o título "PRESENÇA" (ex.: PAMPLONA): usa a coluna logo depois do HORÁRIO
  if (col.pres < 0 && col.hora >= 0) col.pres = col.hora + 1;
  return col;
}

function opAbaDeLoja(aba) {
  const nome = aba.getName().trim();
  if (OP_ABAS_IGNORADAS.indexOf(nome.toUpperCase()) >= 0) return false;
  if (/^ontem/i.test(nome) || aba.isSheetHidden()) return false;
  return true;
}

function opMontarCsv() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const hoje = new Date();
  const min = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - OP_DIAS_ANTES);
  const max = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() + OP_DIAS_DEPOIS);
  const agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
  const saida = [['LOJA', 'DATA', 'HORARIO', 'CODIGO', 'NOME', 'PRESENCA', 'LINHA', 'ATUALIZADO']];

  planilha.getSheets().forEach(function (aba) {
    if (!opAbaDeLoja(aba)) return;
    const loja = aba.getName().trim();
    const ultima = aba.getLastRow();
    if (ultima < 2) return;
    const inicio = Math.max(1, ultima - OP_LINHAS_FINAIS + 1);
    const linhas = aba.getRange(inicio, 1, ultima - inicio + 1, Math.min(OP_COLUNAS, aba.getMaxColumns())).getDisplayValues();
    let col = null;

    linhas.forEach(function (l, i) {
      if (String(l[0]).trim().toUpperCase() === 'DATA') { col = opColunas(l); return; }
      const m = String(l[0]).trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
      if (!m || !col || col.pres < 0) return;
      const d = new Date(+m[3], +m[2] - 1, +m[1]);
      if (d < min || d > max) return;
      const pega = function (k) { return k >= 0 ? String(l[k]).trim() : ''; };
      saida.push([loja, String(l[0]).trim(), pega(col.hora), pega(col.cod), pega(col.nome), pega(col.pres), String(inicio + i), '']);
    });
  });

  if (saida.length > 1) saida[1][7] = agora;
  else saida.push(['', '', '', '', '', '', '', agora]);
  return saida.map(function (linha) {
    return linha.map(function (c) { return /[",\n]/.test(c) ? '"' + String(c).replace(/"/g, '""') + '"' : c; }).join(',');
  }).join('\n');
}

// acha a linha na aba e confere se ainda é a mesma pessoa no mesmo dia
function opLocalizar(p) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(String(p.loja || ''));
  if (!aba || !opAbaDeLoja(aba)) return { erro: { ok: false, erro: 'Não achei a aba da loja ' + p.loja + '.' } };

  const linha = parseInt(p.linha, 10);
  if (!(linha >= 2) || linha > aba.getLastRow()) return { erro: { ok: false, erro: 'Linha fora da planilha.', mudou: true } };

  // acha o cabeçalho do bloco desta linha (subindo no máximo 300 linhas)
  const de = Math.max(1, linha - 300);
  const trecho = aba.getRange(de, 1, linha - de + 1, Math.min(OP_COLUNAS, aba.getMaxColumns())).getDisplayValues();
  const atual = trecho[trecho.length - 1];
  let col = null;
  for (let i = trecho.length - 2; i >= 0; i--) {
    if (String(trecho[i][0]).trim().toUpperCase() === 'DATA') { col = opColunas(trecho[i]); break; }
  }
  if (!col || col.pres < 0 || col.nome < 0) return { erro: { ok: false, erro: 'Não achei o cabeçalho deste dia na aba ' + p.loja + '.' } };

  const igual = function (a, b) { return String(a).trim().toUpperCase() === String(b).trim().toUpperCase(); };
  if (!igual(atual[0], p.data) || !igual(atual[col.nome], p.nome)) {
    return { erro: { ok: false, mudou: true, erro: 'Essa linha mudou na planilha. Atualizei a lista, confira e tente de novo.' } };
  }
  return { aba: aba, linha: linha, atual: atual, col: col };
}

function opLimparCache() {
  const cache = CacheService.getScriptCache();
  cache.remove('op_escala');   // a tela do operador já lê a mudança
  cache.remove('presenca');    // o painel público também
}

function opMarcar(p) {
  const achou = opLocalizar(p);
  if (achou.erro) return achou.erro;
  const aba = achou.aba, linha = achou.linha, atual = achou.atual, col = achou.col;

  const celula = aba.getRange(linha, col.pres + 1);
  const agora = String(atual[col.pres]).trim().toUpperCase();

  if (p.acao === 'chegou') {
    if (agora === 'OK') return { ok: true, presenca: 'OK', jaEstava: true };
    if (agora !== '' && agora !== 'N/D') {
      return { ok: false, erro: 'Na planilha está escrito "' + atual[col.pres] + '" para essa pessoa. Mude direto na planilha.' };
    }
    celula.setValue('OK');
  } else {
    if (agora !== 'OK') return { ok: true, presenca: atual[col.pres], jaEstava: true };
    celula.setValue('N/D');
  }
  SpreadsheetApp.flush();
  opLimparCache();
  return { ok: true, presenca: p.acao === 'chegou' ? 'OK' : 'N/D' };
}

// escreve o código do motoboy na coluna COD; o NOME vem sozinho pela fórmula da planilha
function opTrocarCodigo(p) {
  const achou = opLocalizar(p);
  if (achou.erro) return achou.erro;
  const aba = achou.aba, linha = achou.linha, atual = achou.atual, col = achou.col;
  if (col.cod < 0) return { ok: false, erro: 'Não achei a coluna COD nesta aba.' };

  const novo = String(p.codigo || '').trim();
  if (!/^[0-9A-Za-z-]{1,12}$/.test(novo)) return { ok: false, erro: 'Digite um código válido (só números ou letras).' };
  if (String(atual[col.cod]).trim().toUpperCase() === novo.toUpperCase()) return { ok: true, nome: atual[col.nome], codigo: novo, jaEstava: true };

  // não deixa trocar quem já está com OK (o OK passaria para outra pessoa)
  if (String(atual[col.pres]).trim().toUpperCase() === 'OK') {
    return { ok: false, erro: 'Essa pessoa já está com OK. Toque em Desfazer antes de trocar.' };
  }

  const celNome = aba.getRange(linha, col.nome + 1);
  if (!celNome.getFormula()) {
    return { ok: false, erro: 'Nesta linha o NOME não puxa pelo código (não tem fórmula). Troque direto na planilha.' };
  }

  const celCod = aba.getRange(linha, col.cod + 1);
  const antes = celCod.getValue();
  celCod.setValue(/^\d+$/.test(novo) ? Number(novo) : novo);
  SpreadsheetApp.flush();

  const nome = String(celNome.getDisplayValue()).trim();
  if (!nome || /N[ÃA]O\s+LOCALIZAD|^#/i.test(nome)) {
    celCod.setValue(antes);              // código não existe no cadastro: volta como estava
    SpreadsheetApp.flush();
    return { ok: false, erro: 'Código ' + novo + ' não encontrado no cadastro. Nada foi mudado.' };
  }
  opLimparCache();
  return { ok: true, nome: nome, codigo: novo };
}

function opResposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Confere, em cada loja, se a coluna NOME puxa pelo código (fórmula). Não mostra nenhum dado pessoal. */
function conferirNomePorCodigo() {
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function (aba) {
    if (!opAbaDeLoja(aba)) return;
    const ultima = aba.getLastRow();
    if (ultima < 3) return;
    const inicio = Math.max(1, ultima - 60 + 1);
    const valores = aba.getRange(inicio, 1, ultima - inicio + 1, OP_COLUNAS).getDisplayValues();
    const formulas = aba.getRange(inicio, 1, ultima - inicio + 1, OP_COLUNAS).getFormulas();
    let col = null, com = 0, sem = 0;
    valores.forEach(function (l, i) {
      if (String(l[0]).trim().toUpperCase() === 'DATA') { col = opColunas(l); return; }
      if (!col || col.nome < 0 || !/^\d{2}\/\d{2}\/\d{4}$/.test(String(l[0]).trim())) return;
      formulas[i][col.nome] ? com++ : sem++;
    });
    Logger.log(aba.getName() + ': ' + com + ' linhas com fórmula no NOME, ' + sem + ' sem fórmula');
  });
}

/** Rode esta no editor para testar: mostra quantas linhas leu, o tempo e se a senha existe (sem mostrar a senha). */
function testarOperador() {
  const t = Date.now();
  const csv = opMontarCsv();
  Logger.log('Leu ' + (csv.split('\n').length - 1) + ' linhas em ' + Math.round((Date.now() - t) / 1000) + ' segundos.');
  const s = PropertiesService.getScriptProperties().getProperty('SENHA_EDICAO');
  Logger.log(s ? 'Senha criada, com ' + s.length + ' caracteres.' : 'Senha NÃO criada ainda.');
}
