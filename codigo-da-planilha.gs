/**
 * LogMax · Presença ao vivo + marcar chegada pelo site
 * Cole este código DENTRO da planilha (Extensões → Apps Script) e implante como "App da Web".
 *
 * LEITURA (doGet): a tela pública e a tela do operador chamam este endereço a cada 30 segundos.
 *   Entrega SÓ: loja, data, horário, código, nome, presença e o número da linha.
 *   NÃO entrega telefone, CPF, PIX nem valores.
 *
 * GRAVAÇÃO (doPost): a tela do operador manda "chegou" ou "desfazer".
 *   - Só funciona com a senha guardada em Configurações do projeto → Propriedades do script → SENHA_EDICAO.
 *   - Só escreve na coluna PRESENÇA, e só depois de conferir que a linha ainda é a mesma
 *     (mesma loja, mesma data, mesmo nome). Nunca toca em PIX, CPF ou valores.
 */
const ABAS_IGNORADAS = ['RESUMO', 'FOLHA DE PAGAMENTO', 'CADASTRO MOTORISTA', 'ENDEREÇOS', 'PAG FNAL'];
const DIAS_ANTES = 7, DIAS_DEPOIS = 7;
const GUARDAR_SEGUNDOS = 25;  // várias telas abertas ao mesmo tempo usam a mesma leitura
const LINHAS_FINAIS = 700;    // lê só o final de cada aba (onde ficam os dias recentes)
const COLUNAS = 8;            // DATA, COD, NOME, TELEFONE, HORARIO, PRESENÇA... (lê só até a coluna H)
const MAX_ERROS_SENHA = 15;   // depois de 15 senhas erradas em 10 minutos, bloqueia por 10 minutos

/* ======================= LEITURA ======================= */

function doGet() {
  const cache = CacheService.getScriptCache();
  let csv = cache.get('presenca');
  if (!csv) {
    csv = montarCsv();
    try { cache.put('presenca', csv, GUARDAR_SEGUNDOS); } catch (e) { /* grande demais para guardar: tudo bem, só fica mais lento */ }
  }
  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV);
}

// acha as colunas a partir da linha de cabeçalho (a que começa com "DATA")
function colunasDoCabecalho(l) {
  const achar = function (txt) { return l.findIndex(function (c) { return String(c).toUpperCase().indexOf(txt) >= 0; }); };
  const col = { cod: achar('COD'), nome: achar('NOME'), hora: achar('HORARIO'), pres: achar('PRESEN') };
  // abas sem o título "PRESENÇA" (ex.: PAMPLONA): usa a coluna logo depois do HORÁRIO
  if (col.pres < 0 && col.hora >= 0) col.pres = col.hora + 1;
  return col;
}

function abaDeLoja(aba) {
  const nome = aba.getName().trim();
  if (ABAS_IGNORADAS.indexOf(nome.toUpperCase()) >= 0) return false;
  if (/^ontem/i.test(nome) || aba.isSheetHidden()) return false;   // cópias de relatório e abas escondidas
  return true;
}

function montarCsv() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const hoje = new Date();
  const min = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - DIAS_ANTES);
  const max = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() + DIAS_DEPOIS);
  const agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
  const saida = [['LOJA', 'DATA', 'HORARIO', 'CODIGO', 'NOME', 'PRESENCA', 'LINHA', 'ATUALIZADO']];

  planilha.getSheets().forEach(function (aba) {
    if (!abaDeLoja(aba)) return;
    const loja = aba.getName().trim();
    const ultima = aba.getLastRow();
    if (ultima < 2) return;
    const inicio = Math.max(1, ultima - LINHAS_FINAIS + 1);
    const linhas = aba.getRange(inicio, 1, ultima - inicio + 1, Math.min(COLUNAS, aba.getMaxColumns())).getDisplayValues();
    let col = null; // posição das colunas do bloco atual

    linhas.forEach(function (l, i) {
      const primeira = String(l[0]).trim().toUpperCase();
      if (primeira === 'DATA') { col = colunasDoCabecalho(l); return; }
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

/* ======================= GRAVAÇÃO ======================= */

function doPost(e) {
  let pedido;
  try { pedido = JSON.parse(e.postData.contents); } catch (err) { return resposta({ ok: false, erro: 'Pedido inválido.' }); }

  // 1) senha
  const cache = CacheService.getScriptCache();
  const erros = +(cache.get('erros_senha') || 0);
  if (erros >= MAX_ERROS_SENHA) return resposta({ ok: false, erro: 'Muitas senhas erradas. Espere 10 minutos.', senha: true });
  const senhaCerta = PropertiesService.getScriptProperties().getProperty('SENHA_EDICAO');
  if (!senhaCerta) return resposta({ ok: false, erro: 'A senha ainda não foi criada nas Propriedades do script (SENHA_EDICAO).', senha: true });
  if (String(pedido.senha || '') !== senhaCerta) {
    cache.put('erros_senha', String(erros + 1), 600);
    return resposta({ ok: false, erro: 'Senha errada.', senha: true });
  }

  // 2) só conferir a senha (tela de entrada)
  if (pedido.acao === 'entrar') {
    return resposta({ ok: true, planilha: SpreadsheetApp.getActiveSpreadsheet().getName() });
  }

  if (pedido.acao !== 'chegou' && pedido.acao !== 'desfazer') return resposta({ ok: false, erro: 'Ação desconhecida.' });

  // 3) uma gravação por vez
  const trava = LockService.getScriptLock();
  if (!trava.tryLock(15000)) return resposta({ ok: false, erro: 'A planilha está ocupada. Tente de novo.' });
  try {
    return resposta(marcar(pedido));
  } finally {
    trava.releaseLock();
  }
}

function marcar(p) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName(String(p.loja || ''));
  if (!aba || !abaDeLoja(aba)) return { ok: false, erro: 'Não achei a aba da loja ' + p.loja + '.' };

  const linha = parseInt(p.linha, 10);
  if (!(linha >= 2) || linha > aba.getLastRow()) return { ok: false, erro: 'Linha fora da planilha.', mudou: true };

  // acha o cabeçalho do bloco desta linha (subindo no máximo 300 linhas)
  const de = Math.max(1, linha - 300);
  const trecho = aba.getRange(de, 1, linha - de + 1, Math.min(COLUNAS, aba.getMaxColumns())).getDisplayValues();
  const atual = trecho[trecho.length - 1];
  let col = null;
  for (let i = trecho.length - 2; i >= 0; i--) {
    if (String(trecho[i][0]).trim().toUpperCase() === 'DATA') { col = colunasDoCabecalho(trecho[i]); break; }
  }
  if (!col || col.pres < 0 || col.nome < 0) return { ok: false, erro: 'Não achei o cabeçalho deste dia na aba ' + p.loja + '.' };

  // confere se a linha ainda é a mesma pessoa no mesmo dia
  const igual = function (a, b) { return String(a).trim().toUpperCase() === String(b).trim().toUpperCase(); };
  if (!igual(atual[0], p.data) || !igual(atual[col.nome], p.nome)) {
    return { ok: false, mudou: true, erro: 'Essa linha mudou na planilha. Atualizei a lista, confira e tente de novo.' };
  }

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
  CacheService.getScriptCache().remove('presenca');   // a próxima leitura já vem com a mudança
  return { ok: true, presenca: p.acao === 'chegou' ? 'OK' : 'N/D' };
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ======================= TESTES (rodar no editor) ======================= */

/** Clique em Executar com esta função e veja o resultado em "Registro de execução". */
function testar() {
  const t = Date.now();
  const csv = montarCsv();
  Logger.log(csv.split('\n').length - 1 + ' linhas em ' + Math.round((Date.now() - t) / 1000) + ' segundos');
  Logger.log('Tamanho: ' + Utilities.newBlob(csv).getBytes().length + ' bytes');
  Logger.log(csv.split('\n').slice(0, 6).join('\n'));
}

/** Confere se a senha foi criada (não mostra a senha). */
function testarSenha() {
  const s = PropertiesService.getScriptProperties().getProperty('SENHA_EDICAO');
  Logger.log(s ? 'Senha criada, com ' + s.length + ' caracteres.' : 'Senha NÃO criada ainda.');
}
