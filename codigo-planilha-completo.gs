/**
 * LogMax · CÓDIGO COMPLETO DA PLANILHA (painel público + tela do operador) — versão 4
 *
 * Cole TUDO num único arquivo .gs do Apps Script da planilha (apague o que estava nele antes).
 * Depois: Implantar → Gerenciar implantações → lápis → Versão: Nova versão → Implantar.
 * (NUNCA "Nova implantação": isso muda o endereço /exec e as telas param.)
 *
 * O QUE MUDOU NA VERSÃO 4 (para não travar):
 *   - A escala é montada UMA vez e guardada pronta (em pedaços, para caber na memória do Google).
 *     Todas as telas (painel e operador) usam a mesma escala guardada.
 *   - Só UMA montagem por vez. Se outra tela pedir enquanto monta, recebe a última escala pronta na hora.
 *   - Se a montagem passar de 4 min e meio, entrega o que já leu (não estoura o limite de 6 min).
 *   - Ao marcar Chegou/Desfazer, só a linha daquela pessoa é atualizada na escala guardada (rápido).
 *
 * O QUE ENTREGA: loja, data, horário, código, nome, presença, nº da linha e hora da chegada.
 * NUNCA entrega telefone, CPF, PIX nem valores.
 *
 * FUNÇÕES PARA RODAR NO EDITOR (botão Executar):
 *   testarTudo           → monta a escala e mostra quantas linhas, quantos segundos e quais abas são lentas.
 *   diagnosticarLentidao → mede o tempo de cada aba, uma por uma (sem dados pessoais).
 */
const LM_SENHA = 'COLOQUE_A_SENHA_AQUI';               // senha dos operadores (para trocar: mude aqui e faça "Nova versão")
const LM_ABAS_IGNORADAS = ['RESUMO', 'FOLHA DE PAGAMENTO', 'CADASTRO MOTORISTA', 'ENDEREÇOS', 'PAG FNAL'];
const LM_DIAS_ANTES = 1, LM_DIAS_DEPOIS = 8;  // de ontem até daqui a 8 dias
const LM_LINHAS_FINAIS = 700;                 // lê só o final de cada aba (onde ficam os dias recentes)
const LM_COLUNAS = 8;                         // DATA, COD, NOME, TELEFONE, HORARIO, PRESENÇA... (até a coluna H)
const LM_FRESCO_SEG = 25;                     // até 25 s a escala guardada vale sem montar de novo
const LM_GUARDA_SEG = 21600;                  // guarda a última escala pronta por até 6 horas
const LM_LIMITE_MONTAR_MS = 270000;           // 4,5 minutos
const LM_MAX_ERROS_SENHA = 15;                // 15 senhas erradas em 10 min → bloqueia 10 min
const LM_ONLINE_SEG = 90;                     // tela que não avisa há 90 s sai do "Online agora"
const LM_CHAVE = 'escala_v4';
const LM_PEDACO = 40000;                      // tamanho de cada pedaço guardado
const LM_PREFIXO_CHEGADA = 'ch|';

/* ============================== PAINEL (leitura pública) ============================== */

function doGet(e) {
  lmOnline('p', e && e.parameter ? e.parameter.v : '');
  return ContentService.createTextOutput(lmEscala()).setMimeType(ContentService.MimeType.CSV);
}

/* ============================== OPERADOR (com senha) ============================== */

function doPost(e) {
  let pedido;
  try { pedido = JSON.parse(e.postData.contents); } catch (err) { return lmResposta({ ok: false, erro: 'Pedido inválido.' }); }

  const cache = CacheService.getScriptCache();
  const erros = +(cache.get('lm_erros_senha') || 0);
  if (erros >= LM_MAX_ERROS_SENHA) return lmResposta({ ok: false, erro: 'Muitas senhas erradas. Espere 10 minutos.', senha: true });
  if (String(pedido.senha || '') !== LM_SENHA) {
    cache.put('lm_erros_senha', String(erros + 1), 600);
    return lmResposta({ ok: false, erro: 'Senha errada.', senha: true });
  }

  if (pedido.acao === 'entrar') return lmResposta({ ok: true, planilha: SpreadsheetApp.getActiveSpreadsheet().getName() });
  if (pedido.acao === 'ler') return lmResposta({ ok: true, csv: lmEscala(), online: lmOnline('o', pedido.sessao) });
  if (pedido.acao !== 'chegou' && pedido.acao !== 'desfazer') return lmResposta({ ok: false, erro: 'Ação desconhecida.' });

  const trava = LockService.getScriptLock();   // uma gravação por vez (separada da montagem da escala)
  if (!trava.tryLock(15000)) return lmResposta({ ok: false, erro: 'A planilha está ocupada. Tente de novo.' });
  try {
    return lmResposta(lmMarcar(pedido));
  } finally {
    trava.releaseLock();
  }
}

function lmResposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ============================== ESCALA GUARDADA ============================== */

// Devolve a escala (CSV). Usa a guardada se estiver fresca; senão, UMA tela monta e as outras recebem a última pronta.
function lmEscala() {
  const guardada = lmLerGuardada();
  if (guardada && Date.now() - guardada.t < LM_FRESCO_SEG * 1000) return guardada.csv;

  const trava = LockService.getDocumentLock() || LockService.getUserLock();
  if (trava.tryLock(guardada ? 200 : 28000)) {
    try {
      const outra = lmLerGuardada();   // outra tela pode ter acabado de montar
      if (outra && Date.now() - outra.t < LM_FRESCO_SEG * 1000) return outra.csv;
      const csv = lmMontar().csv;
      lmGuardar(csv, Date.now());
      return csv;
    } finally {
      trava.releaseLock();
    }
  }
  if (guardada) return guardada.csv;   // alguém está montando agora: entrega a última pronta, na hora
  const depois = lmLerGuardada();
  return depois ? depois.csv : lmCsv([]);
}

function lmGuardar(csv, t) {
  const partes = [];
  for (let i = 0; i < csv.length; i += LM_PEDACO) partes.push(csv.slice(i, i + LM_PEDACO));
  const obj = {};
  partes.forEach(function (p, i) { obj[LM_CHAVE + '_' + i] = p; });
  obj[LM_CHAVE] = JSON.stringify({ n: partes.length, t: t });
  try { CacheService.getScriptCache().putAll(obj, LM_GUARDA_SEG); } catch (e) { console.warn('Não consegui guardar a escala: ' + e); }
}

function lmLerGuardada() {
  const cache = CacheService.getScriptCache();
  const m = cache.get(LM_CHAVE);
  if (!m) return null;
  let meta;
  try { meta = JSON.parse(m); } catch (e) { return null; }
  const chaves = [];
  for (let i = 0; i < meta.n; i++) chaves.push(LM_CHAVE + '_' + i);
  const v = cache.getAll(chaves);
  let csv = '';
  for (let i = 0; i < chaves.length; i++) { if (v[chaves[i]] == null) return null; csv += v[chaves[i]]; }
  return { csv: csv, t: meta.t };
}

// Depois do Chegou/Desfazer: muda só a linha daquela pessoa na escala guardada.
function lmAtualizarGuardada(loja, linha, presenca, chegada) {
  const g = lmLerGuardada();
  if (!g) return;
  const linhas = g.csv.split('\n');
  for (let i = 1; i < linhas.length; i++) {
    const c = lmLerLinhaCsv(linhas[i]);
    if (c[0] === loja && c[6] === String(linha)) {
      c[5] = presenca; c[8] = chegada || '';
      linhas[i] = lmLinhaCsv(c);
      lmGuardar(linhas.join('\n'), g.t);
      return;
    }
  }
}

/* ============================== MONTAR A ESCALA ============================== */

function lmColunas(l) {
  const achar = function (txt) { return l.findIndex(function (c) { return String(c).toUpperCase().indexOf(txt) >= 0; }); };
  const col = { cod: achar('COD'), nome: achar('NOME'), hora: achar('HORARIO'), pres: achar('PRESEN') };
  if (col.pres < 0 && col.hora >= 0) col.pres = col.hora + 1;   // abas sem título "PRESENÇA" (ex.: PAMPLONA)
  return col;
}

function lmAbaDeLoja(aba) {
  const nome = aba.getName().trim();
  if (LM_ABAS_IGNORADAS.indexOf(nome.toUpperCase()) >= 0) return false;
  if (/^ontem/i.test(nome) || aba.isSheetHidden()) return false;
  return true;
}

function lmMontar() {
  const comeco = Date.now();
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const hoje = new Date();
  const min = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - LM_DIAS_ANTES);
  const max = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() + LM_DIAS_DEPOIS);
  const chegadas = lmLerChegadas();
  const saida = [];
  const lentas = [], puladas = [];

  planilha.getSheets().forEach(function (aba) {
    if (!lmAbaDeLoja(aba)) return;
    const loja = aba.getName().trim();
    if (Date.now() - comeco > LM_LIMITE_MONTAR_MS) { puladas.push(loja); return; }
    const t0 = Date.now();
    const ultima = aba.getLastRow();
    if (ultima < 2) return;
    const inicio = Math.max(1, ultima - LM_LINHAS_FINAIS + 1);
    const linhas = aba.getRange(inicio, 1, ultima - inicio + 1, Math.min(LM_COLUNAS, aba.getMaxColumns())).getDisplayValues();
    let col = null;
    linhas.forEach(function (l, i) {
      if (String(l[0]).trim().toUpperCase() === 'DATA') { col = lmColunas(l); return; }
      const m = String(l[0]).trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
      if (!m || !col || col.pres < 0) return;
      const d = new Date(+m[3], +m[2] - 1, +m[1]);
      if (d < min || d > max) return;
      const pega = function (k) { return k >= 0 ? String(l[k]).replace(/\s*[\r\n]+\s*/g, ' ').trim() : ''; };
      const data = String(l[0]).trim(), nome = pega(col.nome);
      saida.push([loja, data, pega(col.hora), pega(col.cod), nome, pega(col.pres), String(inicio + i), '',
                  lmChegadaTexto(chegadas, loja, data, inicio + i, nome)]);
    });
    const seg = (Date.now() - t0) / 1000;
    if (seg > 5) lentas.push(loja + ' (' + seg.toFixed(0) + ' s)');
  });

  if (lentas.length) console.warn('Abas lentas: ' + lentas.join(', '));
  if (puladas.length) console.warn('Passou de 4,5 min. Abas que ficaram de fora desta vez: ' + puladas.join(', '));
  return { csv: lmCsv(saida), segundos: (Date.now() - comeco) / 1000, lentas: lentas, puladas: puladas, linhas: saida.length };
}

function lmCsv(linhas) {
  const agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
  const saida = [['LOJA', 'DATA', 'HORARIO', 'CODIGO', 'NOME', 'PRESENCA', 'LINHA', 'ATUALIZADO', 'CHEGADA']].concat(linhas);
  if (saida.length > 1) saida[1][7] = agora;
  else saida.push(['', '', '', '', '', '', '', agora, '']);
  return saida.map(lmLinhaCsv).join('\n');
}

function lmLinhaCsv(linha) {
  return linha.map(function (c) { c = String(c == null ? '' : c); return /[",\n]/.test(c) ? '"' + c.replace(/"/g, '""') + '"' : c; }).join(',');
}

function lmLerLinhaCsv(t) {
  const out = []; let campo = '', aspas = false;
  for (let i = 0; i < t.length; i++) {
    const c = t[i];
    if (aspas) { if (c === '"' && t[i + 1] === '"') { campo += '"'; i++; } else if (c === '"') aspas = false; else campo += c; }
    else if (c === '"') aspas = true;
    else if (c === ',') { out.push(campo); campo = ''; }
    else campo += c;
  }
  out.push(campo);
  while (out.length < 9) out.push('');
  return out;
}

/* ============================== CHEGOU / DESFAZER ============================== */

function lmLocalizar(p) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(String(p.loja || ''));
  if (!aba || !lmAbaDeLoja(aba)) return { erro: { ok: false, erro: 'Não achei a aba da loja ' + p.loja + '.' } };
  const linha = parseInt(p.linha, 10);
  if (!(linha >= 2) || linha > aba.getLastRow()) return { erro: { ok: false, erro: 'Linha fora da planilha.', mudou: true } };

  const de = Math.max(1, linha - 300);   // acha o cabeçalho do bloco desta linha (subindo até 300 linhas)
  const trecho = aba.getRange(de, 1, linha - de + 1, Math.min(LM_COLUNAS, aba.getMaxColumns())).getDisplayValues();
  const atual = trecho[trecho.length - 1];
  let col = null;
  for (let i = trecho.length - 2; i >= 0; i--) {
    if (String(trecho[i][0]).trim().toUpperCase() === 'DATA') { col = lmColunas(trecho[i]); break; }
  }
  if (!col || col.pres < 0 || col.nome < 0) return { erro: { ok: false, erro: 'Não achei o cabeçalho deste dia na aba ' + p.loja + '.' } };

  const igual = function (a, b) { return String(a).replace(/\s*[\r\n]+\s*/g, ' ').trim().toUpperCase() === String(b).trim().toUpperCase(); };
  if (!igual(atual[0], p.data) || !igual(atual[col.nome], p.nome)) {
    return { erro: { ok: false, mudou: true, erro: 'Essa linha mudou na planilha. Atualizei a lista, confira e tente de novo.' } };
  }
  return { aba: aba, linha: linha, atual: atual, col: col };
}

function lmMarcar(p) {
  const achou = lmLocalizar(p);
  if (achou.erro) return achou.erro;
  const aba = achou.aba, linha = achou.linha, atual = achou.atual, col = achou.col;
  const celula = aba.getRange(linha, col.pres + 1);
  const agora = String(atual[col.pres]).trim().toUpperCase();
  const dataLinha = String(atual[0]).trim(), nomeLinha = String(atual[col.nome]).replace(/\s*[\r\n]+\s*/g, ' ').trim();
  let chegada = '';

  if (p.acao === 'chegou') {
    if (agora === 'OK') return { ok: true, presenca: 'OK', jaEstava: true, chegada: lmChegadaTexto(lmLerChegadas(), p.loja, dataLinha, linha, nomeLinha) };
    if (agora !== '' && agora !== 'N/D') {
      return { ok: false, erro: 'Na planilha está escrito "' + atual[col.pres] + '" para essa pessoa. Mude direto na planilha.' };
    }
    celula.setValue('OK');
    chegada = lmGuardarChegada(p.loja, dataLinha, linha, nomeLinha);
  } else {
    lmApagarChegada(p.loja, dataLinha, linha);
    if (agora !== 'OK') return { ok: true, presenca: atual[col.pres], jaEstava: true };
    celula.setValue('N/D');
  }
  SpreadsheetApp.flush();
  const presenca = p.acao === 'chegou' ? 'OK' : 'N/D';
  lmAtualizarGuardada(p.loja, linha, presenca, chegada);
  return { ok: true, presenca: presenca, chegada: chegada };
}

/* ============================== HORA DA CHEGADA (guardada no Apps Script, não na planilha) ============================== */

function lmChaveChegada(loja, data, linha) { return LM_PREFIXO_CHEGADA + loja + '|' + data + '|' + linha; }
function lmLerChegadas() { try { return PropertiesService.getScriptProperties().getProperties(); } catch (e) { return {}; } }
function lmChegadaTexto(todas, loja, data, linha, nome) {   // "HH:mm|minutos desde 00:00 do dia da escala"
  const v = todas[lmChaveChegada(loja, data, linha)];
  if (!v) return '';
  try {
    const o = JSON.parse(v);
    if (String(o.n).trim().toUpperCase() !== String(nome).trim().toUpperCase()) return '';
    return o.h + '|' + o.m;
  } catch (e) { return ''; }
}
function lmGuardarChegada(loja, data, linha, nome) {
  const a = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd HH:mm').match(/(\d+)-(\d+)-(\d+) (\d+):(\d+)/);
  const e = String(data).match(/(\d{2})\/(\d{2})\/(\d{4})/);
  if (!a || !e) return '';
  const dias = Math.round((Date.UTC(+a[1], +a[2] - 1, +a[3]) - Date.UTC(+e[3], +e[2] - 1, +e[1])) / 864e5);
  const min = dias * 1440 + (+a[4]) * 60 + (+a[5]);
  const hora = a[4] + ':' + a[5];
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(lmChaveChegada(loja, data, linha), JSON.stringify({ n: String(nome).trim(), h: hora, m: min }));
    lmLimparChegadasVelhas(props);
  } catch (err) { return ''; }
  return hora + '|' + min;
}
function lmApagarChegada(loja, data, linha) {
  try { PropertiesService.getScriptProperties().deleteProperty(lmChaveChegada(loja, data, linha)); } catch (e) {}
}
function lmLimparChegadasVelhas(props) {   // apaga horas anotadas de mais de 10 dias atrás
  const limite = Date.now() - 10 * 864e5;
  props.getKeys().forEach(function (k) {
    if (k.indexOf(LM_PREFIXO_CHEGADA) !== 0) return;
    const m = k.match(/\|(\d{2})\/(\d{2})\/(\d{4})\|/);
    if (m && Date.UTC(+m[3], +m[2] - 1, +m[1]) < limite) props.deleteProperty(k);
  });
}

/* ============================== QUEM ESTÁ COM A TELA ABERTA ============================== */

function lmOnline(tipo, id) {   // 'o' = operador, 'p' = painel
  const cache = CacheService.getScriptCache(), agora = Date.now();
  const ler = function () { try { return JSON.parse(cache.get('lm_online') || '{}'); } catch (e) { return {}; } };
  let mapa = ler();
  if (id && /^[A-Za-z0-9]{6,40}$/.test(String(id))) {
    try {
      mapa[id] = [tipo, agora];
      Object.keys(mapa).forEach(function (k) { if (agora - mapa[k][1] > LM_ONLINE_SEG * 1000) delete mapa[k]; });
      cache.put('lm_online', JSON.stringify(mapa), 600);
    } catch (e) { /* se falhar, só não conta desta vez */ }
  }
  const conta = { o: 0, p: 0 };
  Object.keys(mapa).forEach(function (k) {
    if (agora - mapa[k][1] <= LM_ONLINE_SEG * 1000) conta[mapa[k][0]] = (conta[mapa[k][0]] || 0) + 1;
  });
  return conta;
}

/* ============================== TESTES (rodar no editor) ============================== */

/** Monta a escala do zero e mostra quantas linhas, quanto tempo e quais abas estão lentas. */
function testarTudo() {
  const r = lmMontar();
  lmGuardar(r.csv, Date.now());
  console.log('Escala montada: ' + r.linhas + ' linhas em ' + r.segundos.toFixed(1) + ' segundos (' +
    Math.round(Utilities.newBlob(r.csv).getBytes().length / 1024) + ' KB).');
  console.log(r.lentas.length ? 'Abas lentas (mais de 5 s): ' + r.lentas.join(', ') : 'Nenhuma aba lenta.');
  if (r.puladas.length) console.log('Ficaram de fora por tempo: ' + r.puladas.join(', '));
  console.log('Guardada ok? ' + (lmLerGuardada() ? 'SIM' : 'NÃO'));
  console.log('Agora faça: Implantar → Gerenciar implantações → lápis → Nova versão → Implantar.');
}

/** Mede o tempo de cada aba, uma por uma (mostra só nome e tamanho das abas, nenhum dado de motoboy). */
function diagnosticarLentidao() {
  const inicioGeral = Date.now();
  const abas = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  console.log('Abas na planilha: ' + abas.length);
  for (let i = 0; i < abas.length; i++) {
    if (Date.now() - inicioGeral > 300000) { console.log('Parei aos 5 minutos para não estourar o tempo.'); break; }
    const aba = abas[i], nome = aba.getName();
    const t0 = Date.now();
    const oculta = aba.isSheetHidden();
    const ultima = aba.getLastRow(), maxLinhas = aba.getMaxRows(), maxColunas = aba.getMaxColumns();
    const t1 = Date.now();
    let leitura = 'não lida';
    if (lmAbaDeLoja(aba) && ultima >= 2) {
      const ini = Math.max(1, ultima - LM_LINHAS_FINAIS + 1);
      aba.getRange(ini, 1, ultima - ini + 1, Math.min(LM_COLUNAS, maxColunas)).getDisplayValues();
      leitura = ((Date.now() - t1) / 1000).toFixed(1) + ' s';
    }
    console.log((i + 1) + '. ' + nome + (oculta ? ' (oculta)' : '') + ' | até a linha ' + ultima +
      ' (aba com ' + maxLinhas + ' linhas e ' + maxColunas + ' colunas) | tamanho: ' + ((t1 - t0) / 1000).toFixed(1) +
      ' s | leitura: ' + leitura);
  }
  console.log('TOTAL: ' + ((Date.now() - inicioGeral) / 1000).toFixed(1) + ' segundos');
}
