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
 * NA VERSÃO 4.1: tela do operador pode COLOCAR motoboy na vaga ou TROCAR quem está, digitando só o código.
 *   O código vai na coluna COD; o NOME vem sozinho pela fórmula da planilha. O OK continua separado.
 *
 * NA VERSÃO 4.2: a tela do operador (com senha) recebe também o TELEFONE, para o botão do WhatsApp.
 *   O painel público NÃO recebe telefone.
 *
 * NA VERSÃO 4.3: aba "Localizar motoboy" na tela do operador (com senha). A busca é feita AQUI na planilha,
 *   na aba de cadastro dos motoboys, e devolve só quem bate com o que foi digitado (no máximo 20):
 *   código, nome, telefone, CPF e PIX. A lista inteira nunca sai da planilha. O painel público não tem isso.
 *   (4.4: escolha do campo de busca — código, nome, telefone ou CPF — e o PIX no resultado.)
 *
 * NA VERSÃO 4.5: "Pagamento no Final" (PGF). O botão PGF da tela do operador grava o valor combinado numa aba
 *   NOVA da planilha chamada "PGF SITE" (criada sozinha; não mexe na PAG FNAL nem em nenhuma outra aba).
 *   A aba Pagamento no Final do site lê essa lista e tem o botão Pago.
 *
 * NA VERSÃO 4.6: "Novo Cadastro" na tela do operador. Escreve UMA linha nova na aba de cadastro, colunas A até H,
 *   na mesma ordem da planilha (A = carimbo de data/hora, preenchido sozinho). O COD FUN é sugerido (maior número + 1).
 *   CPF que já existe é bloqueado e a resposta traz o código já cadastrado.
 *
 * NA VERSÃO 4.7: "Relatar Problema". Ocorrências (leve, média, grave) ficam numa aba NOVA "PROBLEMAS SITE".
 *   A tela avisa (vermelho piscando + som) quando alguém com problema GRAVE é escalado.
 *
 * O QUE ENTREGA: loja, data, horário, código, nome, presença, nº da linha e hora da chegada.
 * Telefone: só para a tela do operador (com senha). NUNCA entrega CPF, PIX nem valores.
 *
 * FUNÇÕES PARA RODAR NO EDITOR (botão Executar):
 *   testarTudo           → monta a escala e mostra quantas linhas, quantos segundos e quais abas são lentas.
 *   diagnosticarLentidao → mede o tempo de cada aba, uma por uma (sem dados pessoais).
 *   conferirNomePorCodigo → diz, em cada loja, se a coluna NOME puxa o nome pelo código (fórmula).
 *   liberarSenha         → libera na hora a trava de "Muitas senhas erradas".
 *   conferirCadastro     → mostra quais colunas do cadastro a busca usa (só os títulos, sem dados pessoais).
 */
const LM_SENHA = 'COLOQUE_A_SENHA_AQUI';               // senha dos operadores (para trocar: mude aqui e faça "Nova versão")
const LM_ABAS_IGNORADAS = ['RESUMO', 'FOLHA DE PAGAMENTO', 'CADASTRO MOTORISTA', 'ENDEREÇOS', 'PAG FNAL', 'PGF SITE', 'PROBLEMAS SITE'];
const LM_DIAS_ANTES = 1, LM_DIAS_DEPOIS = 8;  // de ontem até daqui a 8 dias
const LM_LINHAS_FINAIS = 700;                 // lê só o final de cada aba (onde ficam os dias recentes)
const LM_COLUNAS = 8;                         // DATA, COD, NOME, TELEFONE, HORARIO, PRESENÇA... (até a coluna H)
const LM_FRESCO_SEG = 25;                     // até 25 s a escala guardada vale sem montar de novo
const LM_GUARDA_SEG = 21600;                  // guarda a última escala pronta por até 6 horas
const LM_LIMITE_MONTAR_MS = 270000;           // 4,5 minutos
const LM_MAX_ERROS_SENHA = 5;                 // 5 senhas erradas → aquele aparelho espera
const LM_TRAVA_SEG = 60;                      // ...no máximo 1 minuto
const LM_ONLINE_SEG = 90;                     // tela que não avisa há 90 s sai do "Online agora"
const LM_CHAVE = 'escala_v4';
const LM_CHAVE_TEL = 'tel_v4';                 // telefones guardados à parte (só vão para o operador)
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

  // trava de senha: por APARELHO. 5 erros → aquele aparelho espera 1 minuto (os outros continuam entrando).
  const cache = CacheService.getScriptCache();
  const aparelho = /^[A-Za-z0-9]{6,40}$/.test(String(pedido.sessao || '')) ? String(pedido.sessao) : 'sem-id';
  const agoraMs = Date.now();
  const ateAp = +(cache.get('lm_trava_' + aparelho) || 0), ateGeral = +(cache.get('lm_trava_geral') || 0);
  const ate = Math.max(ateAp, ateGeral);
  if (ate > agoraMs) {
    const espera = Math.ceil((ate - agoraMs) / 1000);
    return lmResposta({ ok: false, erro: 'Muitas senhas erradas. Tente de novo em ' + espera + ' s.', senha: true, espera: espera });
  }
  if (String(pedido.senha || '') !== LM_SENHA) {
    const errosAp = +(cache.get('lm_erros_' + aparelho) || 0) + 1, errosGeral = +(cache.get('lm_erros_geral') || 0) + 1;
    if (errosGeral >= 100) { cache.put('lm_trava_geral', String(agoraMs + LM_TRAVA_SEG * 1000), LM_TRAVA_SEG); cache.remove('lm_erros_geral'); }
    else cache.put('lm_erros_geral', String(errosGeral), 600);
    if (errosAp >= LM_MAX_ERROS_SENHA) {
      cache.put('lm_trava_' + aparelho, String(agoraMs + LM_TRAVA_SEG * 1000), LM_TRAVA_SEG);
      cache.remove('lm_erros_' + aparelho);
      return lmResposta({ ok: false, erro: 'Muitas senhas erradas. Tente de novo em ' + LM_TRAVA_SEG + ' s.', senha: true, espera: LM_TRAVA_SEG });
    }
    cache.put('lm_erros_' + aparelho, String(errosAp), 600);
    const resta = LM_MAX_ERROS_SENHA - errosAp;
    return lmResposta({ ok: false, erro: 'Senha errada.' + (resta <= 2 ? (resta > 1 ? ' Restam ' + resta + ' tentativas' : ' Resta 1 tentativa') + ' antes de esperar 1 minuto.' : ''), senha: true });
  }
  cache.remove('lm_erros_' + aparelho);

  if (pedido.acao === 'entrar') return lmResposta({ ok: true, planilha: SpreadsheetApp.getActiveSpreadsheet().getName() });
  if (pedido.acao === 'ler') {
    const csv = lmEscala();   // primeiro a escala (se precisar, ela também guarda os telefones)
    return lmResposta({ ok: true, csv: csv, tel: lmLerTelefones(), pgf: lmPgfLista(), problemas: lmProblemas(), online: lmOnline('o', pedido.sessao) });
  }
  if (pedido.acao === 'buscar') return lmResposta(lmBuscar(pedido.q, pedido.campo));
  if (pedido.acao === 'cadastroInfo') return lmResposta(lmCadastroInfo());
  if (pedido.acao === 'problemaNovo' || pedido.acao === 'problemaEncerrar') {
    const travaP = LockService.getScriptLock();
    if (!travaP.tryLock(15000)) return lmResposta({ ok: false, erro: 'A planilha está ocupada. Tente de novo.' });
    try { return lmResposta(pedido.acao === 'problemaNovo' ? lmProblemaNovo(pedido) : lmProblemaEncerrar(pedido.id, pedido.motivo)); }
    finally { travaP.releaseLock(); }
  }
  if (pedido.acao === 'cadastrar') {
    const travaCad = LockService.getScriptLock();
    if (!travaCad.tryLock(15000)) return lmResposta({ ok: false, erro: 'A planilha está ocupada. Tente de novo.' });
    try { return lmResposta(lmCadastrar(pedido)); } finally { travaCad.releaseLock(); }
  }
  if (['pgf', 'pgfPago', 'pgfRemover'].indexOf(pedido.acao) >= 0) {
    const travaPgf = LockService.getScriptLock();
    if (!travaPgf.tryLock(15000)) return lmResposta({ ok: false, erro: 'A planilha está ocupada. Tente de novo.' });
    try {
      if (pedido.acao === 'pgf') return lmResposta(lmPgfMarcar(pedido));
      if (pedido.acao === 'pgfPago') return lmResposta(lmPgfPago(pedido.id, !!pedido.pago));
      return lmResposta(lmPgfRemover(pedido.id));
    } finally { travaPgf.releaseLock(); }
  }
  if (['chegou', 'desfazer', 'codigo'].indexOf(pedido.acao) < 0) return lmResposta({ ok: false, erro: 'Ação desconhecida.' });

  const trava = LockService.getScriptLock();   // uma gravação por vez (separada da montagem da escala)
  if (!trava.tryLock(15000)) return lmResposta({ ok: false, erro: 'A planilha está ocupada. Tente de novo.' });
  try {
    return lmResposta(pedido.acao === 'codigo' ? lmTrocarCodigo(pedido) : lmMarcar(pedido));
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
      const r = lmMontar();
      lmGuardarTexto(LM_CHAVE_TEL, JSON.stringify(r.tel), Date.now());
      lmGuardar(r.csv, Date.now());
      const csv = r.csv;
      return csv;
    } finally {
      trava.releaseLock();
    }
  }
  if (guardada) return guardada.csv;   // alguém está montando agora: entrega a última pronta, na hora
  const depois = lmLerGuardada();
  return depois ? depois.csv : lmCsv([]);
}

// Guarda um texto grande em pedaços (a memória do Google aceita no máximo ~100 KB por item).
function lmGuardarTexto(chave, texto, t) {
  const partes = [];
  for (let i = 0; i < texto.length; i += LM_PEDACO) partes.push(texto.slice(i, i + LM_PEDACO));
  if (!partes.length) partes.push('');
  const obj = {};
  partes.forEach(function (p, i) { obj[chave + '_' + i] = p; });
  obj[chave] = JSON.stringify({ n: partes.length, t: t });
  try { CacheService.getScriptCache().putAll(obj, LM_GUARDA_SEG); } catch (e) { console.warn('Não consegui guardar ' + chave + ': ' + e); }
}
function lmLerTexto(chave) {
  const cache = CacheService.getScriptCache();
  const m = cache.get(chave);
  if (!m) return null;
  let meta;
  try { meta = JSON.parse(m); } catch (e) { return null; }
  const chaves = [];
  for (let i = 0; i < meta.n; i++) chaves.push(chave + '_' + i);
  const v = cache.getAll(chaves);
  let texto = '';
  for (let i = 0; i < chaves.length; i++) { if (v[chaves[i]] == null) return null; texto += v[chaves[i]]; }
  return { texto: texto, t: meta.t };
}
function lmGuardar(csv, t) { lmGuardarTexto(LM_CHAVE, csv, t); }
function lmLerGuardada() { const g = lmLerTexto(LM_CHAVE); return g ? { csv: g.texto, t: g.t } : null; }
function lmLerTelefones() {
  const g = lmLerTexto(LM_CHAVE_TEL);
  if (!g) return {};
  try { return JSON.parse(g.texto || '{}'); } catch (e) { return {}; }
}
function lmMudarTelefone(loja, linha, tel) {
  const mapa = lmLerTelefones();
  if (tel) mapa[loja + '|' + linha] = tel; else delete mapa[loja + '|' + linha];
  lmGuardarTexto(LM_CHAVE_TEL, JSON.stringify(mapa), Date.now());
}

// Depois do Chegou/Desfazer/troca: muda só a linha daquela pessoa na escala guardada.
// mudancas = { coluna: valor }  (3 = CODIGO, 4 = NOME, 5 = PRESENCA, 8 = CHEGADA)
function lmMudarGuardada(loja, linha, mudancas) {
  const g = lmLerGuardada();
  if (!g) return;
  const linhas = g.csv.split('\n');
  for (let i = 1; i < linhas.length; i++) {
    const c = lmLerLinhaCsv(linhas[i]);
    if (c[0] === loja && c[6] === String(linha)) {
      Object.keys(mudancas).forEach(function (k) { c[+k] = mudancas[k] == null ? '' : String(mudancas[k]); });
      linhas[i] = lmLinhaCsv(c);
      lmGuardar(linhas.join('\n'), g.t);
      return;
    }
  }
}
function lmAtualizarGuardada(loja, linha, presenca, chegada) { lmMudarGuardada(loja, linha, { 5: presenca, 8: chegada || '' }); }

/* ============================== MONTAR A ESCALA ============================== */

function lmColunas(l) {
  const achar = function (txt) { return l.findIndex(function (c) { return String(c).toUpperCase().indexOf(txt) >= 0; }); };
  const col = { cod: achar('COD'), nome: achar('NOME'), hora: achar('HORARIO'), pres: achar('PRESEN') };
  col.tel = achar('TELEF'); if (col.tel < 0) col.tel = achar('CELULAR'); if (col.tel < 0) col.tel = achar('FONE');
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
  const tel = {};   // "LOJA|linha" → telefone só com números (só vai para a tela do operador)
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
      const fone = lmLimparTelefone(pega(col.tel));
      if (fone) tel[loja + '|' + (inicio + i)] = fone;
      saida.push([loja, data, pega(col.hora), pega(col.cod), nome, pega(col.pres), String(inicio + i), '',
                  lmChegadaTexto(chegadas, loja, data, inicio + i, nome)]);
    });
    const seg = (Date.now() - t0) / 1000;
    if (seg > 5) lentas.push(loja + ' (' + seg.toFixed(0) + ' s)');
  });

  if (lentas.length) console.warn('Abas lentas: ' + lentas.join(', '));
  if (puladas.length) console.warn('Passou de 4,5 min. Abas que ficaram de fora desta vez: ' + puladas.join(', '));
  return { csv: lmCsv(saida), tel: tel, segundos: (Date.now() - comeco) / 1000, lentas: lentas, puladas: puladas, linhas: saida.length };
}

// "(11) 98765-4321" → "5511987654321"; devolve "" se não parecer um telefone do Brasil
function lmLimparTelefone(t) {
  let d = String(t || '').replace(/\D/g, '');
  if (d.length === 10 || d.length === 11) d = '55' + d;
  return (d.length === 12 || d.length === 13) && d.indexOf('55') === 0 ? d : '';
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

/* ============================== COLOCAR / TROCAR MOTOBOY PELO CÓDIGO ============================== */

// Escreve o código na coluna COD; o NOME vem sozinho pela fórmula da planilha. Não mexe no OK.
function lmTrocarCodigo(p) {
  const achou = lmLocalizar(p);
  if (achou.erro) return achou.erro;
  const aba = achou.aba, linha = achou.linha, atual = achou.atual, col = achou.col;
  if (col.cod < 0) return { ok: false, erro: 'Não achei a coluna COD nesta aba.' };

  const novo = String(p.codigo || '').trim();
  if (!/^[0-9A-Za-z-]{1,12}$/.test(novo)) return { ok: false, erro: 'Digite um código válido (só números ou letras).' };
  if (String(atual[col.cod]).trim().toUpperCase() === novo.toUpperCase()) {
    return { ok: true, nome: String(atual[col.nome]).trim(), codigo: novo, jaEstava: true };
  }
  // quem já está com OK não pode ser trocado (o OK passaria para outra pessoa no pagamento)
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

  const nome = String(celNome.getDisplayValue()).replace(/\s*[\r\n]+\s*/g, ' ').trim();
  if (!nome || /N[ÃA]O\s+LOCALIZAD|^#/i.test(nome)) {
    celCod.setValue(antes);   // código não existe no cadastro: volta como estava
    SpreadsheetApp.flush();
    return { ok: false, erro: 'Código ' + novo + ' não encontrado no cadastro. Nada foi mudado.' };
  }
  lmApagarChegada(p.loja, String(atual[0]).trim(), linha);   // hora de chegada era da pessoa anterior
  lmMudarGuardada(p.loja, linha, { 3: novo, 4: nome, 8: '' });
  const fone = col.tel >= 0 ? lmLimparTelefone(aba.getRange(linha, col.tel + 1).getDisplayValue()) : '';
  lmMudarTelefone(p.loja, linha, fone);
  return { ok: true, nome: nome, codigo: novo, tel: fone };
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

/* ============================== LOCALIZAR MOTOBOY (cadastro) ============================== */

const LM_CHAVE_CAD = 'cad_v2';
const LM_CAD_FRESCO_MIN = 10;   // o cadastro é relido da planilha a cada 10 minutos

function lmAbaCadastro() {
  const abas = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (let i = 0; i < abas.length; i++) if (/CADASTRO/i.test(abas[i].getName())) return abas[i];
  return null;
}

// Lê a aba de cadastro e acha as colunas pelo título. Guarda só na memória interna da planilha (não sai daqui).
function lmLerCadastro(forcar) {
  if (!forcar) {
    const g = lmLerTexto(LM_CHAVE_CAD);
    if (g && Date.now() - g.t < LM_CAD_FRESCO_MIN * 60000) { try { return JSON.parse(g.texto); } catch (e) {} }
  }
  const aba = lmAbaCadastro();
  if (!aba) return { erro: 'Não achei a aba de cadastro dos motoboys (uma aba com "CADASTRO" no nome).' };
  const ultima = aba.getLastRow(), ncol = Math.min(aba.getLastColumn(), 40);
  if (ultima < 2) return { erro: 'A aba de cadastro está vazia.' };
  const vals = aba.getRange(1, 1, ultima, ncol).getDisplayValues();
  let cab = -1;
  for (let i = 0; i < Math.min(15, vals.length); i++) {
    if (vals[i].some(function (c) { return /^\s*NOME/i.test(String(c)); })) { cab = i; break; }
  }
  if (cab < 0) return { erro: 'Não achei a linha de títulos (com "NOME") no cadastro.' };
  const t = vals[cab].map(function (c) { return String(c).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim(); });
  const achar = function (teste) { for (let j = 0; j < t.length; j++) if (teste(t[j])) return j; return -1; };
  const col = {
    cod: achar(function (x) { return /^(COD|CODIGO|ID|MATRICULA)\b/.test(x); }),
    nome: achar(function (x) { return /^NOME/.test(x); }),
    tel: achar(function (x) { return /TELEF|CELULAR|WHATS|FONE/.test(x); }),
    cpf: achar(function (x) { return /\bCPF\b/.test(x); }),
    pix: achar(function (x) { return /^(CHAVE\s*(DO\s*)?)?PIX$/.test(x); })
  };
  if (col.pix < 0) col.pix = achar(function (x) { return /PIX/.test(x) && !/TIPO/.test(x); });
  const lista = [];
  for (let i = cab + 1; i < vals.length; i++) {
    const l = vals[i], nome = col.nome >= 0 ? String(l[col.nome]).trim() : '';
    if (!nome) continue;
    lista.push([col.cod >= 0 ? String(l[col.cod]).trim() : '', nome,
                col.tel >= 0 ? String(l[col.tel]).trim() : '', col.cpf >= 0 ? String(l[col.cpf]).trim() : '',
                col.pix >= 0 ? String(l[col.pix]).trim() : '']);
  }
  const cad = { aba: aba.getName(), linhaTitulos: cab + 1, titulos: { cod: col.cod >= 0 ? vals[cab][col.cod] : '', nome: col.nome >= 0 ? vals[cab][col.nome] : '',
                tel: col.tel >= 0 ? vals[cab][col.tel] : '', cpf: col.cpf >= 0 ? vals[cab][col.cpf] : '',
                pix: col.pix >= 0 ? vals[cab][col.pix] : '' }, lista: lista };
  lmGuardarTexto(LM_CHAVE_CAD, JSON.stringify(cad), Date.now());
  return cad;
}

// Procura no campo escolhido (codigo | nome | telefone | cpf). Sem campo: tenta adivinhar. Devolve no máximo 20.
function lmBuscar(q, campo) {
  q = String(q || '').trim();
  if (q.length < 1) return { ok: false, erro: 'Digite o que procurar.' };
  const cad = lmLerCadastro(false);
  if (cad.erro) return { ok: false, erro: cad.erro };
  const semAcento = function (x) { return String(x).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, ''); };
  const digitos = q.replace(/\D/g, '');
  if (['codigo', 'nome', 'telefone', 'cpf'].indexOf(campo) < 0) {
    const soNumero = digitos.length >= 2 && digitos.length >= q.replace(/[\s().\-\/+]/g, '').length;
    campo = !soNumero ? 'nome' : digitos.length <= 6 ? 'codigo' : 'telefone';
  }
  if (campo === 'nome' && q.length < 2) return { ok: false, erro: 'Digite pelo menos 2 letras do nome.' };
  if ((campo === 'telefone' || campo === 'cpf') && digitos.length < 4) return { ok: false, erro: 'Digite pelo menos 4 números.' };
  if (campo === 'codigo' && !q) return { ok: false, erro: 'Digite o código.' };
  const palavras = semAcento(q).split(/\s+/).filter(Boolean);
  const achados = [];
  cad.lista.forEach(function (m) {
    const cod = m[0], nome = m[1], tel = m[2].replace(/\D/g, ''), cpf = m[3].replace(/\D/g, '');
    let nota = 0;
    if (campo === 'codigo') {
      if (semAcento(cod).trim() === semAcento(q).trim() || (digitos && cod.replace(/\D/g, '') === digitos)) nota = 3;
    } else if (campo === 'telefone') {
      if (tel.indexOf(digitos) >= 0 || (digitos.indexOf('55') === 0 && digitos.length > 11 && tel.indexOf(digitos.slice(2)) >= 0)) nota = 2;
    } else if (campo === 'cpf') {
      if (cpf.indexOf(digitos) >= 0) nota = cpf === digitos ? 3 : 2;
    } else {
      const n = semAcento(nome);
      if (palavras.every(function (p) { return n.indexOf(p) >= 0; })) nota = n.indexOf(palavras[0]) === 0 ? 2 : 1;
    }
    if (nota) achados.push({ nota: nota, cod: cod, nome: nome, tel: m[2], whats: lmLimparTelefone(m[2]), cpf: m[3], pix: m[4] || '' });
  });
  achados.sort(function (a, b) { return b.nota - a.nota || a.nome.localeCompare(b.nome); });
  return { ok: true, campo: campo, total: achados.length, itens: achados.slice(0, 20).map(function (a) { delete a.nota; return a; }) };
}

/* ============================== NOVO CADASTRO (aba de cadastro, colunas A até H) ============================== */

const LM_CAD_COLS = 8;   // A até H

function lmSemAcento(x) { return String(x).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim(); }

// Lê os títulos A–H e diz que tipo de campo é cada um
function lmCadastroEstrutura() {
  const aba = lmAbaCadastro();
  if (!aba) return { erro: 'Não achei a aba de cadastro dos motoboys.' };
  const ultima = aba.getLastRow();
  const titulos = aba.getRange(1, 1, 1, LM_CAD_COLS).getDisplayValues()[0].map(function (t) { return String(t).trim(); });
  const tipos = titulos.map(function (t, i) {
    const u = lmSemAcento(t);
    if (i === 0) return 'carimbo';
    if (/^(COD|CODIGO|ID|MATRICULA)\b/.test(u)) return 'codigo';
    if (/TIPO/.test(u) && /PIX|CHAVE/.test(u)) return 'tipoPix';
    if (/PIX/.test(u)) return 'pix';
    if (/\bCPF\b/.test(u)) return 'cpf';
    if (/TELEF|CELULAR|WHATS|FONE/.test(u)) return 'telefone';
    if (/NOME/.test(u)) return 'nome';
    if (/UNIDADE|LOJA/.test(u)) return 'unidade';
    if (/MAIL/.test(u)) return 'email';
    if (/NASC|DATA/.test(u)) return 'data';
    return 'texto';
  });
  return { aba: aba, ultima: ultima, titulos: titulos, tipos: tipos };
}

function lmCpfValido(c) {
  const d = String(c).replace(/\D/g, '');
  if (d.length !== 11 || /^(\d)\1{10}$/.test(d)) return false;
  for (let t = 9; t < 11; t++) {
    let soma = 0;
    for (let i = 0; i < t; i++) soma += +d[i] * (t + 1 - i);
    if (((soma * 10) % 11) % 10 !== +d[t]) return false;
  }
  return true;
}

// Para montar o formulário: títulos, tipos, opções das colunas com poucas respostas e o próximo código
function lmCadastroInfo() {
  const e = lmCadastroEstrutura();
  if (e.erro) return { ok: false, erro: e.erro };
  const linhas = e.ultima >= 2 ? e.aba.getRange(2, 1, e.ultima - 1, LM_CAD_COLS).getDisplayValues() : [];
  const iCod = e.tipos.indexOf('codigo');
  let maior = 0;
  if (iCod >= 0) linhas.forEach(function (l) { const n = parseInt(String(l[iCod]).replace(/\D/g, ''), 10); if (n > maior && n < 1e7) maior = n; });
  const opcoes = {};
  e.tipos.forEach(function (t, i) {
    if (t !== 'tipoPix' && t !== 'texto' && t !== 'unidade') return;
    const conta = {};
    linhas.slice(t === 'unidade' ? -1500 : -400).forEach(function (l) { const v = String(l[i]).trim(); if (v) conta[v] = (conta[v] || 0) + 1; });
    if (t === 'unidade') { opcoes[i] = Object.keys(conta).sort(function (a, b) { return a.localeCompare(b); }).slice(0, 60); return; }
    const lista = Object.keys(conta).sort(function (a, b) { return conta[b] - conta[a]; });
    if (t === 'tipoPix' || (lista.length > 0 && lista.length <= 8)) opcoes[i] = lista.slice(0, 10);
  });
  return { ok: true, titulos: e.titulos, tipos: e.tipos, opcoes: opcoes, proximoCodigo: maior + 1, total: linhas.length };
}

function lmCadastrar(p) {
  if (!p.confirmouDocumento) return { ok: false, erro: 'É preciso confirmar que você tem o documento do motoca.' };
  const e = lmCadastroEstrutura();
  if (e.erro) return { ok: false, erro: e.erro };
  const v = p.valores || {};
  const iCod = e.tipos.indexOf('codigo'), iCpf = e.tipos.indexOf('cpf'), iNome = e.tipos.indexOf('nome'), iTel = e.tipos.indexOf('telefone');
  const linhaNova = [];
  for (let i = 0; i < LM_CAD_COLS; i++) linhaNova.push(String(v[i] == null ? '' : v[i]).replace(/\s+/g, ' ').trim());

  // conferências
  if (iNome >= 0) { linhaNova[iNome] = linhaNova[iNome].toUpperCase(); if (linhaNova[iNome].length < 5) return { ok: false, erro: 'Escreva o nome completo.' }; }
  if (iCpf >= 0) {
    if (!lmCpfValido(linhaNova[iCpf])) return { ok: false, erro: 'Esse CPF não é válido. Confira os números.', campo: iCpf };
    const d = linhaNova[iCpf].replace(/\D/g, '');
    linhaNova[iCpf] = d.slice(0, 3) + '.' + d.slice(3, 6) + '.' + d.slice(6, 9) + '-' + d.slice(9);
  }
  if (iTel >= 0) {
    const d = linhaNova[iTel].replace(/\D/g, '');
    if (d.length < 10 || d.length > 11) return { ok: false, erro: 'Telefone com DDD, por favor (10 ou 11 números).', campo: iTel };
    linhaNova[iTel] = d.length === 11 ? '(' + d.slice(0, 2) + ') ' + d.slice(2, 7) + '-' + d.slice(7) : '(' + d.slice(0, 2) + ') ' + d.slice(2, 6) + '-' + d.slice(6);
  }
  for (let i = 1; i < LM_CAD_COLS; i++) {
    if (i === iCod) continue;
    if (!linhaNova[i] && ['nome', 'cpf', 'telefone', 'pix', 'tipoPix', 'unidade'].indexOf(e.tipos[i]) >= 0) return { ok: false, erro: 'Preencha: ' + e.titulos[i] + '.', campo: i };
  }

  // CPF ou código repetido? (lê a planilha agora, não a memória)
  const linhas = e.ultima >= 2 ? e.aba.getRange(2, 1, e.ultima - 1, LM_CAD_COLS).getDisplayValues() : [];
  let maior = 0;
  if (iCpf >= 0) {
    const d = linhaNova[iCpf].replace(/\D/g, '');
    for (let k = 0; k < linhas.length; k++) {
      if (String(linhas[k][iCpf]).replace(/\D/g, '') === d) {
        return { ok: false, duplicado: true, erro: 'Esse CPF já está cadastrado.',
                 codigo: iCod >= 0 ? String(linhas[k][iCod]).trim() : '', nome: iNome >= 0 ? String(linhas[k][iNome]).trim() : '' };
      }
    }
  }
  if (iCod >= 0) {
    linhas.forEach(function (l) { const n = parseInt(String(l[iCod]).replace(/\D/g, ''), 10); if (n > maior && n < 1e7) maior = n; });
    const cod = String(linhaNova[iCod]).replace(/\D/g, '') || String(maior + 1);
    const usado = linhas.some(function (l) { return String(l[iCod]).replace(/\D/g, '') === cod; });
    if (usado) return { ok: false, erro: 'O COD FUN ' + cod + ' já é de outro motoca. Sugestão: ' + (maior + 1) + '.', proximoCodigo: maior + 1, campo: iCod };
    linhaNova[iCod] = Number(cod);
  }
  linhaNova[0] = new Date();   // A: carimbo de data/hora

  const nova = e.ultima + 1;
  e.aba.getRange(nova, 1, 1, LM_CAD_COLS).setValues([linhaNova]);
  SpreadsheetApp.flush();
  CacheService.getScriptCache().remove(LM_CHAVE_CAD + '_0');   // a busca já encontra o motoca novo
  CacheService.getScriptCache().remove(LM_CHAVE_CAD);
  return { ok: true, codigo: iCod >= 0 ? String(linhaNova[iCod]) : '', nome: iNome >= 0 ? linhaNova[iNome] : '', linha: nova, proximoCodigo: Number(linhaNova[iCod] || maior) + 1 };
}

/* ============================== RELATAR PROBLEMA (aba "PROBLEMAS SITE") ============================== */

const LM_ABA_PROB = 'PROBLEMAS SITE';
const LM_PROB_TITULOS = ['ID', 'DATA E HORA', 'COD FUN', 'NOME', 'GRAVIDADE', 'O QUE ACONTECEU', 'REGISTRADO POR', 'ATIVO', 'ENCERRADO EM', 'MOTIVO DO ENCERRAMENTO'];

function lmAbaProb(criar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(LM_ABA_PROB);
  if (!aba && criar) {
    aba = ss.insertSheet(LM_ABA_PROB);
    aba.getRange(1, 1, 1, LM_PROB_TITULOS.length).setValues([LM_PROB_TITULOS]).setFontWeight('bold');
    aba.setFrozenRows(1);
    ['A', 'B', 'C', 'I'].forEach(function (c) { aba.getRange(c + ':' + c).setNumberFormat('@'); });
  }
  return aba;
}

function lmProbLerAba(aba) {
  if (!aba || aba.getLastRow() < 2) return [];
  const vals = aba.getRange(2, 1, aba.getLastRow() - 1, LM_PROB_TITULOS.length).getDisplayValues();
  const lista = [];
  vals.forEach(function (v, i) {
    if (!String(v[0]).trim()) return;
    lista.push({ id: String(v[0]).trim(), quando: String(v[1]).trim(), cod: String(v[2]).trim(), nome: String(v[3]).trim(),
                 gravidade: String(v[4]).trim().toLowerCase(), texto: String(v[5]).trim(), por: String(v[6]).trim(),
                 ativo: String(v[7]).trim().toUpperCase() !== 'NÃO', encerradoEm: String(v[8]).trim(), motivo: String(v[9]).trim(), _linhaAba: i + 2 });
  });
  return lista;
}

function lmProblemas() {   // lista para a tela (guardada 10 min; refeita a cada mudança)
  const cache = CacheService.getScriptCache();
  const c = cache.get('prob_v1');
  if (c) { try { return JSON.parse(c); } catch (e) {} }
  const lista = lmProbLerAba(lmAbaProb(false)).map(function (x) { delete x._linhaAba; return x; }).slice(-600);
  try { cache.put('prob_v1', JSON.stringify(lista), 600); } catch (e) {}
  return lista;
}

function lmProblemaNovo(p) {
  const grav = String(p.gravidade || '').toLowerCase();
  if (['leve', 'media', 'grave'].indexOf(grav) < 0) return { ok: false, erro: 'Escolha a gravidade: leve, média ou grave.' };
  const texto = String(p.texto || '').replace(/\s+/g, ' ').trim();
  if (texto.length < 5) return { ok: false, erro: 'Conte o que aconteceu (pelo menos algumas palavras).' };
  const cod = String(p.codigo || '').replace(/\D/g, '');
  if (!cod) return { ok: false, erro: 'Digite o código (COD FUN) do motoca.' };
  const cad = lmLerCadastro(false);
  let nome = '';
  if (!cad.erro) {
    for (let i = 0; i < cad.lista.length; i++) if (String(cad.lista[i][0]).replace(/\D/g, '') === cod) { nome = cad.lista[i][1]; break; }
    if (!nome) return { ok: false, erro: 'Não achei o COD FUN ' + cod + ' no cadastro.' };
  }
  const aba = lmAbaProb(true);
  const id = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
  const quando = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
  const por = String(p.por || '').replace(/\s+/g, ' ').trim().slice(0, 60);
  aba.appendRow([id, quando, cod, nome, grav === 'media' ? 'média' : grav, texto.slice(0, 1000), por, 'SIM', '', '']);
  CacheService.getScriptCache().remove('prob_v1');
  return { ok: true, item: { id: id, quando: quando, cod: cod, nome: nome, gravidade: grav === 'media' ? 'média' : grav, texto: texto.slice(0, 1000), por: por, ativo: true, encerradoEm: '', motivo: '' } };
}

function lmProblemaEncerrar(id, motivo) {
  const aba = lmAbaProb(false);
  const x = lmProbLerAba(aba).filter(function (i) { return i.id === String(id); })[0];
  if (!x) return { ok: false, erro: 'Não achei esse registro.' };
  const quando = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
  aba.getRange(x._linhaAba, 8, 1, 3).setValues([['NÃO', quando, String(motivo || '').slice(0, 300)]]);
  CacheService.getScriptCache().remove('prob_v1');
  x.ativo = false; x.encerradoEm = quando; x.motivo = String(motivo || ''); delete x._linhaAba;
  return { ok: true, item: x };
}

/* ============================== PAGAMENTO NO FINAL (aba "PGF SITE") ============================== */

const LM_ABA_PGF = 'PGF SITE';
const LM_PGF_TITULOS = ['ID', 'DATA', 'LOJA', 'HORÁRIO', 'CÓDIGO', 'NOME', 'TELEFONE', 'CHAVE PIX', 'VALOR', 'MARCADO EM', 'PAGO', 'PAGO EM', 'LINHA DA ESCALA'];
const LM_PGF_DIAS = 15;   // a tela recebe os PGF dos últimos 15 dias (e dos próximos)

function lmAbaPgf(criar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(LM_ABA_PGF);
  if (!aba && criar) {
    aba = ss.insertSheet(LM_ABA_PGF);
    aba.getRange(1, 1, 1, LM_PGF_TITULOS.length).setValues([LM_PGF_TITULOS]).setFontWeight('bold');
    aba.setFrozenRows(1);
    ['A', 'B', 'E', 'G', 'H', 'M'].forEach(function (c) { aba.getRange(c + ':' + c).setNumberFormat('@'); });   // como texto
    aba.getRange('I:I').setNumberFormat('"R$" #,##0.00');
  }
  return aba;
}

function lmNumero(v) {   // "80", "80,50", "R$ 1.080,50" → número
  let t = String(v == null ? '' : v).replace(/[^\d,.\-]/g, '');
  if (t.indexOf(',') >= 0) t = t.replace(/\./g, '').replace(',', '.');
  const n = parseFloat(t);
  return isFinite(n) ? Math.round(n * 100) / 100 : NaN;
}

function lmPgfLerAba(aba) {
  if (!aba || aba.getLastRow() < 2) return [];
  const n = aba.getLastRow() - 1, faixa = aba.getRange(2, 1, n, LM_PGF_TITULOS.length);
  const vals = faixa.getDisplayValues(), brutos = faixa.getValues();
  const lista = [];
  vals.forEach(function (v, i) {
    if (!String(v[0]).trim()) return;
    const valor = typeof brutos[i][8] === 'number' ? brutos[i][8] : lmNumero(v[8]);
    lista.push({ id: String(v[0]).trim(), data: String(v[1]).trim(), loja: String(v[2]).trim(), horario: String(v[3]).trim(),
                 cod: String(v[4]).trim(), nome: String(v[5]).trim(), tel: String(v[6]).trim(), pix: String(v[7]).trim(),
                 valor: isFinite(valor) ? valor : 0, marcadoEm: String(v[9]).trim(), pago: String(v[10]).trim().toUpperCase() === 'SIM',
                 pagoEm: String(v[11]).trim(), linha: String(v[12]).trim(), whats: lmLimparTelefone(v[6]), _linhaAba: i + 2 });
  });
  return lista;
}

// Lista para a tela (guardada 10 min; é refeita a cada mudança)
function lmPgfLista() {
  const cache = CacheService.getScriptCache();
  const c = cache.get('pgf_v1');
  if (c) { try { return JSON.parse(c); } catch (e) {} }
  const hoje = new Date(), limite = Date.UTC(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - LM_PGF_DIAS);
  const lista = lmPgfLerAba(lmAbaPgf(false)).filter(function (x) {
    const m = x.data.match(/(\d{2})\/(\d{2})\/(\d{4})/);
    return !m || Date.UTC(+m[3], +m[2] - 1, +m[1]) >= limite;
  }).map(function (x) { delete x._linhaAba; return x; });
  try { cache.put('pgf_v1', JSON.stringify(lista), 600); } catch (e) {}
  return lista;
}

// Botão PGF: grava (ou muda) o valor combinado daquela pessoa naquele dia
function lmPgfMarcar(p) {
  const valor = lmNumero(p.valor);
  if (!(valor > 0) || valor > 10000) return { ok: false, erro: 'Digite um valor válido (ex.: 80 ou 80,50).' };
  const achou = lmLocalizar(p);
  if (achou.erro) return achou.erro;
  const atual = achou.atual, col = achou.col, linha = achou.linha;
  const limpa = function (x) { return String(x == null ? '' : x).replace(/\s*[\r\n]+\s*/g, ' ').trim(); };
  const nome = limpa(atual[col.nome]);
  if (!nome || /N[ÃA]O\s+LOCALIZAD/i.test(nome)) return { ok: false, erro: 'Vaga sem motoca não tem pagamento.' };
  const cod = col.cod >= 0 ? limpa(atual[col.cod]) : '', tel = col.tel >= 0 ? limpa(atual[col.tel]) : '';
  const horario = col.hora >= 0 ? limpa(atual[col.hora]) : '', data = limpa(atual[0]);

  let pix = '';   // chave PIX vem do cadastro
  const cad = lmLerCadastro(false);
  if (!cad.erro) {
    const so = function (x) { return String(x).replace(/\D/g, ''); };
    let m = null;
    for (let i = 0; i < cad.lista.length && !m; i++) if (cod && so(cad.lista[i][0]) && so(cad.lista[i][0]) === so(cod)) m = cad.lista[i];
    for (let i = 0; i < cad.lista.length && !m; i++) if (cad.lista[i][1].toUpperCase() === nome.toUpperCase()) m = cad.lista[i];
    if (m) pix = m[4] || '';
  }

  const aba = lmAbaPgf(true);
  const existente = lmPgfLerAba(aba).filter(function (x) { return x.loja === p.loja && x.data === data && x.linha === String(linha); })[0];
  const agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
  let item;
  if (existente) {
    if (existente.pago) return { ok: false, erro: 'Esse PGF já foi pago. Desfaça o "Pago" antes de mudar o valor.' };
    aba.getRange(existente._linhaAba, 9).setValue(valor);
    aba.getRange(existente._linhaAba, 6, 1, 3).setValues([[nome, tel, pix || existente.pix]]);
    item = existente; item.valor = valor; item.nome = nome; item.tel = tel; item.pix = pix || existente.pix; item.whats = lmLimparTelefone(tel);
  } else {
    const id = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    aba.appendRow([id, data, p.loja, horario, cod, nome, tel, pix, valor, agora, '', '', String(linha)]);
    item = { id: id, data: data, loja: p.loja, horario: horario, cod: cod, nome: nome, tel: tel, pix: pix, valor: valor,
             marcadoEm: agora, pago: false, pagoEm: '', linha: String(linha), whats: lmLimparTelefone(tel) };
  }
  delete item._linhaAba;
  CacheService.getScriptCache().remove('pgf_v1');
  return { ok: true, item: item, semPix: !item.pix };
}

function lmPgfPago(id, pago) {
  const aba = lmAbaPgf(false);
  const x = lmPgfLerAba(aba).filter(function (i) { return i.id === String(id); })[0];
  if (!x) return { ok: false, erro: 'Não achei esse PGF na planilha (pode ter sido apagado).', sumiu: true };
  const agora = pago ? Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm') : '';
  aba.getRange(x._linhaAba, 11, 1, 2).setValues([[pago ? 'SIM' : '', agora]]);
  CacheService.getScriptCache().remove('pgf_v1');
  x.pago = pago; x.pagoEm = agora; delete x._linhaAba;
  return { ok: true, item: x };
}

function lmPgfRemover(id) {
  const aba = lmAbaPgf(false);
  const x = lmPgfLerAba(aba).filter(function (i) { return i.id === String(id); })[0];
  if (!x) return { ok: true, removido: true };
  if (x.pago) return { ok: false, erro: 'Esse PGF já foi pago. Desfaça o "Pago" antes de remover.' };
  aba.deleteRow(x._linhaAba);
  CacheService.getScriptCache().remove('pgf_v1');
  return { ok: true, removido: true };
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
  lmGuardarTexto(LM_CHAVE_TEL, JSON.stringify(r.tel), Date.now());
  console.log('Telefones encontrados: ' + Object.keys(r.tel).length + ' (só vão para a tela do operador).');
  console.log('Escala montada: ' + r.linhas + ' linhas em ' + r.segundos.toFixed(1) + ' segundos (' +
    Math.round(Utilities.newBlob(r.csv).getBytes().length / 1024) + ' KB).');
  console.log(r.lentas.length ? 'Abas lentas (mais de 5 s): ' + r.lentas.join(', ') : 'Nenhuma aba lenta.');
  if (r.puladas.length) console.log('Ficaram de fora por tempo: ' + r.puladas.join(', '));
  console.log('Guardada ok? ' + (lmLerGuardada() ? 'SIM' : 'NÃO'));
  console.log('Agora faça: Implantar → Gerenciar implantações → lápis → Nova versão → Implantar.');
}

/** Diz, em cada loja, se a coluna NOME puxa o nome pelo código (fórmula). Não mostra dados pessoais. */
function conferirNomePorCodigo() {
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function (aba) {
    if (!lmAbaDeLoja(aba)) return;
    const ultima = aba.getLastRow();
    if (ultima < 3) return;
    const inicio = Math.max(1, ultima - 60 + 1);
    const faixa = aba.getRange(inicio, 1, ultima - inicio + 1, Math.min(LM_COLUNAS, aba.getMaxColumns()));
    const valores = faixa.getDisplayValues(), formulas = faixa.getFormulas();
    let col = null, com = 0, sem = 0;
    valores.forEach(function (l, i) {
      if (String(l[0]).trim().toUpperCase() === 'DATA') { col = lmColunas(l); return; }
      if (!col || col.nome < 0 || !/^\d{2}\/\d{2}\/\d{4}$/.test(String(l[0]).trim())) return;
      formulas[i][col.nome] ? com++ : sem++;
    });
    console.log(aba.getName() + ': ' + com + ' linhas com fórmula no NOME, ' + sem + ' sem fórmula');
  });
}

/** Libera na hora a trava de "Muitas senhas erradas" (de todos os aparelhos). */
function liberarSenha() {
  const cache = CacheService.getScriptCache();
  cache.remove('lm_erros_senha'); cache.remove('lm_erros_geral'); cache.remove('lm_trava_geral');
  console.log('Pronto: trava geral liberada. (A trava de um aparelho some sozinha em até 1 minuto.)');
}

/** Mostra quais colunas do cadastro a busca usa (só os títulos e quantos motoboys; nenhum dado pessoal). */
function conferirCadastro() {
  const cad = lmLerCadastro(true);
  if (cad.erro) { console.log('PROBLEMA: ' + cad.erro); return; }
  console.log('Aba de cadastro: ' + cad.aba + ' (títulos na linha ' + cad.linhaTitulos + ')');
  console.log('Código: ' + (cad.titulos.cod || 'NÃO ACHEI') + ' | Nome: ' + (cad.titulos.nome || 'NÃO ACHEI') +
    ' | Telefone: ' + (cad.titulos.tel || 'NÃO ACHEI') + ' | CPF: ' + (cad.titulos.cpf || 'NÃO ACHEI') +
    ' | PIX: ' + (cad.titulos.pix || 'NÃO ACHEI'));
  console.log('Motoboys no cadastro: ' + cad.lista.length);
  const e = lmCadastroEstrutura();
  if (!e.erro) console.log('Colunas A–H (Novo Cadastro): ' + e.titulos.map(function (t, i) { return 'ABCDEFGH'[i] + ' = ' + (t || '(vazia)') + ' [' + e.tipos[i] + ']'; }).join(' | '));
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
