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
 * O QUE ENTREGA: loja, data, horário, código, nome, presença, nº da linha e hora da chegada.
 * Telefone: só para a tela do operador (com senha). NUNCA entrega CPF, PIX nem valores.
 *
 * FUNÇÕES PARA RODAR NO EDITOR (botão Executar):
 *   testarTudo           → monta a escala e mostra quantas linhas, quantos segundos e quais abas são lentas.
 *   diagnosticarLentidao → mede o tempo de cada aba, uma por uma (sem dados pessoais).
 *   conferirNomePorCodigo → diz, em cada loja, se a coluna NOME puxa o nome pelo código (fórmula).
 *   conferirCadastro     → mostra quais colunas do cadastro a busca usa (só os títulos, sem dados pessoais).
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

  const cache = CacheService.getScriptCache();
  const erros = +(cache.get('lm_erros_senha') || 0);
  if (erros >= LM_MAX_ERROS_SENHA) return lmResposta({ ok: false, erro: 'Muitas senhas erradas. Espere 10 minutos.', senha: true });
  if (String(pedido.senha || '') !== LM_SENHA) {
    cache.put('lm_erros_senha', String(erros + 1), 600);
    return lmResposta({ ok: false, erro: 'Senha errada.', senha: true });
  }

  if (pedido.acao === 'entrar') return lmResposta({ ok: true, planilha: SpreadsheetApp.getActiveSpreadsheet().getName() });
  if (pedido.acao === 'ler') {
    const csv = lmEscala();   // primeiro a escala (se precisar, ela também guarda os telefones)
    return lmResposta({ ok: true, csv: csv, tel: lmLerTelefones(), online: lmOnline('o', pedido.sessao) });
  }
  if (pedido.acao === 'buscar') return lmResposta(lmBuscar(pedido.q, pedido.campo));
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

/** Mostra quais colunas do cadastro a busca usa (só os títulos e quantos motoboys; nenhum dado pessoal). */
function conferirCadastro() {
  const cad = lmLerCadastro(true);
  if (cad.erro) { console.log('PROBLEMA: ' + cad.erro); return; }
  console.log('Aba de cadastro: ' + cad.aba + ' (títulos na linha ' + cad.linhaTitulos + ')');
  console.log('Código: ' + (cad.titulos.cod || 'NÃO ACHEI') + ' | Nome: ' + (cad.titulos.nome || 'NÃO ACHEI') +
    ' | Telefone: ' + (cad.titulos.tel || 'NÃO ACHEI') + ' | CPF: ' + (cad.titulos.cpf || 'NÃO ACHEI') +
    ' | PIX: ' + (cad.titulos.pix || 'NÃO ACHEI'));
  console.log('Motoboys no cadastro: ' + cad.lista.length);
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
