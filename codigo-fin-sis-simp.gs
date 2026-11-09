/**
 * FIN_SIS_SIMP — FOLHA GERAL (aba Financeiro da tela do operador)
 *
 * É um PROJETO SEPARADO do SIST_SIMP (script.google.com → Novo projeto → nome FIN_SIS_SIMP).
 * Não mexe em nada do SIST_SIMP. Tem o seu próprio endereço /exec.
 *
 * O que faz:
 *  - lê das abas das lojas: COD, NOME, TELEFONE, HORÁRIO, PIX, TOTAL A PAGAR e STATUS
 *    (de 7 dias atrás até hoje);
 *  - o site SÓ muda o STATUS (uma pessoa por vez), usando as opções do menu da própria planilha;
 *  - "Todos pagos" (por loja e dia) pede a senha da folha e muda para PAGO só quem está pendente.
 *
 * Implantar: Implantar → Nova implantação → Tipo: App da Web → Executar como: Eu → Quem pode acessar: Qualquer pessoa.
 * Depois, para atualizar: Implantar → Gerenciar implantações → lápis → Nova versão (nunca "Nova implantação" de novo).
 */

const FIN_PLANILHA_ID = '1SUSjukHorUEWx1d-U5KijU8BuzCPmglW0WisuygqcXo';   // ESCALAS LOGMAX - FOOD
const FIN_SENHA_OPERADOR = 'COLOQUE_A_SENHA_AQUI';   // a MESMA senha do site do operador (se trocar lá, troque aqui também)
const FIN_SENHA_TODOS = 'COLOQUE_A_SENHA_DA_FOLHA';  // senha do botão "Todos pagos"
const FIN_ABAS_IGNORADAS = ['RESUMO', 'FOLHA DE PAGAMENTO', 'CADASTRO MOTORISTA', 'ENDEREÇOS', 'PAG FNAL', 'PGF SITE', 'PROBLEMAS SITE'];
const FIN_DIAS_ANTES = 7;          // mostra de 7 dias atrás...
const FIN_DIAS_DEPOIS = 0;         // ...até hoje
const FIN_LINHAS_FINAIS = 700;     // lê só o final de cada aba
const FIN_COLUNAS = 20;            // lê até a coluna T
const FIN_FRESCO_SEG = 60;         // a folha guardada vale 60 s sem ler a planilha de novo
const FIN_GUARDA_SEG = 21600;      // guarda a última folha por até 6 h
const FIN_CHAVE = 'folha_v1';
const FIN_PEDACO = 40000;
const FIN_OPCOES_PADRAO = ['PENDENTE', 'PAGO', 'NÃO PAGAR'];

/* ============================== ENTRADA ============================== */

function doGet() {
  return ContentService.createTextOutput('FIN_SIS_SIMP no ar.');
}

function doPost(e) {
  let p;
  try { p = JSON.parse(e.postData.contents); } catch (err) { return finResposta({ ok: false, erro: 'Pedido inválido.' }); }
  const cache = CacheService.getScriptCache();
  const aparelho = /^[A-Za-z0-9]{6,40}$/.test(String(p.sessao || '')) ? String(p.sessao) : 'sem-id';

  // trava por aparelho: 5 senhas erradas → espera 1 minuto
  const ate = +(cache.get('fin_trava_' + aparelho) || 0);
  if (ate > Date.now()) return finResposta({ ok: false, erro: 'Muitas senhas erradas. Tente de novo em ' + Math.ceil((ate - Date.now()) / 1000) + ' s.' });
  if (String(p.senha || '') !== FIN_SENHA_OPERADOR) {
    finErro(cache, aparelho);
    return finResposta({ ok: false, senha: true, erro: 'A senha da Folha Geral não confere com a do site (confira FIN_SENHA_OPERADOR no FIN_SIS_SIMP).' });
  }

  if (p.acao === 'folha') return finResposta(finFolha());

  if (p.acao === 'folhaStatus' || p.acao === 'folhaTodosPagos') {
    if (p.acao === 'folhaTodosPagos' && String(p.senhaFolha || '') !== FIN_SENHA_TODOS) {
      finErro(cache, aparelho);
      return finResposta({ ok: false, erro: 'Senha da folha errada.' });
    }
    const trava = LockService.getScriptLock();
    if (!trava.tryLock(15000)) return finResposta({ ok: false, erro: 'A planilha está ocupada. Tente de novo.' });
    try {
      cache.remove('fin_erros_' + aparelho);
      return finResposta(p.acao === 'folhaStatus' ? finMudarStatus(p) : finTodosPagos(p));
    } finally { trava.releaseLock(); }
  }
  return finResposta({ ok: false, erro: 'Ação desconhecida.' });
}

function finErro(cache, aparelho) {
  const n = +(cache.get('fin_erros_' + aparelho) || 0) + 1;
  if (n >= 5) { cache.put('fin_trava_' + aparelho, String(Date.now() + 60000), 60); cache.remove('fin_erros_' + aparelho); }
  else cache.put('fin_erros_' + aparelho, String(n), 600);
}

function finResposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ============================== LER A FOLHA ============================== */

function finFolha() {
  const g = finLerGuardada();
  if (g && Date.now() - g.t < FIN_FRESCO_SEG * 1000) return g.dados;
  const lock = LockService.getUserLock();   // só uma montagem por vez (separada da trava de gravação)
  if (!lock.tryLock(g ? 100 : 280000)) {
    if (g) return g.dados;   // alguém já está montando: devolve a última pronta
    return { ok: false, erro: 'A folha está sendo montada. Tente de novo em instantes.' };
  }
  try {
    const g2 = finLerGuardada();
    if (g2 && Date.now() - g2.t < FIN_FRESCO_SEG * 1000) return g2.dados;
    const dados = finMontar();
    finGuardar(dados);
    return dados;
  } finally { lock.releaseLock(); }
}

function finSemAcento(x) { return String(x).toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim(); }

function finColunas(l) {
  const t = l.map(finSemAcento);
  const achar = function (f) { return t.findIndex(f); };
  const col = {
    cod: achar(function (c) { return c.indexOf('COD') >= 0; }),
    nome: achar(function (c) { return c.indexOf('NOME') >= 0; }),
    tel: achar(function (c) { return /TELEF|CELULAR|FONE/.test(c); }),
    hora: achar(function (c) { return c.indexOf('HORARIO') >= 0; }),
    pix: achar(function (c) { return c.indexOf('PIX') >= 0 && c.indexOf('TIPO') < 0; }),
    total: achar(function (c) { return c.indexOf('TOTAL A PAGAR') >= 0; }),
    status: achar(function (c) { return c === 'STATUS'; })
  };
  if (col.total < 0) col.total = achar(function (c) { return /A PAGAR$/.test(c); });
  if (col.status < 0) col.status = achar(function (c) { return c.indexOf('STATUS') >= 0; });
  return col;
}

function finAbaDeLoja(aba) {
  const nome = aba.getName().trim();
  if (FIN_ABAS_IGNORADAS.indexOf(nome.toUpperCase()) >= 0) return false;
  if (/^ontem/i.test(nome) || aba.isSheetHidden()) return false;
  return true;
}

function finNumero(v) {   // "R$ 1.080,50" → 1080.5 ; "R$ -" → 0
  let t = String(v == null ? '' : v).replace(/[^\d,.\-]/g, '');
  if (t.indexOf(',') >= 0) t = t.replace(/\./g, '').replace(',', '.');
  const n = parseFloat(t);
  return isFinite(n) ? Math.round(n * 100) / 100 : 0;
}

function finOpcoesDaCelula(cel) {   // lê as opções do menu (validação) da célula STATUS
  try {
    const v = cel.getDataValidation();
    if (!v) return null;
    const tipo = v.getCriteriaType(), val = v.getCriteriaValues();
    if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return val[0].map(String).filter(String);
    if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) return val[0].getDisplayValues().map(function (r) { return String(r[0]).trim(); }).filter(String);
  } catch (e) {}
  return null;
}

function finPixCadastro(planilha) {   // código → chave PIX (para quem está sem PIX na aba da loja)
  const mapa = {};
  try {
    const aba = planilha.getSheets().filter(function (a) { return /CADASTRO/i.test(a.getName()); })[0];
    if (!aba || aba.getLastRow() < 2) return mapa;
    const d = aba.getRange(1, 1, aba.getLastRow(), Math.min(12, aba.getMaxColumns())).getDisplayValues();
    const t = d[0].map(finSemAcento);
    const cCod = t.findIndex(function (c) { return c.indexOf('COD') >= 0; });
    const cPix = t.findIndex(function (c) { return c.indexOf('PIX') >= 0 && c.indexOf('TIPO') < 0; });
    if (cCod < 0 || cPix < 0) return mapa;
    for (let i = 1; i < d.length; i++) {
      const c = String(d[i][cCod]).trim(), x = String(d[i][cPix]).trim();
      if (c && x) mapa[c] = x;
    }
  } catch (e) {}
  return mapa;
}

function finMontar() {
  const comeco = Date.now();
  const planilha = SpreadsheetApp.openById(FIN_PLANILHA_ID);
  const hoje = new Date();
  const min = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - FIN_DIAS_ANTES);
  const max = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() + FIN_DIAS_DEPOIS);
  const pixCad = finPixCadastro(planilha);
  const itens = [];
  let opcoes = null;
  const puladas = [];

  planilha.getSheets().forEach(function (aba) {
    if (!finAbaDeLoja(aba)) return;
    const loja = aba.getName().trim();
    if (Date.now() - comeco > 270000) { puladas.push(loja); return; }
    const ultima = aba.getLastRow();
    if (ultima < 2) return;
    const inicio = Math.max(1, ultima - FIN_LINHAS_FINAIS + 1);
    const linhas = aba.getRange(inicio, 1, ultima - inicio + 1, Math.min(FIN_COLUNAS, aba.getMaxColumns())).getDisplayValues();
    let col = null;
    linhas.forEach(function (l, i) {
      if (String(l[0]).trim().toUpperCase() === 'DATA') { col = finColunas(l); return; }
      const m = String(l[0]).trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
      if (!m || !col || col.total < 0 || col.status < 0) return;
      const d = new Date(+m[3], +m[2] - 1, +m[1]);
      if (d < min || d > max) return;
      const pega = function (k) { return k >= 0 ? String(l[k]).replace(/\s*[\r\n]+\s*/g, ' ').trim() : ''; };
      const nome = pega(col.nome), cod = pega(col.cod);
      if (!nome || /N[ÃA]O LOCALIZADO/i.test(nome)) return;
      if (!opcoes) opcoes = finOpcoesDaCelula(aba.getRange(inicio + i, col.status + 1));
      const pix = pega(col.pix) || pixCad[cod] || '';
      itens.push({ id: loja + '|' + (inicio + i), loja: loja, linha: inicio + i, data: String(l[0]).trim(), cod: cod, nome: nome,
                   tel: pega(col.tel), horario: pega(col.hora), pix: pix, total: finNumero(pega(col.total)), status: pega(col.status) });
    });
  });
  if (puladas.length) console.warn('Passou de 4,5 min. Ficaram de fora: ' + puladas.join(', '));
  return { ok: true, opcoes: opcoes && opcoes.length ? opcoes : FIN_OPCOES_PADRAO, itens: itens, montadaEm: new Date().toISOString() };
}

/* ============================== GUARDAR (em pedaços) ============================== */

function finGuardar(dados) {
  const texto = JSON.stringify(dados), cache = CacheService.getScriptCache();
  const n = Math.ceil(texto.length / FIN_PEDACO) || 1, obj = {};
  for (let i = 0; i < n; i++) obj[FIN_CHAVE + '_' + i] = texto.slice(i * FIN_PEDACO, (i + 1) * FIN_PEDACO);
  obj[FIN_CHAVE + '_n'] = String(n);
  obj[FIN_CHAVE + '_t'] = String(Date.now());
  try { cache.putAll(obj, FIN_GUARDA_SEG); } catch (e) { console.warn('Não guardou a folha: ' + e); }
}

function finLerGuardada() {
  try {
    const cache = CacheService.getScriptCache();
    const base = cache.getAll([FIN_CHAVE + '_n', FIN_CHAVE + '_t']);
    const n = +base[FIN_CHAVE + '_n'];
    if (!n) return null;
    const chaves = []; for (let i = 0; i < n; i++) chaves.push(FIN_CHAVE + '_' + i);
    const ps = cache.getAll(chaves);
    let texto = '';
    for (let i = 0; i < n; i++) { if (ps[chaves[i]] == null) return null; texto += ps[chaves[i]]; }
    return { dados: JSON.parse(texto), t: +base[FIN_CHAVE + '_t'] || 0 };
  } catch (e) { return null; }
}

function finMudarGuardada(mudancas) {   // { id: status } → atualiza a folha guardada sem reler tudo
  const g = finLerGuardada();
  if (!g) return;
  g.dados.itens.forEach(function (x) { if (mudancas.hasOwnProperty(x.id)) x.status = mudancas[x.id]; });
  const t = g.t;
  finGuardar(g.dados);
  try { CacheService.getScriptCache().put(FIN_CHAVE + '_t', String(t), FIN_GUARDA_SEG); } catch (e) {}
}

/* ============================== GRAVAR STATUS ============================== */

function finAchaColunas(aba, linha) {   // procura a linha de títulos "DATA" logo acima
  const ini = Math.max(1, linha - 80);
  const bloco = aba.getRange(ini, 1, linha - ini + 1, Math.min(FIN_COLUNAS, aba.getMaxColumns())).getDisplayValues();
  for (let i = bloco.length - 2; i >= 0; i--) {
    if (String(bloco[i][0]).trim().toUpperCase() === 'DATA') return { col: finColunas(bloco[i]), linhaDados: bloco[bloco.length - 1] };
  }
  return null;
}

function finMudarStatus(p) {
  const planilha = SpreadsheetApp.openById(FIN_PLANILHA_ID);
  const aba = planilha.getSheetByName(String(p.loja || ''));
  const linha = parseInt(p.linha, 10);
  if (!aba || !finAbaDeLoja(aba) || !(linha >= 2)) return { ok: false, erro: 'Não achei essa loja/linha na planilha.', mudou: true };
  const achou = finAchaColunas(aba, linha);
  if (!achou || achou.col.status < 0) return { ok: false, erro: 'Não achei a coluna STATUS desse dia.', mudou: true };
  const nomeAgora = String(achou.linhaDados[achou.col.nome] || '').trim();
  if (finSemAcento(nomeAgora) !== finSemAcento(p.nome)) return { ok: false, erro: 'A planilha mudou (nessa linha agora está ' + (nomeAgora || 'outra pessoa') + '). Atualizando…', mudou: true };
  const cel = aba.getRange(linha, achou.col.status + 1);
  const opcoes = finOpcoesDaCelula(cel) || FIN_OPCOES_PADRAO;
  const novo = String(p.status || '');
  if (opcoes.indexOf(novo) < 0) return { ok: false, erro: 'Status inválido: ' + novo };
  cel.setValue(novo);
  const id = aba.getName().trim() + '|' + linha, m = {}; m[id] = novo;
  finMudarGuardada(m);
  return { ok: true, item: { id: id, status: novo } };
}

function finTodosPagos(p) {
  const planilha = SpreadsheetApp.openById(FIN_PLANILHA_ID);
  const aba = planilha.getSheetByName(String(p.loja || ''));
  const data = String(p.data || '').trim();
  if (!aba || !finAbaDeLoja(aba) || !/^\d{2}\/\d{2}\/\d{4}$/.test(data)) return { ok: false, erro: 'Não achei essa loja/dia.' };
  const ultima = aba.getLastRow();
  const inicio = Math.max(1, ultima - FIN_LINHAS_FINAIS + 1);
  const linhas = aba.getRange(inicio, 1, ultima - inicio + 1, Math.min(FIN_COLUNAS, aba.getMaxColumns())).getDisplayValues();
  let col = null, pago = null;
  const mudados = [], m = {};
  linhas.forEach(function (l, i) {
    if (String(l[0]).trim().toUpperCase() === 'DATA') { col = finColunas(l); return; }
    if (String(l[0]).trim() !== data || !col || col.status < 0) return;
    const nome = String(l[col.nome] || '').trim();
    if (!nome || /N[ÃA]O LOCALIZADO/i.test(nome)) return;
    const st = finSemAcento(l[col.status]);
    if (/^PAGO/.test(st) || /NAO/.test(st)) return;   // já pago ou "não pagar": não mexe
    const cel = aba.getRange(inicio + i, col.status + 1);
    if (!pago) {
      const ops = finOpcoesDaCelula(cel) || FIN_OPCOES_PADRAO;
      pago = ops.filter(function (o) { return /^PAGO/.test(finSemAcento(o)); })[0] || 'PAGO';
    }
    cel.setValue(pago);
    const id = aba.getName().trim() + '|' + (inicio + i);
    m[id] = pago; mudados.push({ id: id, status: pago });
  });
  if (mudados.length) finMudarGuardada(m);
  return { ok: true, mudou: mudados.length, itens: mudados };
}

/* ============================== TESTE (rode no editor) ============================== */

function testarFolha() {
  const t0 = Date.now();
  const d = finMontar();
  const porLoja = {};
  d.itens.forEach(function (x) { porLoja[x.loja] = (porLoja[x.loja] || 0) + 1; });
  console.log('Opções do STATUS: ' + JSON.stringify(d.opcoes));
  console.log('Pessoas na folha (7 dias): ' + d.itens.length + ' | sem PIX: ' + d.itens.filter(function (x) { return !x.pix; }).length);
  console.log('Por loja: ' + JSON.stringify(porLoja));
  console.log('Tamanho: ' + JSON.stringify(d).length + ' caracteres | Tempo: ' + ((Date.now() - t0) / 1000).toFixed(1) + ' s');
  finGuardar(d);
}
