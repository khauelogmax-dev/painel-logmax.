/**
 * LogMax · Presença ao vivo
 * Cole este código DENTRO da planilha "ESCALAS LOGMAX - FOOD" (Extensões → Apps Script)
 * e implante como "App da Web". A tela chama este endereço a cada 30 segundos.
 *
 * Entrega SÓ: loja, data, horário, código, nome, presença e telefone (para o botão de WhatsApp).
 * NÃO entrega CPF, PIX nem valores.
 * Período: de HOJE até o DOMINGO desta semana.
 */
const ABAS_IGNORADAS = ['RESUMO', 'FOLHA DE PAGAMENTO', 'CADASTRO MOTORISTA', 'ENDEREÇOS', 'PAG FNAL'];
const GUARDAR_SEGUNDOS = 25;  // várias telas abertas ao mesmo tempo usam a mesma leitura
const LINHAS_FINAIS = 700;    // lê só o final de cada aba (onde ficam os dias recentes)
const COLUNAS = 8;            // DATA, COD, NOME, TELEFONE, HORARIO, PRESENÇA... (lê só até a coluna H)

function doGet() {
  const cache = CacheService.getScriptCache();
  let csv = cache.get('presenca');
  if (!csv) {
    csv = montarCsv();
    if (csv.length < 95000) cache.put('presenca', csv, GUARDAR_SEGUNDOS);
  }
  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV);
}

function montarCsv() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  // hoje no horário de São Paulo, e o domingo desta semana
  const p = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd').split('-');
  const min = new Date(+p[0], +p[1] - 1, +p[2]);
  const diaSemana = min.getDay();                                   // 0 = domingo
  const max = new Date(+p[0], +p[1] - 1, +p[2] + (diaSemana === 0 ? 0 : 7 - diaSemana));
  const agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
  const saida = [['LOJA', 'DATA', 'HORARIO', 'CODIGO', 'NOME', 'PRESENCA', 'TELEFONE', 'ATUALIZADO']];

  planilha.getSheets().forEach(function (aba) {
    const loja = aba.getName().trim();
    if (ABAS_IGNORADAS.indexOf(loja.toUpperCase()) >= 0) return;
    if (/^ontem/i.test(loja) || aba.isSheetHidden()) return;   // cópias de relatório e abas escondidas
    const ultima = aba.getLastRow();
    if (ultima < 2) return;
    const inicio = Math.max(1, ultima - LINHAS_FINAIS + 1);
    const linhas = aba.getRange(inicio, 1, ultima - inicio + 1, Math.min(COLUNAS, aba.getMaxColumns())).getDisplayValues();
    let col = null; // posição das colunas do bloco atual

    linhas.forEach(function (l) {
      const primeira = String(l[0]).trim().toUpperCase();
      if (primeira === 'DATA') {
        const achar = function (txt) { return l.findIndex(function (c) { return String(c).toUpperCase().indexOf(txt) >= 0; }); };
        col = { cod: achar('COD'), nome: achar('NOME'), hora: achar('HORARIO'), pres: achar('PRESEN'), tel: achar('TELEFONE') };
        return;
      }
      const m = String(l[0]).trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
      if (!m || !col || col.pres < 0) return;
      const d = new Date(+m[3], +m[2] - 1, +m[1]);
      if (d < min || d > max) return;
      const pega = function (i) { return i >= 0 ? String(l[i]).trim() : ''; };
      saida.push([loja, String(l[0]).trim(), pega(col.hora), pega(col.cod), pega(col.nome), pega(col.pres), pega(col.tel), '']);
    });
  });

  if (saida.length > 1) saida[1][7] = agora;
  else saida.push(['', '', '', '', '', '', '', agora]);
  return saida.map(function (linha) {
    return linha.map(function (c) { return /[",\n]/.test(c) ? '"' + String(c).replace(/"/g, '""') + '"' : c; }).join(',');
  }).join('\n');
}

/** Para testar no editor: clique em Executar com esta função e veja o resultado em "Registro de execução". */
function testar() {
  const t = Date.now();
  const csv = montarCsv();
  Logger.log(csv.split('\n').length - 1 + ' linhas em ' + Math.round((Date.now() - t) / 1000) + ' segundos');
  Logger.log(csv.split('\n').slice(0, 6).join('\n'));
}
