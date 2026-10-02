/**
 * ESTUDO DE CASO – PONTOS DE DESCARTE IRREGULAR
 * GCM APGO – Inspetoria de Inteligência
 *
 * Script vinculado à Planilha Google. Ele:
 *  - monta as abas (Visitas, Consolidação, Indicadores, Pontos, Escala);
 *  - publica o aplicativo de campo (Index.html) como aplicativo da Web;
 *  - recebe cada visita, salva as fotos no Drive e grava a linha na aba Visitas;
 *  - exporta uma cópia da planilha em Excel (.xlsx) para o Drive.
 */

/* ===================== CONFIGURAÇÃO ===================== */
const CFG = {
  INICIO: '2026-10-05',          // 1º dia da pesquisa (SEG)
  DIAS_CAMPO: 28,                // 4 semanas de visitas
  DIAS_TOTAL: 32,                // até 05/11 (consolidação)
  EQUIPES: ['Charlie', 'Delta', 'Alfa', 'Bravo'],
  DIAS_SEMANA: ['SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SÁB', 'DOM'],
  GRUPOS: [[1, 6], [7, 12], [13, 18], [19, 24], [25, 30], [31, 36], [37, 41]],
  PASTA_FOTOS: 'Estudo Descarte Irregular – Fotos',
  PASTA_EXCEL: 'Estudo Descarte Irregular – Exportações Excel',
  PIN_PADRAO: '2026',
  TZ: 'America/Sao_Paulo'
};

// [Nº, Endereço, Situação (documento SDU), Grau de necessidade, Regional]
const PONTOS = [
  [1, 'RUA C-18 – JARDIM CASCATA', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [2, 'AVENIDA DOS IMIGRANTES – VILA DELFIORE', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [3, 'RUA PORTUGAL COM AVENIDA DOS IMIGRANTES ENTRE SETORES: MADRE GERMANA E QUINTA DA BOA VISTA', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [4, 'AVENIDA JUSCELINO KUBITSCHEK – SETOR MADRE GERMANA I', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [5, 'RUA JI-26 – JARDIM DOS IPÊS', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [6, 'RUA JI-12 – JARDIM DOS IPÊS', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [7, 'AVENIDA DOM CRUZ – JARDIM DOM BOSCO/VILA IZAURA', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [8, 'RUA SAQUAREMA QD. 10 LT. 10 – JARDIM HIMALAIA', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [9, 'RUA BOMSUCESSO – SETOR BURITI SERENO', 'SEM MONITORAMENTO', '', 'Buriti Sereno*'],
  [10, 'AVENIDA CAPYABA QD. 84 LT. 7 – JARDIM HELVÉCIA', 'MONITORADO', '', 'A confirmar'],
  [11, 'RUA AMAZONAS QD. 24 – ST NOSSA SENHORA DE LOURDES', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [12, 'AVENIDA W-2 QD 31 – SETOR SANTA LUZIA', 'SEM MONITORAMENTO', '', 'Santa Luzia*'],
  [13, 'RUA 10 QD. 33 – SETOR SANTA LUZIA', 'SEM MONITORAMENTO', '', 'Santa Luzia*'],
  [14, 'TRAVESSA SANTO ANTÔNIO – ST BELA VISTA', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [15, 'AVENIDA W-6 SETOR SANTA LUZIA', 'SEM MONITORAMENTO', '', 'Santa Luzia*'],
  [16, 'AVENIDA 12 DE OUTUBRO, 100 – PARQUE FLAMBOYANT', 'SEM MONITORAMENTO', 'LIMPEZA REALIZADA NA DATA DE 08/09', 'A confirmar'],
  [17, 'ALAMEDA B – CHÁCARA SÃO PEDRO', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [18, 'RUA LT 01 – PRÓXIMO AVENIDA JATAÍ', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [19, 'AVENIDA 15 DE NOVEMBRO – JARDIM MONTE CRISTO', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [20, 'RUA CRISTIANO GALDINO GONÇALVES – INDEPENDÊNCIA MANSÕES', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [21, 'RUA 59 QD 192 LT 01 – INDEPENDÊNCIA MANSÕES', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [22, 'AVENIDA MARIA LUIZA DAS DORES – INDEPENDÊNCIA MANSÕES', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [23, 'AVENIDA DA LUZ – GOIÂNIA PARQUE SUL (03 RUAS)', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [24, 'AVENIDA EUGÊNIO ELIAS DE DEUS – INDEPENDÊNCIA MANSÕES', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [25, 'AVENIDA MIRANDA QD 45 – GOIÂNIA PARQUE SUL', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [26, 'RUA MARIA JACOB CHAER DE SOUZA – RESIDENCIAL ARAGUAIA', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [27, 'ALAMEDA BICUDO QD. 75 – RESIDENCIAL NORTE SUL', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [28, 'ALAMEDA FLORESTA – JARDIM TROPICAL', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [29, 'RUA 14-E – SETOR GARAVELO', 'SEM MONITORAMENTO', '', 'Garavelo*'],
  [30, 'ALAMEDA RUI RODRIGUES QD 20 – ST COLONIAL SUL', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [31, 'RUA 12 QD 101 – BAIXADA BLACK JD TIRADENTES', 'MONITORADO', '', 'A confirmar'],
  [32, 'RUA 35 ENTRE OS SETORES JD TIRADENTES E NOVA CIDADE', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [33, 'AVENIDA SÃO PAULO – ENTRE OS SETORES NOVA CIDADE E JD FLORENÇA', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [34, 'RUA BERGAMO QD 23 – JARDIM FLORENÇA', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [35, 'RUA OUTRO QD 01 – SETOR FABRÍCIO', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [36, 'RUA ÔNIX – SETOR COMENDADOR WALMOR', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [37, 'RUA DAS PERDIZES QD 41 – SETOR COLINA AZUL', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [38, 'AVENIDA SÃO FRANCISCO – PARQUE HAYALA', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [39, 'PRIMEIRA AVENIDA QD. 1 B – BAIRRO CARDOSO', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [40, 'RUA DOS CARIOCAS – VILA MARIANA', 'SEM MONITORAMENTO', '', 'A confirmar'],
  [41, 'AVENIDA PRADO JÚNIOR – BURITI SERENO', 'SEM MONITORAMENTO', '', 'Buriti Sereno*']
];

const ABA = {
  VISITAS: 'Visitas',
  CONSOL: 'Consolidação',
  IND: 'Indicadores',
  PONTOS: 'Pontos',
  ESCALA: 'Escala'
};

const CAB_VISITAS = [
  'Carimbo de envio', 'Data da visita', 'Hora', 'Semana', 'Dia', 'Equipe', 'Agentes / matrícula',
  'Nº ponto', 'Endereço', 'Regional', 'Latitude', 'Longitude', 'Precisão GPS (m)',
  'Existe lixo?', 'Tipo de área', 'Tipo de resíduo', 'Volume (0-3)', 'Comparação c/ visita anterior',
  'Limpeza pela Prefeitura?', 'Data da limpeza', 'Flagrante de descarte?', 'Apreensão de veículo?',
  'Placa / tipo de veículo', 'Nº RAI / BO', 'Condição do tempo', 'Nº de fotos', 'Fotos (Drive)',
  'Observações', 'ID do envio'
];
// índices (0-based) usados no código
const C = {
  SEMANA: 3, PONTO: 7, VOLUME: 16, DATA: 1, EQUIPE: 5, ID: 28
};

const COR = { navy: '#1F3864', claro: '#D9E2F3', cinza: '#F2F2F2' };
const COR_EQUIPE = { Charlie: '#DDEBF7', Delta: '#E2EFDA', Alfa: '#FFF2CC', Bravo: '#FCE4D6' };

/* ===================== MENU ===================== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Estudo Descarte')
    .addItem('1. Configurar planilha (primeira vez)', 'configurarPlanilha')
    .addSeparator()
    .addItem('Exportar cópia em Excel (.xlsx) para o Drive', 'exportarExcel')
    .addItem('Abrir pasta de fotos', 'abrirPastaFotos')
    .addItem('Alterar PIN de acesso do aplicativo', 'alterarPin')
    .addToUi();
}

/* ===================== APLICATIVO DA WEB ===================== */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('GCM APGO – Descarte Irregular')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Chamadas do aplicativo quando ele é aberto fora do Google (GitHub Pages): {fn, args} em JSON. */
function doPost(e) {
  const API = { getConfig: getConfig, getStatus: getStatus, enviarVisita: enviarVisita };
  let saida;
  try {
    const req = JSON.parse(e.postData.contents);
    if (!API[req.fn]) throw new Error('Função inválida.');
    saida = { r: API[req.fn].apply(null, req.args || []) };
  } catch (err) {
    saida = { erro: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(saida)).setMimeType(ContentService.MimeType.JSON);
}

function pinValido_(pin) {
  const atual = PropertiesService.getScriptProperties().getProperty('PIN') || CFG.PIN_PADRAO;
  return String(pin || '').trim() === String(atual).trim();
}

/** Dados iniciais para o aplicativo. */
function getConfig(pin) {
  if (!pinValido_(pin)) return { ok: false, msg: 'PIN incorreto.' };
  return {
    ok: true,
    inicio: CFG.INICIO,
    diasCampo: CFG.DIAS_CAMPO,
    diasTotal: CFG.DIAS_TOTAL,
    equipes: CFG.EQUIPES,
    diasSemana: CFG.DIAS_SEMANA,
    grupos: CFG.GRUPOS,
    pontos: PONTOS.map(p => ({ n: p[0], end: p[1], sit: p[2], reg: p[4] })),
    status: getStatus_()
  };
}

function getStatus(pin) {
  if (!pinValido_(pin)) return { ok: false, msg: 'PIN incorreto.' };
  return { ok: true, status: getStatus_() };
}

/** Para cada ponto: lista das visitas já registradas {s: semana, v: volume, d: data, e: equipe}. */
function getStatus_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(ABA.VISITAS);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  const dados = sh.getRange(2, 1, sh.getLastRow() - 1, CAB_VISITAS.length).getValues();
  dados.forEach(r => {
    const n = Number(r[C.PONTO]);
    if (!n) return;
    (out[n] = out[n] || []).push({
      s: Number(r[C.SEMANA]),
      v: r[C.VOLUME] === '' ? null : Number(r[C.VOLUME]),
      d: r[C.DATA] instanceof Date ? Utilities.formatDate(r[C.DATA], CFG.TZ, 'dd/MM/yyyy') : String(r[C.DATA]),
      e: String(r[C.EQUIPE])
    });
  });
  return out;
}

/** Recebe uma visita do aplicativo de campo. */
function enviarVisita(pin, v) {
  if (!pinValido_(pin)) return { ok: false, msg: 'PIN incorreto. Entre novamente no aplicativo.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(ABA.VISITAS);
    if (!sh) return { ok: false, msg: 'A planilha ainda não foi configurada.' };

    const ponto = PONTOS.find(p => p[0] === Number(v.ponto));
    if (!ponto) return { ok: false, msg: 'Ponto inválido.' };
    const semana = Number(v.semana);
    const volume = Number(v.volume);

    // evita duplicar quando o celular reenvia o mesmo registro
    if (sh.getLastRow() > 1) {
      const ids = sh.getRange(2, C.ID + 1, sh.getLastRow() - 1, 1).getValues().flat();
      if (ids.indexOf(v.id) !== -1) return { ok: true, duplicado: true };
    }

    // comparação automática com a visita anterior do mesmo ponto
    let comparacao = '1ª visita';
    const anteriores = (getStatus_()[ponto[0]] || []).filter(x => x.s < semana && x.v !== null);
    if (anteriores.length) {
      anteriores.sort((a, b) => a.s - b.s);
      const ant = anteriores[anteriores.length - 1].v;
      if (ant > 0 && volume === 0) comparacao = 'Eliminado';
      else if (volume > ant) comparacao = 'Aumentou';
      else if (volume < ant) comparacao = 'Diminuiu';
      else comparacao = 'Manteve';
    }

    // fotos
    const fotos = v.fotos || [];
    let linkFotos = '';
    if (fotos.length) {
      const pastaVisita = pastaDaVisita_(ponto, semana, v.data);
      const pref = 'P' + pad2_(ponto[0]) + '_S' + semana + '_' + v.data + '_' + v.equipe + '_';
      fotos.forEach((b64, i) => {
        const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', pref + (i + 1) + '.jpg');
        pastaVisita.createFile(blob);
      });
      linkFotos = '=HYPERLINK("' + pastaVisita.getUrl() + '","Abrir ' + fotos.length + ' foto(s)")';
    }

    const dt = v.data.split('-'); // yyyy-mm-dd
    const linha = [
      new Date(),
      new Date(Number(dt[0]), Number(dt[1]) - 1, Number(dt[2])),
      v.hora,
      semana,
      v.dia,
      v.equipe,
      v.agentes,
      ponto[0],
      ponto[1],
      ponto[4],
      v.lat === '' ? '' : Number(v.lat),
      v.lng === '' ? '' : Number(v.lng),
      v.precisao === '' ? '' : Number(v.precisao),
      v.lixo,
      v.area,
      (v.residuos || []).join(', '),
      volume,
      comparacao,
      v.limpeza,
      v.dataLimpeza || '',
      v.flagrante,
      v.apreensao,
      v.placa || '',
      v.rai || '',
      v.tempo,
      fotos.length,
      linkFotos,
      v.obs || '',
      v.id
    ];
    sh.appendRow(linha);
    return { ok: true, comparacao: comparacao };
  } finally {
    lock.releaseLock();
  }
}

/* ===================== DRIVE ===================== */
function pastaRaiz_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PASTA_FOTOS_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* recria abaixo */ }
  }
  const arquivo = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  const pais = arquivo.getParents();
  const base = pais.hasNext() ? pais.next() : DriveApp.getRootFolder();
  const pasta = base.createFolder(CFG.PASTA_FOTOS);
  props.setProperty('PASTA_FOTOS_ID', pasta.getId());
  return pasta;
}

function subpasta_(pai, nome) {
  const it = pai.getFoldersByName(nome);
  return it.hasNext() ? it.next() : pai.createFolder(nome);
}

function pastaDaVisita_(ponto, semana, dataISO) {
  const nomePonto = 'Ponto ' + pad2_(ponto[0]) + ' – ' + ponto[1].substring(0, 60);
  const p = subpasta_(pastaRaiz_(), nomePonto);
  const d = dataISO.split('-');
  return subpasta_(p, 'Semana ' + semana + ' – ' + d[2] + '-' + d[1] + '-' + d[0]);
}

function abrirPastaFotos() {
  const url = pastaRaiz_().getUrl();
  const html = HtmlService.createHtmlOutput(
    '<p style="font-family:Arial">Pasta de fotos: <a href="' + url + '" target="_blank">abrir no Drive</a></p>'
  ).setWidth(360).setHeight(90);
  SpreadsheetApp.getUi().showModalDialog(html, 'Fotos do estudo');
}

function exportarExcel() {
  const ss = SpreadsheetApp.getActive();
  SpreadsheetApp.flush();
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
  const resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
  const nome = 'Estudo Descarte Irregular – ' + Utilities.formatDate(new Date(), CFG.TZ, 'dd-MM-yyyy HH\'h\'mm') + '.xlsx';
  const arquivo = DriveApp.getFileById(ss.getId());
  const pais = arquivo.getParents();
  const base = pais.hasNext() ? pais.next() : DriveApp.getRootFolder();
  const destino = subpasta_(base, CFG.PASTA_EXCEL);
  const f = destino.createFile(resp.getBlob().setName(nome));
  const html = HtmlService.createHtmlOutput(
    '<p style="font-family:Arial">Cópia em Excel criada: <a href="' + f.getUrl() + '" target="_blank">' + nome + '</a></p>'
  ).setWidth(420).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Exportação concluída');
}

function alterarPin() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('PIN de acesso', 'Digite o novo PIN que os agentes usarão no aplicativo:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const novo = r.getResponseText().trim();
  if (novo.length < 4) { ui.alert('Use pelo menos 4 caracteres.'); return; }
  PropertiesService.getScriptProperties().setProperty('PIN', novo);
  ui.alert('PIN alterado. Informe o novo PIN às equipes.');
}

/* ===================== MONTAGEM DA PLANILHA ===================== */
function configurarPlanilha() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(CFG.TZ);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PIN')) props.setProperty('PIN', CFG.PIN_PADRAO);

  montarVisitas_(ss);
  montarPontos_(ss);
  montarEscala_(ss);
  montarConsolidacao_(ss);
  montarIndicadores_(ss);

  // ordem das abas
  [ABA.IND, ABA.CONSOL, ABA.VISITAS, ABA.ESCALA, ABA.PONTOS].forEach((n, i) => {
    const sh = ss.getSheetByName(n);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
  });
  const padrao = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (padrao && ss.getSheets().length > 1) ss.deleteSheet(padrao);
  pastaRaiz_();
  ss.setActiveSheet(ss.getSheetByName(ABA.IND));

  SpreadsheetApp.getUi().alert(
    'Planilha configurada!\n\nPIN de acesso do aplicativo: ' + props.getProperty('PIN') +
    '\n\nPróximo passo: Implantar > Nova implantação > Aplicativo da Web.'
  );
}

function aba_(ss, nome) {
  return ss.getSheetByName(nome) || ss.insertSheet(nome);
}

function cabecalho_(sh, valores) {
  const r = sh.getRange(1, 1, 1, valores.length);
  r.setValues([valores]).setBackground(COR.navy).setFontColor('#FFFFFF').setFontWeight('bold')
    .setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true);
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 42);
}

function montarVisitas_(ss) {
  const sh = aba_(ss, ABA.VISITAS);
  cabecalho_(sh, CAB_VISITAS); // preserva os dados já lançados
  const larg = [130, 95, 60, 65, 50, 75, 170, 60, 300, 110, 90, 90, 80, 75, 110, 200, 80, 120, 105, 95, 90, 95, 140, 110, 110, 70, 120, 280, 150];
  larg.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('A2:A').setNumberFormat('dd/MM/yyyy HH:mm');
  sh.getRange('B2:B').setNumberFormat('dd/MM/yyyy');
  sh.getRange('K2:L').setNumberFormat('0.000000');
  sh.setFrozenColumns(0);
  const regras = Object.keys(COR_EQUIPE).map(e =>
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(e).setBackground(COR_EQUIPE[e])
      .setRanges([sh.getRange('F2:F')]).build());
  regras.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Sim').setBackground('#F8CBAD')
    .setRanges([sh.getRange('V2:V')]).build());
  sh.setConditionalFormatRules(regras);
}

function montarPontos_(ss) {
  const sh = aba_(ss, ABA.PONTOS);
  sh.clear();
  cabecalho_(sh, ['Nº', 'Endereço', 'Situação (documento SDU)', 'Grau de necessidade', 'Regional', 'Dia da visita']);
  const linhas = PONTOS.map(p => {
    const g = CFG.GRUPOS.findIndex(x => p[0] >= x[0] && p[0] <= x[1]);
    return [p[0], p[1], p[2], p[3], p[4], CFG.DIAS_SEMANA[g]];
  });
  sh.getRange(2, 1, linhas.length, 6).setValues(linhas).setWrap(true).setVerticalAlignment('middle');
  [45, 420, 160, 170, 120, 95].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(2, 1, linhas.length, 1).setHorizontalAlignment('center');
  sh.getRange(2, 6, linhas.length, 1).setHorizontalAlignment('center').setFontWeight('bold');
  sh.getRange(1, 1, linhas.length + 1, 6).setBorder(true, true, true, true, true, true, '#8EA9DB', SpreadsheetApp.BorderStyle.SOLID);
}

function montarEscala_(ss) {
  const sh = aba_(ss, ABA.ESCALA);
  sh.clear();
  cabecalho_(sh, ['Data', 'Dia', 'Equipe', 'Fase', 'Pontos', 'Visitas previstas', 'Visitas realizadas']);
  const ini = isoParaData_(CFG.INICIO);
  const linhas = [];
  for (let k = 0; k < CFG.DIAS_TOTAL; k++) {
    const d = new Date(ini.getTime()); d.setDate(d.getDate() + k);
    const dia = CFG.DIAS_SEMANA[(d.getDay() + 6) % 7];
    const eq = CFG.EQUIPES[k % 4];
    if (k < CFG.DIAS_CAMPO) {
      const g = CFG.GRUPOS[(d.getDay() + 6) % 7];
      linhas.push([d, dia, eq, 'Semana ' + (Math.floor(k / 7) + 1), g[0] + ' a ' + g[1], g[1] - g[0] + 1]);
    } else {
      linhas.push([d, dia, eq, 'Consolidação', 'Tabulação, análise e relatório', '']);
    }
  }
  sh.getRange(2, 1, linhas.length, 6).setValues(linhas);
  const f = linhas.map((l, i) => [l[5] === '' ? '' : '=COUNTIF(' + ABA.VISITAS + '!$B:$B,$A' + (i + 2) + ')']);
  sh.getRange(2, 7, f.length, 1).setFormulas(f);
  sh.getRange('A2:A').setNumberFormat('dd/MM/yyyy');
  [95, 55, 90, 110, 220, 120, 130].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(2, 1, linhas.length, 7).setHorizontalAlignment('center');
  sh.getRange(2, 2, linhas.length, 1).setFontWeight('bold');
  const regras = Object.keys(COR_EQUIPE).map(e =>
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(e).setBackground(COR_EQUIPE[e])
      .setRanges([sh.getRange('C2:C')]).build());
  sh.setConditionalFormatRules(regras);
  sh.getRange(1, 1, linhas.length + 1, 7).setBorder(true, true, true, true, true, true, '#8EA9DB', SpreadsheetApp.BorderStyle.SOLID);
}

function montarConsolidacao_(ss) {
  const sh = aba_(ss, ABA.CONSOL);
  sh.clear();
  cabecalho_(sh, ['Nº', 'Endereço', 'Regional', 'Dia', 'Área (último registro)', 'S1', 'S2', 'S3', 'S4',
    'Visitas realizadas', 'Limpeza Prefeitura', 'Apreensão de veículo', 'Tendência']);
  const V = ABA.VISITAS + '!';
  const linhas = [];
  const formulas = [];
  PONTOS.forEach((p, i) => {
    const r = i + 2;
    const g = CFG.GRUPOS.findIndex(x => p[0] >= x[0] && p[0] <= x[1]);
    linhas.push([p[0], p[1], p[4], CFG.DIAS_SEMANA[g]]);
    const semana = w => '=IFERROR(INDEX(FILTER(' + V + '$Q:$Q;' + V + '$H:$H=$A' + r + ';' + V + '$D:$D=' + w + ');COUNTIFS(' + V + '$H:$H;$A' + r + ';' + V + '$D:$D;' + w + '));"")';
    formulas.push([
      '=IFERROR(INDEX(FILTER(' + V + '$O:$O;' + V + '$H:$H=$A' + r + ');COUNTIF(' + V + '$H:$H;$A' + r + '));"")',
      semana(1), semana(2), semana(3), semana(4),
      '=COUNTIF(' + V + '$H:$H;$A' + r + ')',
      '=IF(COUNTIFS(' + V + '$H:$H;$A' + r + ';' + V + '$S:$S;"Sim")>0;"Sim";"Não")',
      '=IF(COUNTIFS(' + V + '$H:$H;$A' + r + ';' + V + '$V:$V;"Sim")>0;"Sim";"Não")',
      '=IF(OR(F' + r + '="";I' + r + '="");IF(J' + r + '=0;"SEM DADOS";"EM ANDAMENTO");' +
        'IF(F' + r + '=0;IF(MAX(G' + r + ':I' + r + ')>0;"NOVO FOCO";"SEM DESCARTE");' +
        'IF(AND(H' + r + '<>"";H' + r + '=0;I' + r + '=0);"ELIMINADO";' +
        'IF(I' + r + '<F' + r + ';"DIMINUIU";IF(I' + r + '=F' + r + ';"MANTEVE";"AUMENTOU")))))'
    ]);
  });
  sh.getRange(2, 1, linhas.length, 4).setValues(linhas);
  sh.getRange(2, 5, formulas.length, 9).setFormulas(formulas.map(f => f.map(converterFormula_)));
  [45, 380, 110, 50, 130, 45, 45, 45, 45, 85, 95, 100, 125].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(2, 2, linhas.length, 1).setWrap(true);
  sh.getRange(2, 1, linhas.length, 1).setHorizontalAlignment('center');
  sh.getRange(2, 3, linhas.length, 11).setHorizontalAlignment('center');
  sh.getRange(2, 13, linhas.length, 1).setFontWeight('bold');
  sh.getRange(1, 1, linhas.length + 1, 13).setBorder(true, true, true, true, true, true, '#8EA9DB', SpreadsheetApp.BorderStyle.SOLID);
  const M = sh.getRange('M2:M42');
  const cores = { 'ELIMINADO': '#C6EFCE', 'DIMINUIU': '#E2EFDA', 'MANTEVE': '#FFF2CC', 'AUMENTOU': '#F8CBAD', 'NOVO FOCO': '#F4B183', 'SEM DESCARTE': '#DDEBF7', 'EM ANDAMENTO': '#EDEDED' };
  const regras = Object.keys(cores).map(t => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t).setBackground(cores[t]).setRanges([M]).build());
  const vols = sh.getRange('F2:I42');
  regras.push(SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(3).setBackground('#F8CBAD').setRanges([vols]).build());
  regras.push(SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(2).setBackground('#FFE699').setRanges([vols]).build());
  regras.push(SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(1).setBackground('#FFF2CC').setRanges([vols]).build());
  regras.push(SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setBackground('#C6EFCE').setRanges([vols]).build());
  sh.setConditionalFormatRules(regras);
  sh.setFrozenColumns(2);
}

function montarIndicadores_(ss) {
  const sh = aba_(ss, ABA.IND);
  sh.clear();
  sh.getRange('A1').setValue('ESTUDO DE CASO – PONTOS DE DESCARTE IRREGULAR • GCM APGO – Inspetoria de Inteligência')
    .setFontSize(13).setFontWeight('bold').setFontColor(COR.navy);
  sh.getRange('A2').setValue('Período: 05/10/2026 a 05/11/2026 • 41 pontos • 4 visitas por ponto • atualização automática')
    .setFontColor('#595959');

  const CS = "'" + ABA.CONSOL + "'!";
  const V = ABA.VISITAS + '!';
  let linha = 4;
  const bloco = (titulo, cab, linhas) => {
    sh.getRange(linha, 1).setValue(titulo).setFontWeight('bold').setFontColor(COR.navy).setFontSize(11);
    linha++;
    const hc = sh.getRange(linha, 1, 1, cab.length);
    hc.setValues([cab]).setBackground(COR.navy).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
    linha++;
    const corpo = sh.getRange(linha, 1, linhas.length, cab.length);
    corpo.setValues(linhas.map(l => l.map(c => (typeof c === 'string' && c.charAt(0) === '=') ? converterFormula_(c) : c)));
    sh.getRange(linha - 1, 1, linhas.length + 1, cab.length)
      .setBorder(true, true, true, true, true, true, '#8EA9DB', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(linha, 2, linhas.length, cab.length - 1).setHorizontalAlignment('center');
    const ini = linha;
    linha += linhas.length + 1;
    return ini;
  };

  const r0 = bloco('Andamento geral', ['Indicador', 'Valor', 'Meta', '%'], [
    ['Visitas realizadas', '=COUNTA(' + V + 'H2:H)', 164, ''],
    ['Pontos com ao menos 1 visita', '=COUNTIF(' + CS + 'J2:J42;">0")', 41, ''],
    ['Pontos com as 4 visitas', '=COUNTIF(' + CS + 'J2:J42;">=4")', 41, '']
  ]);
  for (let i = 0; i < 3; i++) {
    sh.getRange(r0 + i, 4).setFormula('=IFERROR(B' + (r0 + i) + '/C' + (r0 + i) + ',0)').setNumberFormat('0%');
  }

  const tend = ['ELIMINADO', 'DIMINUIU', 'MANTEVE', 'AUMENTOU', 'NOVO FOCO', 'SEM DESCARTE', 'EM ANDAMENTO', 'SEM DADOS'];
  const r1 = bloco('Tendência dos pontos', ['Tendência', 'Nº de pontos', '% dos 41'],
    tend.map(t => [t, '=COUNTIF(' + CS + 'M2:M42;"' + t + '")', '']));
  tend.forEach((t, i) => sh.getRange(r1 + i, 3).setFormula('=B' + (r1 + i) + '/41').setNumberFormat('0%'));

  bloco('Tipo de área (último registro de cada ponto)', ['Área', 'Nº de pontos'], [
    ['Pública', '=COUNTIF(' + CS + 'E2:E42;"Pública")'],
    ['Particular', '=COUNTIF(' + CS + 'E2:E42;"Particular")'],
    ['Não identificada', '=COUNTIF(' + CS + 'E2:E42;"Não identificada")']
  ]);

  bloco('Limpeza, flagrantes e apreensões', ['Indicador', 'Quantidade'], [
    ['Pontos com limpeza da Prefeitura no período', '=COUNTIF(' + CS + 'K2:K42;"Sim")'],
    ['Registros de limpeza informados nas visitas', '=COUNTIF(' + V + 'S2:S;"Sim")'],
    ['Flagrantes de descarte', '=COUNTIF(' + V + 'U2:U;"Sim")'],
    ['Apreensões de veículos', '=COUNTIF(' + V + 'V2:V;"Sim")'],
    ['Pontos com apreensão', '=COUNTIF(' + CS + 'L2:L42;"Sim")']
  ]);

  const r4 = bloco('Visitas por equipe', ['Equipe', 'Realizadas', 'Previstas', '%'],
    CFG.EQUIPES.map(e => [e, '=COUNTIF(' + V + 'F2:F;"' + e + '")', 41, '']));
  CFG.EQUIPES.forEach((e, i) => {
    sh.getRange(r4 + i, 4).setFormula('=IFERROR(B' + (r4 + i) + '/C' + (r4 + i) + ',0)').setNumberFormat('0%');
    sh.getRange(r4 + i, 1).setBackground(COR_EQUIPE[e]).setFontWeight('bold');
  });

  const r5 = bloco('Visitas por semana', ['Semana', 'Realizadas', 'Previstas', '%'],
    [1, 2, 3, 4].map(w => ['Semana ' + w, '=COUNTIF(' + V + 'D2:D;' + w + ')', 41, '']));
  for (let i = 0; i < 4; i++) {
    sh.getRange(r5 + i, 4).setFormula('=IFERROR(B' + (r5 + i) + '/C' + (r5 + i) + ',0)').setNumberFormat('0%');
  }

  bloco('Pontos monitorados (documento SDU)', ['Ponto', 'Tendência'], [
    ['10 – Av. Capyaba (Jd. Helvécia)', '=' + CS + 'M11'],
    ['31 – Rua 12, Baixada Black (Jd. Tiradentes)', '=' + CS + 'M32']
  ]);

  sh.getRange(linha, 1).setValue('Prioridade para instalação de câmeras (pontos que mantiveram, aumentaram ou surgiram – maior volume primeiro)')
    .setFontWeight('bold').setFontColor(COR.navy).setFontSize(11);
  linha++;
  sh.getRange(linha, 1).setFormula(converterFormula_(
    '=IFERROR(QUERY(' + CS + 'A1:M42;"select A, B, I, M where M = \'AUMENTOU\' or M = \'MANTEVE\' or M = \'NOVO FOCO\' order by I desc";1);"Sem dados suficientes ainda")'));

  [330, 380, 90, 80].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setHiddenGridlines(true);
}

/* ===================== UTILITÁRIOS ===================== */
// As fórmulas são escritas com ";" e convertidas para "," (padrão aceito pelo Apps Script).
function converterFormula_(f) {
  let out = '', aspas = false;
  for (let i = 0; i < f.length; i++) {
    const ch = f.charAt(i);
    if (ch === '"') aspas = !aspas;
    out += (!aspas && ch === ';') ? ',' : ch;
  }
  return out;
}

function pad2_(n) { return (n < 10 ? '0' : '') + n; }

function isoParaData_(iso) {
  const p = iso.split('-');
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}
