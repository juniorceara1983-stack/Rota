// ============================================================
// Sistema de Monitoramento de Rotas
// Backend – Google Apps Script (Code.gs)
// ============================================================

const SHEET_ROTAS    = 'Rotas';
const SHEET_GPS      = 'GPS_Log';
const SHEET_VEICULOS = 'Veiculos';
const SHEET_CONFIG   = 'Config';

// ── Roteamento de Páginas ────────────────────────────────────
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'driver';
  if (page === 'supervisor') {
    return HtmlService.createHtmlOutputFromFile('Supervisor')
      .setTitle('Painel do Supervisor – Sistema de Rotas')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createHtmlOutputFromFile('Driver')
    .setTitle('Painel do Motorista – Sistema de Rotas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ── Utilitário: obter ou criar aba ───────────────────────────
function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    switch (name) {
      case SHEET_ROTAS:
        sheet.appendRow([
          'ID', 'Nome_Motorista', 'Placa', 'Modelo', 'Obs_Inicial',
          'Hora_Inicio', 'Lat_Inicio', 'Long_Inicio',
          'Status', 'Obs_Viagem', 'Hora_Fim', 'Lat_Fim', 'Long_Fim'
        ]);
        sheet.setFrozenRows(1);
        break;
      case SHEET_GPS:
        sheet.appendRow(['ID_Rota', 'Timestamp', 'Latitude', 'Longitude', 'Velocidade']);
        sheet.setFrozenRows(1);
        break;
      case SHEET_VEICULOS:
        sheet.appendRow(['Placa', 'Modelo', 'Status']);
        sheet.setFrozenRows(1);
        break;
      case SHEET_CONFIG:
        sheet.appendRow(['Chave', 'Valor']);
        sheet.appendRow(['Intervalo_GPS', '10']);
        sheet.setFrozenRows(1);
        break;
    }
  }
  return sheet;
}

// ── Iniciar Rota ─────────────────────────────────────────────
function iniciarRota(data) {
  try {
    const sheet = getOrCreateSheet(SHEET_ROTAS);
    const id    = Utilities.getUuid();
    const ts    = data.timestamp || new Date().toISOString();
    sheet.appendRow([
      id,
      data.nome,
      (data.placa || '').toUpperCase(),
      data.modelo,
      data.obsInicial || '',
      ts,
      data.latitude  || '',
      data.longitude || '',
      'Em_Transito',
      '',
      '',
      '',
      ''
    ]);
    _atualizarStatusVeiculo((data.placa || '').toUpperCase(), 'Em_Transito');
    return { success: true, id: id, timestamp: ts };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── Registrar Ponto GPS ──────────────────────────────────────
function registrarPontoGPS(data) {
  try {
    const sheet = getOrCreateSheet(SHEET_GPS);
    sheet.appendRow([
      data.rotaId,
      data.timestamp || new Date().toISOString(),
      data.latitude,
      data.longitude,
      data.velocidade !== undefined ? data.velocidade : ''
    ]);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── Adicionar Observação de Viagem ───────────────────────────
function adicionarObservacao(data) {
  try {
    const sheet  = getOrCreateSheet(SHEET_ROTAS);
    const rows   = sheet.getDataRange().getValues();
    const ts     = data.timestamp || new Date().toISOString();
    const nova   = '[' + _formatTs(ts) + '] ' + data.observacao;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.rotaId) {
        const atual = rows[i][9] ? rows[i][9] + '\n' + nova : nova;
        sheet.getRange(i + 1, 10).setValue(atual);
        return { success: true };
      }
    }
    return { success: false, error: 'Rota não encontrada.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── Finalizar Rota ───────────────────────────────────────────
function finalizarRota(data) {
  try {
    const sheet = getOrCreateSheet(SHEET_ROTAS);
    const rows  = sheet.getDataRange().getValues();
    const ts    = data.timestamp || new Date().toISOString();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.rotaId) {
        sheet.getRange(i + 1,  9).setValue('Finalizado');
        sheet.getRange(i + 1, 11).setValue(ts);
        sheet.getRange(i + 1, 12).setValue(data.latitude  || '');
        sheet.getRange(i + 1, 13).setValue(data.longitude || '');
        _atualizarStatusVeiculo(rows[i][2], 'Disponivel');
        return { success: true, timestamp: ts };
      }
    }
    return { success: false, error: 'Rota não encontrada.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── Forçar Finalização (pelo supervisor) ─────────────────────
function forcarFinalizarRota(rotaId) {
  return finalizarRota({
    rotaId:    rotaId,
    latitude:  '',
    longitude: '',
    timestamp: new Date().toISOString()
  });
}

// ── Liberar Veículo ──────────────────────────────────────────
function liberarVeiculo(data) {
  try {
    const sheet = getOrCreateSheet(SHEET_VEICULOS);
    const rows  = sheet.getDataRange().getValues();
    const placa = (data.placa || '').toUpperCase();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === placa) {
        sheet.getRange(i + 1, 2).setValue(data.modelo);
        sheet.getRange(i + 1, 3).setValue('Disponivel');
        return { success: true, message: 'Veículo atualizado com sucesso.' };
      }
    }
    sheet.appendRow([placa, data.modelo, 'Disponivel']);
    return { success: true, message: 'Veículo liberado e registrado.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── Consultas do Supervisor ──────────────────────────────────
function getRotasAtivas() {
  try {
    const rows  = getOrCreateSheet(SHEET_ROTAS).getDataRange().getValues();
    const rotas = [];
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][8] === 'Em_Transito') {
        rotas.push({
          id:         rows[i][0],
          nome:       rows[i][1],
          placa:      rows[i][2],
          modelo:     rows[i][3],
          obsInicial: rows[i][4],
          horaInicio: String(rows[i][5]),
          latInicio:  rows[i][6],
          longInicio: rows[i][7],
          obsViagem:  rows[i][9]
        });
      }
    }
    return { success: true, rotas: rotas };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getVeiculos() {
  try {
    const rows     = getOrCreateSheet(SHEET_VEICULOS).getDataRange().getValues();
    const veiculos = [];
    for (let i = 1; i < rows.length; i++) {
      veiculos.push({ placa: rows[i][0], modelo: rows[i][1], status: rows[i][2] });
    }
    return { success: true, veiculos: veiculos };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getConfig() {
  try {
    const rows   = getOrCreateSheet(SHEET_CONFIG).getDataRange().getValues();
    const config = {};
    for (let i = 1; i < rows.length; i++) {
      config[rows[i][0]] = rows[i][1];
    }
    if (!config['Intervalo_GPS']) config['Intervalo_GPS'] = '10';
    return { success: true, config: config };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function setConfig(data) {
  try {
    const sheet = getOrCreateSheet(SHEET_CONFIG);
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === 'Intervalo_GPS') {
        sheet.getRange(i + 1, 2).setValue(String(data.intervaloGPS));
        return { success: true };
      }
    }
    sheet.appendRow(['Intervalo_GPS', String(data.intervaloGPS)]);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── Detalhes de uma Rota (com log GPS) ───────────────────────
function getDetalhesRota(rotaId) {
  try {
    const rotaRows = getOrCreateSheet(SHEET_ROTAS).getDataRange().getValues();
    const gpsRows  = getOrCreateSheet(SHEET_GPS).getDataRange().getValues();

    let rota = null;
    for (let i = 1; i < rotaRows.length; i++) {
      if (rotaRows[i][0] === rotaId) {
        rota = {
          id:         rotaRows[i][0],
          nome:       rotaRows[i][1],
          placa:      rotaRows[i][2],
          modelo:     rotaRows[i][3],
          obsInicial: rotaRows[i][4],
          horaInicio: String(rotaRows[i][5]),
          latInicio:  rotaRows[i][6],
          longInicio: rotaRows[i][7],
          status:     rotaRows[i][8],
          obsViagem:  rotaRows[i][9],
          horaFim:    String(rotaRows[i][10]),
          latFim:     rotaRows[i][11],
          longFim:    rotaRows[i][12]
        };
        break;
      }
    }
    if (!rota) return { success: false, error: 'Rota não encontrada.' };

    const gpsLog = [];
    for (let i = 1; i < gpsRows.length; i++) {
      if (gpsRows[i][0] === rotaId) {
        gpsLog.push({
          timestamp:  String(gpsRows[i][1]),
          latitude:   gpsRows[i][2],
          longitude:  gpsRows[i][3],
          velocidade: gpsRows[i][4]
        });
      }
    }
    return { success: true, rota: rota, gpsLog: gpsLog };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── Gerar Relatório CSV ──────────────────────────────────────
function gerarRelatorio(filtros) {
  try {
    const rotaRows = getOrCreateSheet(SHEET_ROTAS).getDataRange().getValues();
    const gpsRows  = getOrCreateSheet(SHEET_GPS).getDataRange().getValues();

    // Índice GPS por rotaId
    const gpsIndex = {};
    for (let i = 1; i < gpsRows.length; i++) {
      const rid = gpsRows[i][0];
      if (!gpsIndex[rid]) gpsIndex[rid] = [];
      gpsIndex[rid].push(gpsRows[i]);
    }

    const dataInicio = filtros.dataInicio ? filtros.dataInicio + 'T00:00:00' : null;
    const dataFim    = filtros.dataFim    ? filtros.dataFim    + 'T23:59:59' : null;
    const placa      = filtros.placa      ? filtros.placa.toUpperCase()       : null;

    // BOM UTF-8 para compatibilidade com Excel
    let csv = '\uFEFF';
    csv += '"ID","Nome_Motorista","Placa","Modelo","Obs_Inicial","Hora_Inicio",' +
           '"Lat_Inicio","Long_Inicio","Status","Obs_Viagem","Hora_Fim","Lat_Fim","Long_Fim"\n';

    const rotasFiltradas = [];
    for (let i = 1; i < rotaRows.length; i++) {
      const row = rotaRows[i];
      if (placa      && row[2] !== placa)                        continue;
      if (dataInicio && String(row[5]) < dataInicio)             continue;
      if (dataFim    && String(row[5]) > dataFim)                continue;
      rotasFiltradas.push(row);
      csv += row.map(function(c) {
        return '"' + String(c).replace(/"/g, '""') + '"';
      }).join(',') + '\n';
    }

    // Log GPS das rotas filtradas
    csv += '\n"=== LOG DE PONTOS GPS ==="\n';
    csv += '"ID_Rota","Timestamp","Latitude","Longitude","Velocidade"\n';
    for (let r = 0; r < rotasFiltradas.length; r++) {
      const pts = gpsIndex[rotasFiltradas[r][0]] || [];
      for (let p = 0; p < pts.length; p++) {
        csv += pts[p].map(function(c) {
          return '"' + String(c).replace(/"/g, '""') + '"';
        }).join(',') + '\n';
      }
    }

    return { success: true, csv: csv, total: rotasFiltradas.length };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── API REST para GitHub Pages ───────────────────────────────
// Recebe chamadas fetch() feitas a partir das páginas no GitHub Pages.
// O corpo deve ser JSON com o campo "action" indicando a operação.
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action;
    var result;

    switch (action) {
      case 'iniciarRota':
        result = iniciarRota(payload);
        break;
      case 'registrarPontoGPS':
        result = registrarPontoGPS(payload);
        break;
      case 'adicionarObservacao':
        result = adicionarObservacao(payload);
        break;
      case 'finalizarRota':
        result = finalizarRota(payload);
        break;
      case 'forcarFinalizarRota':
        result = forcarFinalizarRota(payload.rotaId);
        break;
      case 'liberarVeiculo':
        result = liberarVeiculo(payload);
        break;
      case 'getRotasAtivas':
        result = getRotasAtivas();
        break;
      case 'getVeiculos':
        result = getVeiculos();
        break;
      case 'getConfig':
        result = getConfig();
        break;
      case 'setConfig':
        result = setConfig(payload);
        break;
      case 'getDetalhesRota':
        result = getDetalhesRota(payload.rotaId);
        break;
      case 'gerarRelatorio':
        result = gerarRelatorio(payload);
        break;
      default:
        result = { success: false, error: 'Ação desconhecida: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Funções privadas ─────────────────────────────────────────
function _atualizarStatusVeiculo(placa, status) {
  const sheet = getOrCreateSheet(SHEET_VEICULOS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === placa) {
      sheet.getRange(i + 1, 3).setValue(status);
      return;
    }
  }
}

function _formatTs(iso) {
  try {
    const d = new Date(iso);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  } catch (_) {
    return iso;
  }
}
