// ============================================================
// Sistema de Monitoramento de Rotas
// Backend – Google Apps Script (Code.gs)
// ============================================================

const SHEET_ROTAS    = 'Rotas';
const SHEET_GPS      = 'GPS_Log';
const SHEET_VEICULOS = 'Veiculos';
const SHEET_MOTORISTAS = 'Motoristas';
const SHEET_CONFIG   = 'Config';
const SHEET_ABAST    = 'Abastecimentos';
const SHEET_CHECKLIST = 'Checklist_Avarias';
const CHECKLIST_STATUS_ABERTO = 'Aberto';
const ABAST_EXPECTED_COLS = 6;
// Colunas totais da aba Rotas após inclusão de Km_Inicio e Km_Fim.
const ROTAS_EXPECTED_COLS = 15;

// ID da planilha – necessário quando o script é chamado via doPost (web app standalone)
const SPREADSHEET_ID = '1qu5ZkDwNnCnsFM_gKjDwbfAOGgdtb0XlOAc-_YqS7dQ';

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
  const ss = SpreadsheetApp.getActiveSpreadsheet() ||
             SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    switch (name) {
      case SHEET_ROTAS:
        sheet.appendRow([
          'ID', 'Nome_Motorista', 'Placa', 'Modelo', 'Obs_Inicial',
          'Hora_Inicio', 'Lat_Inicio', 'Long_Inicio',
          'Status', 'Obs_Viagem', 'Hora_Fim', 'Lat_Fim', 'Long_Fim',
          'Km_Inicio', 'Km_Fim'
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
      case SHEET_MOTORISTAS:
        sheet.appendRow(['Nome']);
        sheet.setFrozenRows(1);
        break;
      case SHEET_CONFIG:
        sheet.appendRow(['Chave', 'Valor']);
        sheet.appendRow(['Intervalo_GPS', '10']);
        sheet.appendRow(['Senha_Supervisor', '123456']);
        sheet.setFrozenRows(1);
        break;
      case SHEET_ABAST:
        sheet.appendRow(['Placa', 'Quilometragem', 'Litros', 'Valor', 'Data_Hora', 'Media_Km_L']);
        sheet.setFrozenRows(1);
        break;
      case SHEET_CHECKLIST:
        sheet.appendRow([
          'ID', 'Placa', 'Parte_Carro', 'Descricao_Avaria', 'Status',
          'Registrado_Em', 'Registrado_Por', 'Ciente_Motorista', 'Ciente_Em'
        ]);
        sheet.setFrozenRows(1);
        break;
    }
  } else if (name === SHEET_ROTAS) {
    _ensureRotasColumns(sheet);
  } else if (name === SHEET_ABAST) {
    _ensureAbastecimentosColumns(sheet);
  }
  return sheet;
}

// ── Iniciar Rota ─────────────────────────────────────────────
function iniciarRota(data) {
  try {
    const sheet = getOrCreateSheet(SHEET_ROTAS);
    _ensureRotasColumns(sheet);
    const rows  = sheet.getDataRange().getValues();
    const nomeMotorista = _normalizarNome(data.nome);
    if (!nomeMotorista) return { success: false, error: 'Nome do motorista não informado.' };
    const motorista = _buscarMotoristaPorNome(nomeMotorista);
    if (!motorista) {
      return { success: false, error: 'Motorista não cadastrado. Solicite cadastro ao supervisor.' };
    }
    for (let i = 1; i < rows.length; i++) {
      if (_normalizarNome(rows[i][1]) === nomeMotorista && rows[i][8] === 'Em_Transito') {
        return { success: false, error: 'Já existe uma rota ativa para este motorista. Finalize a rota anterior antes de iniciar uma nova.' };
      }
    }
    const id    = Utilities.getUuid();
    const ts    = data.timestamp || new Date().toISOString();
    sheet.appendRow([
      id,
      motorista.nome,
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
      '',
      _toNumberOrBlank(data.kmInicio),
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
    _ensureRotasColumns(sheet);
    const rows  = sheet.getDataRange().getValues();
    const ts    = data.timestamp || new Date().toISOString();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.rotaId) {
        sheet.getRange(i + 1,  9).setValue('Finalizado');
        sheet.getRange(i + 1, 11).setValue(ts);
        sheet.getRange(i + 1, 12).setValue(data.latitude  || '');
        sheet.getRange(i + 1, 13).setValue(data.longitude || '');
        sheet.getRange(i + 1, 15).setValue(_toNumberOrBlank(data.kmFim));
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
  try {
    const sheet = getOrCreateSheet(SHEET_ROTAS);
    _ensureRotasColumns(sheet);
    const rows  = sheet.getDataRange().getValues();
    const ts    = new Date().toISOString();
    const rotaIdStr = String(rotaId).trim();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === rotaIdStr) {
        // Só atualiza colunas que já existem, evitando erros em abas com
        // número reduzido de colunas.
        sheet.getRange(i + 1,  9).setValue('Finalizado');
        sheet.getRange(i + 1, 11).setValue(ts);
        sheet.getRange(i + 1, 12).setValue('');
        sheet.getRange(i + 1, 13).setValue('');
        // A coluna Km_Fim (15) só existe quando a planilha está atualizada;
        // em abas legadas com menos colunas, ignoramos o erro e apenas
        // registramos no log — a finalização principal já foi concluída.
        try {
          sheet.getRange(i + 1, 15).setValue('');
        } catch (kmErr) {
          try { Logger.log('forcarFinalizarRota: coluna Km_Fim indisponível: ' + kmErr.message); } catch (_) {}
        }
        _atualizarStatusVeiculo(rows[i][2], 'Disponivel');
        return { success: true, timestamp: ts };
      }
    }
    return { success: false, error: 'Rota não encontrada (id=' + rotaIdStr + ').' };
  } catch (err) {
    return { success: false, error: 'Falha ao forçar finalização: ' + err.message };
  }
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
    const hoje  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    let   countHoje = 0;
    for (let i = 1; i < rows.length; i++) {
      const horaInicio = String(rows[i][5]);
      if (horaInicio.startsWith(hoje)) countHoje++;
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
    return { success: true, rotas: rotas, countHoje: countHoje };
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

function cadastrarMotorista(data) {
  try {
    const nome = String((data && data.nome) || '').trim();
    if (!nome) return { success: false, error: 'Informe o nome do motorista.' };
    const sheet = getOrCreateSheet(SHEET_MOTORISTAS);
    const rows  = sheet.getDataRange().getValues();
    const nomeNorm = _normalizarNome(nome);
    for (let i = 1; i < rows.length; i++) {
      if (_normalizarNome(rows[i][0]) === nomeNorm) {
        sheet.getRange(i + 1, 1).setValue(nome);
        return { success: true, message: 'Motorista atualizado com sucesso.' };
      }
    }
    sheet.appendRow([nome]);
    return { success: true, message: 'Motorista cadastrado com sucesso.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getMotoristas() {
  try {
    const rows = getOrCreateSheet(SHEET_MOTORISTAS).getDataRange().getValues();
    const motoristas = [];
    for (let i = 1; i < rows.length; i++) {
      const nome = String(rows[i][0] || '').trim();
      if (!nome) continue;
      motoristas.push({ nome: nome });
    }
    return { success: true, motoristas: motoristas };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function validarMotorista(data) {
  try {
    const nomeNorm = _normalizarNome(data && data.nome);
    if (!nomeNorm) return { success: false, error: 'Informe o nome do motorista.' };
    const motorista = _buscarMotoristaPorNome(nomeNorm);
    if (!motorista) {
      return { success: false, error: 'Motorista não cadastrado. Solicite cadastro ao supervisor.' };
    }
    return { success: true, nome: motorista.nome };
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
    if (!config['Retencao_Dias']) config['Retencao_Dias'] = '30';
    return { success: true, config: config };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function setConfig(data) {
  try {
    const sheet = getOrCreateSheet(SHEET_CONFIG);
    const rows  = sheet.getDataRange().getValues();
    const updates = {};
    if (data.intervaloGPS !== undefined) updates['Intervalo_GPS'] = String(data.intervaloGPS);
    if (data.retencaoDias !== undefined) updates['Retencao_Dias'] = String(data.retencaoDias);
    const pendentes = Object.assign({}, updates);
    for (let i = 1; i < rows.length; i++) {
      const chave = rows[i][0];
      if (updates[chave] !== undefined) {
        sheet.getRange(i + 1, 2).setValue(updates[chave]);
        delete pendentes[chave];
      }
    }
    Object.keys(pendentes).forEach(function (k) {
      sheet.appendRow([k, pendentes[k]]);
    });
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
          longFim:    rotaRows[i][12],
          kmInicio:   rotaRows[i][13] !== undefined ? rotaRows[i][13] : '',
          kmFim:      rotaRows[i][14] !== undefined ? rotaRows[i][14] : ''
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
           '"Lat_Inicio","Long_Inicio","Status","Obs_Viagem","Hora_Fim","Lat_Fim","Long_Fim","Km_Inicio","Km_Fim"\n';

    const rotasFiltradas = [];
    for (let i = 1; i < rotaRows.length; i++) {
      const row = rotaRows[i];
      if (placa      && row[2] !== placa)                        continue;
      if (dataInicio && String(row[5]) < dataInicio)             continue;
      if (dataFim    && String(row[5]) > dataFim)                continue;
      rotasFiltradas.push(row);
      const rowCsv = [
        row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7],
        row[8], row[9], row[10], row[11], row[12],
        row[13] !== undefined ? row[13] : '',
        row[14] !== undefined ? row[14] : ''
      ];
      csv += rowCsv.map(function(c) {
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
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: 'Corpo da requisição ausente.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
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
      case 'cadastrarMotorista':
        result = cadastrarMotorista(payload);
        break;
      case 'getMotoristas':
        result = getMotoristas();
        break;
      case 'validarMotorista':
        result = validarMotorista(payload);
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
      case 'gerarRelatorioData':
        result = gerarRelatorioData(payload);
        break;
      case 'validarVeiculo':
        result = validarVeiculo(payload);
        break;
      case 'verificarSenha':
        result = verificarSenha(payload.senha);
        break;
      case 'removerVeiculo':
        result = removerVeiculo(payload.placa);
        break;
      case 'limparDadosAntigos':
        result = limparDadosAntigos(payload.dias);
        break;
      case 'alterarSenha':
        result = alterarSenha(payload);
        break;
      case 'buscarPorPlaca':
        result = buscarPorPlaca(payload.placa);
        break;
      case 'buscarRotaAtivaMotorista':
        result = buscarRotaAtivaMotorista(payload.nome);
        break;
      case 'registrarAbastecimento':
        result = registrarAbastecimento(payload);
        break;
      case 'listarAbastecimentos':
        result = listarAbastecimentos(payload);
        break;
      case 'registrarChecklistAvaria':
        result = registrarChecklistAvaria(payload);
        break;
      case 'listarChecklistAvarias':
        result = listarChecklistAvarias(payload);
        break;
      case 'confirmarCienciaChecklist':
        result = confirmarCienciaChecklist(payload);
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

// ── Validar Veículo ──────────────────────────────────────────
function validarVeiculo(data) {
  try {
    const placa  = (data.placa  || '').toUpperCase().trim();
    const modelo = (data.modelo || '').trim().toLowerCase();
    const sheet  = getOrCreateSheet(SHEET_VEICULOS);
    const rows   = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === placa) {
        if (rows[i][1].trim().toLowerCase() !== modelo) {
          return { success: false, error: 'Modelo incorreto. Modelo cadastrado: ' + rows[i][1] };
        }
        if (rows[i][2] === 'Em_Transito') {
          return { success: false, error: 'Veículo já está em uma rota ativa.' };
        }
        return { success: true };
      }
    }
    return { success: false, error: 'Veículo não encontrado. Solicite ao supervisor que libere o veículo.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function _normalizarNome(nome) {
  return String(nome || '').trim().toLowerCase();
}

function _buscarMotoristaPorNome(nomeNormalizado) {
  const rows = getOrCreateSheet(SHEET_MOTORISTAS).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const nome = String(rows[i][0] || '').trim();
    if (!nome) continue;
    if (_normalizarNome(nome) === nomeNormalizado) {
      return { nome: nome };
    }
  }
  return null;
}

// ── Verificar Senha do Supervisor ────────────────────────────
function verificarSenha(senha) {
  try {
    const sheet = getOrCreateSheet(SHEET_CONFIG);
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === 'Senha_Supervisor') {
        return { success: String(rows[i][1]) === String(senha) };
      }
    }
    // Senha não cadastrada: inicializa com padrão e aceita "123456"
    sheet.appendRow(['Senha_Supervisor', '123456']);
    return { success: senha === '123456' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── Alterar Senha do Supervisor ──────────────────────────────
function alterarSenha(data) {
  try {
    const check = verificarSenha(data.senhaAtual);
    if (!check.success) return { success: false, error: 'Senha atual incorreta.' };
    const sheet = getOrCreateSheet(SHEET_CONFIG);
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === 'Senha_Supervisor') {
        sheet.getRange(i + 1, 2).setValue(String(data.novaSenha));
        return { success: true };
      }
    }
    sheet.appendRow(['Senha_Supervisor', String(data.novaSenha)]);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── Gerar Dados do Relatório (estruturado para PDF) ──────────
function gerarRelatorioData(filtros) {
  try {
    const rotaRows = getOrCreateSheet(SHEET_ROTAS).getDataRange().getValues();
    const gpsRows  = getOrCreateSheet(SHEET_GPS).getDataRange().getValues();

    const gpsIndex = {};
    for (let i = 1; i < gpsRows.length; i++) {
      const rid = gpsRows[i][0];
      if (!gpsIndex[rid]) gpsIndex[rid] = [];
      gpsIndex[rid].push({
        timestamp:  String(gpsRows[i][1]),
        latitude:   gpsRows[i][2],
        longitude:  gpsRows[i][3],
        velocidade: gpsRows[i][4]
      });
    }

    const dataInicio = filtros.dataInicio ? filtros.dataInicio + 'T00:00:00' : null;
    const dataFim    = filtros.dataFim    ? filtros.dataFim    + 'T23:59:59' : null;
    const placa      = filtros.placa      ? filtros.placa.toUpperCase()       : null;

    const rotas = [];
    for (let i = 1; i < rotaRows.length; i++) {
      const row = rotaRows[i];
      if (placa      && row[2] !== placa)            continue;
      if (dataInicio && String(row[5]) < dataInicio) continue;
      if (dataFim    && String(row[5]) > dataFim)    continue;
      const rid = String(row[0]);
      rotas.push({
        id:         rid,
        nome:       String(row[1]),
        placa:      String(row[2]),
        modelo:     String(row[3]),
        obsInicial: String(row[4]),
        horaInicio: String(row[5]),
        latInicio:  String(row[6]),
        longInicio: String(row[7]),
        status:     String(row[8]),
        obsViagem:  String(row[9]),
        horaFim:    String(row[10]),
        latFim:     String(row[11]),
        longFim:    String(row[12]),
        kmInicio:   row[13] !== undefined ? String(row[13]) : '',
        kmFim:      row[14] !== undefined ? String(row[14]) : '',
        gpsLog:     gpsIndex[rid] || []
      });
    }
    return { success: true, rotas: rotas, total: rotas.length };
  } catch (err) {
    return { success: false, error: err.message };
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

// ── Buscar Rota por Placa ────────────────────────────────────
function buscarPorPlaca(placa) {
  try {
    if (!placa) return { success: false, error: 'Placa não informada.' };
    const placaUp   = placa.toUpperCase().trim();
    const rotaRows  = getOrCreateSheet(SHEET_ROTAS).getDataRange().getValues();
    const gpsRows   = getOrCreateSheet(SHEET_GPS).getDataRange().getValues();

    // Coletar todas as rotas da placa, do mais recente ao mais antigo
    const rotas = [];
    for (let i = rotaRows.length - 1; i >= 1; i--) {
      if (String(rotaRows[i][2]).toUpperCase() === placaUp) {
        rotas.push({
          id:         String(rotaRows[i][0]),
          nome:       String(rotaRows[i][1]),
          placa:      String(rotaRows[i][2]),
          modelo:     String(rotaRows[i][3]),
          obsInicial: String(rotaRows[i][4]),
          horaInicio: String(rotaRows[i][5]),
          latInicio:  String(rotaRows[i][6]),
          longInicio: String(rotaRows[i][7]),
          status:     String(rotaRows[i][8]),
          obsViagem:  String(rotaRows[i][9]),
          horaFim:    String(rotaRows[i][10]),
          latFim:     String(rotaRows[i][11]),
          longFim:    String(rotaRows[i][12]),
          kmInicio:   rotaRows[i][13] !== undefined ? String(rotaRows[i][13]) : '',
          kmFim:      rotaRows[i][14] !== undefined ? String(rotaRows[i][14]) : ''
        });
      }
    }

    if (rotas.length === 0) {
      return { success: false, error: 'Nenhuma rota encontrada para a placa ' + placaUp + '.' };
    }

    // Última posição GPS da rota mais recente
    const rotaAtual = rotas[0];
    const rotaIds = {};
    for (let i = 0; i < rotas.length; i++) rotaIds[rotas[i].id] = true;
    const localizacoes = [];
    for (let i = gpsRows.length - 1; i >= 1; i--) {
      const rid = String(gpsRows[i][0]);
      if (rotaIds[rid]) {
        localizacoes.push({
          rotaId: rid,
          timestamp:  String(gpsRows[i][1]),
          latitude:   gpsRows[i][2],
          longitude:  gpsRows[i][3],
          velocidade: gpsRows[i][4]
        });
      }
    }
    let ultimaPos = localizacoes.length > 0 ? localizacoes[0] : null;
    // Fallback: ponto de início da rota
    if (!ultimaPos && rotaAtual.latInicio && rotaAtual.latInicio !== '') {
      ultimaPos = {
        timestamp:  rotaAtual.horaInicio,
        latitude:   parseFloat(rotaAtual.latInicio),
        longitude:  parseFloat(rotaAtual.longInicio),
        velocidade: ''
      };
      localizacoes.push({
        rotaId: rotaAtual.id,
        timestamp:  rotaAtual.horaInicio,
        latitude:   parseFloat(rotaAtual.latInicio),
        longitude:  parseFloat(rotaAtual.longInicio),
        velocidade: ''
      });
    }

    return {
      success: true,
      rota: rotaAtual,
      ultimaPos: ultimaPos,
      localizacoes: localizacoes,
      totalRotas: rotas.length
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── Buscar Rota Ativa por Motorista ──────────────────────────
function buscarRotaAtivaMotorista(nome) {
  try {
    const sheet = getOrCreateSheet(SHEET_ROTAS);
    const rows  = sheet.getDataRange().getValues();
    const nomeMotorista = (nome || '').trim().toLowerCase();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][1]).trim().toLowerCase() === nomeMotorista && rows[i][8] === 'Em_Transito') {
        return {
          success: true,
          rota: {
            id:         String(rows[i][0]),
            nome:       String(rows[i][1]),
            placa:      String(rows[i][2]),
            modelo:     String(rows[i][3]),
            obsInicial: String(rows[i][4]),
            horaInicio: String(rows[i][5]),
            kmInicio:   rows[i][13] !== undefined ? String(rows[i][13]) : ''
          }
        };
      }
    }
    return { success: false };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function registrarAbastecimento(data) {
  try {
    const placa = (data.placa || '').toUpperCase().trim();
    const km = _toFiniteNumber(data.quilometragem);
    const litros = _toFiniteNumber(data.litros);
    const valor = _toFiniteNumber(data.valor);
    const dataHora = data.dataHora || new Date().toISOString();
    if (!placa) return { success: false, error: 'Placa não informada.' };
    if (km === null || km < 0) return { success: false, error: 'Quilometragem inválida.' };
    if (litros === null || litros <= 0) return { success: false, error: 'Litros inválidos.' };
    if (valor === null || valor <= 0) return { success: false, error: 'Valor inválido.' };

    const sheet = getOrCreateSheet(SHEET_ABAST);
    _ensureAbastecimentosColumns(sheet);
    const rows = sheet.getDataRange().getValues();
    const ultimoKm = _obterUltimoKmAbastecimentoPorPlaca(rows, placa);
    if (ultimoKm !== null && km <= ultimoKm) {
      return { success: false, error: 'Quilometragem deve ser maior que o último KM de abastecimento (' + ultimoKm + ').' };
    }
    let mediaKmL = 0;
    if (ultimoKm !== null && litros > 0) {
      mediaKmL = Math.round(((km - ultimoKm) / litros) * 100) / 100;
    }
    sheet.appendRow([placa, km, litros, valor, dataHora, mediaKmL]);
    return { success: true, mediaKmL: mediaKmL, ultimoKm: ultimoKm };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function listarAbastecimentos(filtros) {
  try {
    const sheet = getOrCreateSheet(SHEET_ABAST);
    _ensureAbastecimentosColumns(sheet);
    const rows = sheet.getDataRange().getValues();
    const placaFiltro = filtros && filtros.placa ? String(filtros.placa).toUpperCase().trim() : '';
    const itens = [];
    for (let i = rows.length - 1; i >= 1; i--) {
      const placa = String(rows[i][0]).toUpperCase();
      if (placaFiltro && placa !== placaFiltro) continue;
      itens.push({
        placa: placa,
        quilometragem: rows[i][1],
        litros: rows[i][2],
        valor: rows[i][3],
        dataHora: String(rows[i][4]),
        mediaKmL: _toFiniteNumber(rows[i][5])
      });
    }
    const resumo = _montarResumoAbastecimento(itens, placaFiltro);
    return { success: true, itens: itens, resumo: resumo };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function registrarChecklistAvaria(data) {
  try {
    const placa = String((data && data.placa) || '').toUpperCase().trim();
    const registradoPor = String((data && data.registradoPor) || '').trim();
    const status = String((data && data.status) || '').trim() || CHECKLIST_STATUS_ABERTO;
    const registradoEm = (data && data.registradoEm) ? String(data.registradoEm) : new Date().toISOString();
    const entradas = _normalizarEntradasChecklist(data);
    if (!placa) return { success: false, error: 'Placa não informada.' };
    if (!entradas.length) return { success: false, error: 'Informe ao menos uma avaria.' };
    const id = Utilities.getUuid();
    const sheet = getOrCreateSheet(SHEET_CHECKLIST);
    const ids = [];
    for (let i = 0; i < entradas.length; i++) {
      const entrada = entradas[i];
      const linhaId = i === 0 ? id : Utilities.getUuid();
      sheet.appendRow([linhaId, placa, entrada.parteCarro, entrada.descricaoAvaria, status, registradoEm, registradoPor, '', '']);
      ids.push(linhaId);
    }
    return { success: true, id: ids[0], ids: ids, total: ids.length };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function listarChecklistAvarias(filtros) {
  try {
    const sheet = getOrCreateSheet(SHEET_CHECKLIST);
    const rows = sheet.getDataRange().getValues();
    const placaFiltro = filtros && filtros.placa ? String(filtros.placa).toUpperCase().trim() : '';
    const statusFiltro = filtros && filtros.status ? String(filtros.status).trim().toLowerCase() : '';
    const itens = [];

    for (let i = rows.length - 1; i >= 1; i--) {
      const id = String(rows[i][0] || '').trim();
      const placa = String(rows[i][1] || '').toUpperCase().trim();
      const parte = String(rows[i][2] || '').trim();
      const descricao = String(rows[i][3] || '').trim();
      const status = String(rows[i][4] || '').trim() || CHECKLIST_STATUS_ABERTO;
      const registradoEm = String(rows[i][5] || '');
      const registradoPor = String(rows[i][6] || '');
      const cienteMotorista = String(rows[i][7] || '');
      const cienteEm = String(rows[i][8] || '');

      if (!id || !placa) continue;
      if (placaFiltro && placa !== placaFiltro) continue;
      if (statusFiltro && status.toLowerCase() !== statusFiltro) continue;

      itens.push({
        id: id,
        placa: placa,
        parteCarro: parte,
        descricaoAvaria: descricao,
        status: status,
        registradoEm: registradoEm,
        registradoPor: registradoPor,
        cienteMotorista: cienteMotorista,
        cienteEm: cienteEm
      });
    }

    return { success: true, itens: itens };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function confirmarCienciaChecklist(data) {
  try {
    const sheet = getOrCreateSheet(SHEET_CHECKLIST);
    const rows = sheet.getDataRange().getValues();
    const ids = (data && data.ids && Array.isArray(data.ids)) ? data.ids.map(function (id) {
      return String(id || '').trim();
    }).filter(function (id) { return !!id; }) : [];
    const idsLookup = {};
    for (let i = 0; i < ids.length; i++) idsLookup[ids[i]] = true;
    const placa = String((data && data.placa) || '').toUpperCase().trim();
    const motorista = String((data && data.motorista) || '').trim();
    const cienteEm = (data && data.timestamp) ? String(data.timestamp) : new Date().toISOString();
    let atualizados = 0;

    for (let i = 1; i < rows.length; i++) {
      const idRow = String(rows[i][0] || '').trim();
      const placaRow = String(rows[i][1] || '').toUpperCase().trim();
      const statusRow = String(rows[i][4] || '').trim() || CHECKLIST_STATUS_ABERTO;
      const matchById = ids.length > 0 && !!idsLookup[idRow];
      const matchByPlaca = ids.length === 0 && placa && placaRow === placa && statusRow === CHECKLIST_STATUS_ABERTO;
      if (!matchById && !matchByPlaca) continue;
      sheet.getRange(i + 1, 8).setValue(motorista);
      sheet.getRange(i + 1, 9).setValue(cienteEm);
      atualizados++;
    }

    return { success: true, atualizados: atualizados };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── Remover Veículo da frota ─────────────────────────────────
function removerVeiculo(placa) {
  try {
    const placaUp = String(placa || '').toUpperCase().trim();
    if (!placaUp) return { success: false, error: 'Placa não informada.' };

    // Impede remoção de veículo em rota ativa.
    const rotaRows = getOrCreateSheet(SHEET_ROTAS).getDataRange().getValues();
    for (let i = 1; i < rotaRows.length; i++) {
      if (String(rotaRows[i][2]).toUpperCase() === placaUp && rotaRows[i][8] === 'Em_Transito') {
        return { success: false, error: 'Veículo está em rota ativa. Encerre a rota antes de remover.' };
      }
    }

    const sheet = getOrCreateSheet(SHEET_VEICULOS);
    const rows  = sheet.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]).toUpperCase() === placaUp) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Veículo ' + placaUp + ' removido.' };
      }
    }
    return { success: false, error: 'Veículo ' + placaUp + ' não encontrado.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── Limpar Dados Antigos ─────────────────────────────────────
// Remove rotas finalizadas, logs GPS e abastecimentos mais antigos que
// `dias` (padrão: valor em Config.Retencao_Dias ou 30). Rotas ativas
// (Em_Transito) nunca são removidas. Pode ser disparado manualmente
// pelo supervisor ou por um time-driven trigger diário.
function limparDadosAntigos(dias) {
  try {
    let diasNum = Number(dias);
    if (!diasNum || diasNum <= 0) {
      const cfg = getConfig();
      diasNum = Number((cfg.config && cfg.config.Retencao_Dias) || 30);
    }
    if (!diasNum || diasNum <= 0) diasNum = 30;

    const limiteMs = Date.now() - diasNum * 24 * 60 * 60 * 1000;

    const resumo = { rotas: 0, gps: 0, abastecimentos: 0 };

    // Coleta IDs de rotas antigas e finalizadas a serem removidas.
    const rotaSheet = getOrCreateSheet(SHEET_ROTAS);
    const rotaRows  = rotaSheet.getDataRange().getValues();
    const idsRemovidos = {};
    for (let i = rotaRows.length - 1; i >= 1; i--) {
      const status = rotaRows[i][8];
      if (status === 'Em_Transito') continue;
      const refIso = rotaRows[i][10] || rotaRows[i][5]; // Hora_Fim ou Hora_Inicio
      const ts = refIso ? new Date(refIso).getTime() : 0;
      if (ts && ts < limiteMs) {
        idsRemovidos[String(rotaRows[i][0])] = true;
        rotaSheet.deleteRow(i + 1);
        resumo.rotas++;
      }
    }

    // Remove pontos GPS de rotas removidas ou cujos timestamps sejam antigos.
    const gpsSheet = getOrCreateSheet(SHEET_GPS);
    const gpsRows  = gpsSheet.getDataRange().getValues();
    for (let i = gpsRows.length - 1; i >= 1; i--) {
      const rid = String(gpsRows[i][0]);
      const ts  = gpsRows[i][1] ? new Date(gpsRows[i][1]).getTime() : 0;
      if (idsRemovidos[rid] || (ts && ts < limiteMs)) {
        gpsSheet.deleteRow(i + 1);
        resumo.gps++;
      }
    }

    // Remove abastecimentos antigos.
    const abastSheet = getOrCreateSheet(SHEET_ABAST);
    const abastRows  = abastSheet.getDataRange().getValues();
    for (let i = abastRows.length - 1; i >= 1; i--) {
      const ts = abastRows[i][4] ? new Date(abastRows[i][4]).getTime() : 0;
      if (ts && ts < limiteMs) {
        abastSheet.deleteRow(i + 1);
        resumo.abastecimentos++;
      }
    }

    return { success: true, dias: diasNum, removidos: resumo };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// Trigger diário opcional: configurar em Apps Script > Triggers para
// rodar uma vez por dia chamando esta função.
function triggerLimpezaDiaria() {
  return limparDadosAntigos();
}

function _ensureRotasColumns(sheet) {
  try {
    const header = sheet.getRange(1, 1, 1, Math.max(ROTAS_EXPECTED_COLS, sheet.getLastColumn())).getValues()[0];
    const idxKmInicio = 13; // coluna 14
    const idxKmFim = 14;    // coluna 15
    if (header[idxKmInicio] !== 'Km_Inicio') sheet.getRange(1, idxKmInicio + 1).setValue('Km_Inicio');
    if (header[idxKmFim] !== 'Km_Fim') sheet.getRange(1, idxKmFim + 1).setValue('Km_Fim');
  } catch (err) {
    Logger.log('Falha ao garantir colunas de quilometragem: ' + err.message);
  }
}

function _ensureAbastecimentosColumns(sheet) {
  try {
    const totalCols = Math.max(sheet.getLastColumn() || 0, 1);
    const header = sheet.getRange(1, 1, 1, totalCols).getValues()[0];
    if (header[5] !== 'Media_Km_L') sheet.getRange(1, 6).setValue('Media_Km_L');
  } catch (err) {
    Logger.log('Falha ao garantir coluna de média de abastecimento: ' + err.message);
  }
}

function _toNumberOrBlank(value) {
  return value !== undefined && value !== null && value !== '' ? Number(value) : '';
}

function _formatTs(iso) {
  try {
    const d = new Date(iso);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  } catch (_) {
    return iso;
  }
}

function _toFiniteNumber(value) {
  const num = Number(value);
  return isFinite(num) ? num : null;
}

function _obterUltimoKmAbastecimentoPorPlaca(rows, placa) {
  if (!rows || rows.length <= 1 || !placa) return null;
  for (let i = rows.length - 1; i >= 1; i--) {
    const placaRow = String(rows[i][0] || '').toUpperCase().trim();
    if (placaRow !== placa) continue;
    return _toFiniteNumber(rows[i][1]);
  }
  return null;
}

function _normalizarEntradasChecklist(data) {
  const itens = (data && Array.isArray(data.itens)) ? data.itens : [{
    parteCarro: data && data.parteCarro,
    descricaoAvaria: data && data.descricaoAvaria
  }];
  return itens.map(function (item) {
    return {
      parteCarro: String((item && item.parteCarro) || '').trim(),
      descricaoAvaria: String((item && item.descricaoAvaria) || '').trim()
    };
  }).filter(function (item) {
    return !!item.parteCarro && !!item.descricaoAvaria;
  });
}

function _obterUltimoKmRegistrado(placa) {
  if (!placa) return { km: null, origem: '' };
  const rows = getOrCreateSheet(SHEET_ROTAS).getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][2] || '').toUpperCase().trim() !== placa) continue;
    const kmFim = _toFiniteNumber(rows[i][14]);
    if (kmFim !== null) return { km: kmFim, origem: 'Km final da rota' };
    const kmInicio = _toFiniteNumber(rows[i][13]);
    if (kmInicio !== null) return { km: kmInicio, origem: 'Km inicial da rota' };
  }
  return { km: null, origem: '' };
}

function _montarResumoAbastecimento(itens, placaFiltro) {
  const resumo = {
    placa: placaFiltro || '',
    totalAbastecimentos: itens.length,
    totalLitros: 0,
    totalValor: 0,
    ultimoKmAbastecimento: null,
    ultimoKmRegistrado: null,
    origemKmRegistrado: '',
    diferencaKm: null,
    kmRodadosTotal: null,
    mediaKmRodados: null,
    consumoMedioKml: null
  };

  for (let i = 0; i < itens.length; i++) {
    resumo.totalLitros += Number(itens[i].litros || 0);
    resumo.totalValor += Number(itens[i].valor || 0);
  }

  if (!placaFiltro || !itens.length) return resumo;

  const ultimoKmAbastecimento = _toFiniteNumber(itens[0].quilometragem);
  if (ultimoKmAbastecimento !== null) resumo.ultimoKmAbastecimento = ultimoKmAbastecimento;

  const ultimoKmRegistrado = _obterUltimoKmRegistrado(placaFiltro);
  resumo.ultimoKmRegistrado = ultimoKmRegistrado.km;
  resumo.origemKmRegistrado = ultimoKmRegistrado.origem;
  if (resumo.ultimoKmRegistrado !== null && resumo.ultimoKmAbastecimento !== null) {
    resumo.diferencaKm = resumo.ultimoKmRegistrado - resumo.ultimoKmAbastecimento;
  }

  const ordemCronologica = itens.slice().sort(function (a, b) {
    return new Date(a.dataHora).getTime() - new Date(b.dataHora).getTime();
  });
  let somaKmRodados = 0;
  let somaLitrosReferencia = 0;
  let intervalosValidos = 0;
  for (let i = 1; i < ordemCronologica.length; i++) {
    const kmAnterior = _toFiniteNumber(ordemCronologica[i - 1].quilometragem);
    const kmAtual = _toFiniteNumber(ordemCronologica[i].quilometragem);
    const litrosAtual = _toFiniteNumber(ordemCronologica[i].litros);
    if (kmAnterior === null || kmAtual === null || litrosAtual === null || litrosAtual <= 0) continue;
    const kmRodados = kmAtual - kmAnterior;
    if (kmRodados <= 0) continue;
    somaKmRodados += kmRodados;
    somaLitrosReferencia += litrosAtual;
    intervalosValidos++;
  }
  if (intervalosValidos > 0) {
    resumo.kmRodadosTotal = somaKmRodados;
    resumo.mediaKmRodados = somaKmRodados / intervalosValidos;
    resumo.consumoMedioKml = somaLitrosReferencia > 0 ? somaKmRodados / somaLitrosReferencia : null;
  }
  return resumo;
}
