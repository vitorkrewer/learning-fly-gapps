/**
 * Cria o Menu personalizado no Google Sheets
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🛠️ Emissor de Certificados')
    .addItem('Cadastrar Instituição', 'abrirCadastroParceiro')
    .addItem('Configurar Certificado', 'abrirEmissorCertificadoParceiro')
    .addItem('Emitir Certificados Configurados', 'abrirEmissaoLoteSalvo')
    .addItem('Relatório de Emissões', 'abrirRelatorioEmissoesCertificadoParceiro')
    .addToUi();
}

// Define o ID da planilha emissora como uma constante para garantir que o script sempre acesse o arquivo correto.
const SHEET_ID_EMISSOR_CERTIFICADOS = '[Insira o ID da sua Planilha]';

// Constantes para os nomes das abas
const SHEET_CONFIG = 'Config.Salvas';
const SHEET_EMITIDOS = 'CertificadosEmitidos';
const SHEET_LOGS = 'Logs.Atividades';
const SHEET_INSTITUICOES = 'Cadastro.Instituicoes';


/**
 * Exibe a interface de emissão de certificados.
 */
function abrirEmissorCertificadoParceiro() {
  const html = HtmlService.createTemplateFromFile('ui-emissao-simples')
    .evaluate()
    .setWidth(800)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'Emissor de Certificados');
}

/**
 * Exibe a interface de relatórios.
 */
function abrirRelatorioEmissoesCertificadoParceiro() {
  const html = HtmlService.createTemplateFromFile('ui-relatorio')
    .evaluate()
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Relatório de Emissões');
}


function abrirCadastroParceiro() {
  const html = HtmlService.createHtmlOutputFromFile('ui-cadastro-parceiro')
    .setWidth(600)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Cadastrar Nova Instituição');
}

/**
 * --- Abre a UI para emitir certificados de um lote salvo ---
 */
function abrirEmissaoLoteSalvo() {
  const html = HtmlService.createHtmlOutputFromFile('ui-emissao-lote-salvo')
    .setWidth(700)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Emitir a partir de Lote Salvo');
}

/**
 * --- Abre a página de ajuda e instruções ---
 */
function abrirAjuda() {
  const html = HtmlService.createHtmlOutputFromFile('ui-ajuda-emissor-certificado')
    .setWidth(900)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Guia de Ajuda - Emissor de Certificados');
}

/**
 * Inclui o conteúdo de outro arquivo HTML (para CSS e JS).
 * @param {string} filename O nome do arquivo a ser incluído.
 * @return {string} O conteúdo do arquivo.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * --- Salva os dados de um novo parceiro na planilha ---
 * @param {object} parceiroData Objeto com os dados do parceiro.
 * @return {string} JSON com o resultado da operação.
 */
function salvarParceiro(parceiroData) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID_EMISSOR_CERTIFICADOS);
    const sheet = ss.getSheetByName(SHEET_INSTITUICOES);
    const idColumn = sheet.getRange("A:A").getValues();

    // Validação para não duplicar ID
    for (let i = 0; i < idColumn.length; i++) {
      if (idColumn[i][0] == parceiroData.idParceiro) {
        return JSON.stringify({ success: false, message: `O ID '${parceiroData.idParceiro}' já existe. Por favor, utilize outro.` });
      }
    }

    sheet.appendRow([
      parceiroData.idParceiro,
      parceiroData.nomeInstituicao,
      parceiroData.responsavel,
      parceiroData.emailContato,
      parceiroData.telefone,
      new Date() // Data do Cadastro
    ]);
    
    return JSON.stringify({ success: true, message: 'Parceiro cadastrado com sucesso!' });
  } catch(e) {
    logAtividade("ERRO Cadastro Parceiro", e.message);
    return JSON.stringify({ success: false, message: `Erro ao salvar: ${e.message}` });
  }
}

/**
 * --- Busca a lista de parceiros para popular o dropdown ---
 * @return {string} JSON com a lista de parceiros.
 */
function buscarParceiros() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID_EMISSOR_CERTIFICADOS);
    const sheet = ss.getSheetByName(SHEET_INSTITUICOES);
    const data = sheet.getDataRange().getValues();
    
    const headers = data.shift(); // Remove cabeçalho
    const partners = data.map(row => {
      return {
        id: row[0],   // IDParceiro
        nome: row[1]  // NomeInstituicao
      };
    });
    return JSON.stringify({ success: true, data: partners });
  } catch (e) {
    logAtividade("ERRO Buscar Parceiros", e.message);
    return JSON.stringify({ success: false, message: `Erro ao buscar parceiros: ${e.message}` });
  }
}

/**
 * Salva uma nova configuração de emissão na aba 'Config.Salvas'.
 * @param {object} configObject O objeto com os dados da configuração.
 * @return {string} O ID da configuração salva.
 */
function salvarConfiguracao(configObject) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID_EMISSOR_CERTIFICADOS);
    const configSheet = ss.getSheetByName(SHEET_CONFIG);
    
    const configId = 'CONF-' + new Date().getTime();
    const dataCriacao = new Date();

    configSheet.appendRow([
      configId,
      configObject.idParceiro,
      configObject.nomeParceiro,
      configObject.NomeCertificado,
      configObject.NomeEvento,
      configObject.DataEvento,
      configObject.DataEmissao,
      configObject.MensagemCorpo,
      configObject.idSheetParticipantes,
      configObject.TemplateDocID,
      configObject.TargetFolderID,
      dataCriacao
    ]);

    logAtividade('Configuração Salva', `ID: ${configId}, Parceiro: ${configObject.nomeParceiro}`);
    return JSON.stringify({ success: true, configId: configId, idSheetParticipantes: configObject.idSheetParticipantes });
  } catch (e) {
    logAtividade('ERRO ao Salvar Config.', e.message);
    return JSON.stringify({ success: false, message: e.message });
  }
}


/**
 * Busca os dados dos participantes de uma planilha externa.
 * @param {string} sheetId O ID da planilha de participantes.
 * @return {string} JSON com os dados dos participantes ou erro.
 */
function buscarParticipantes(sheetId) {
  try {
    const participantsSS = SpreadsheetApp.openById(sheetId);
    const sheet = participantsSS.getSheets()[0]; // Pega a primeira aba
    const data = sheet.getDataRange().getValues();
    
    const headers = data.shift(); 
    
    const participantes = data.map(row => {
      let obj = {};
      headers.forEach((header, i) => {
        obj[header] = row[i];
      });
      return obj;
    });

    logAtividade('Busca de Participantes', `Planilha ID: ${sheetId}, Encontrados: ${participantes.length}`);
    return JSON.stringify({ success: true, data: participantes });

  } catch (e) {
    logAtividade('ERRO ao Buscar Participantes', `Planilha ID: ${sheetId}, Erro: ${e.message}`);
    return JSON.stringify({ success: false, message: 'Erro ao acessar a planilha de participantes. Verifique o ID e as permissões. Detalhes: ' + e.message });
  }
}

/**
 * Gera um código alfanumérico aleatório de 8 caracteres.
 * Usado para garantir a unicidade de cada certificado emitido.
 * @return {string} O código gerado.
 */
function gerarCodigoVerificador() {
  return Math.random().toString(36).substring(2, 10);
}

/**
 * Prepara uma string para ser usada em um nome de arquivo (estilo semântico).
 * Converte para minúsculas, remove acentos, substitui espaços por hífens 
 * e remove caracteres especiais.
 * @param {string} texto O texto original.
 * @return {string} O texto sanitizado.
 */
function sanitizarStringParaNomeArquivo(texto) {
  if (!texto) return '';
  return texto
    .toString()
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // Remove acentos
    .replace(/\s+/g, '-')       // Substitui espaços e quebras de linha por hífens
    .replace(/[^\w\-]+/g, '')   // Remove tudo que não for letra, número, _ ou -
    .replace(/\-\-+/g, '-');     // Garante que não haja hífens duplos
}

/**
 * Função principal que ATUALIZA os dados dos participantes e EMITE os certificados.
 * @param {object} emissaoData Objeto contendo config, participantes, e IDs.
 * @return {string} JSON com o resultado da operação.
 */
function emitirCertificados(emissaoData) {
    const { config, participantes, templateId, folderId, configId } = emissaoData;
    
    const emissorSS = SpreadsheetApp.openById(SHEET_ID_EMISSOR_CERTIFICADOS);
    const sheetEmitidos = emissorSS.getSheetByName(SHEET_EMITIDOS);
    
    const targetFolder = DriveApp.getFolderById(folderId);
    const templateFile = DriveApp.getFileById(templateId);

    const participantesSheet = SpreadsheetApp.openById(config.idSheetParticipantes).getSheets()[0];
    const headers = participantesSheet.getDataRange().getValues()[0];
    const statusColIndex = headers.indexOf('StatusEmissao') + 1;

    if (participantes.length > 0) {
        try {
            logAtividade('Atualização de Dados', `Iniciando atualização de dados para ${participantes.length} participantes.`);
            const dataToWrite = participantes.map(() => [
                config.idParceiro, config.nomeParceiro, config.NomeEvento, new Date(config.DataEvento)
            ]);
            const rangeToUpdate = participantesSheet.getRange(2, 1, participantes.length, 4);
            rangeToUpdate.setValues(dataToWrite);
            logAtividade('Atualização de Dados Concluída', `Dados do parceiro e evento atualizados.`);
        } catch (e) {
            const errorMessage = `Falha ao atualizar dados na Planilha de Participantes. Verifique permissões/estrutura. Detalhes: ${e.message}`;
            logAtividade('ERRO na Atualização de Dados', errorMessage);
            return JSON.stringify({ success: false, message: errorMessage });
        }
    }
    
    let erros = [];
    let sucessos = 0;

    const dataEmissaoProgramada = new Date(config.DataEmissaoProgramada || config.DataEmissao);
    const dataEmissaoFormatada = dataEmissaoProgramada.toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo' });

    participantes.forEach((participante, index) => {
        try {
            const codigoVerificador = gerarCodigoVerificador();
            const nomeParticipanteSanitizado = sanitizarStringParaNomeArquivo(participante.NomeParticipante);
            
            const tempFileName = `Temp-Certificado-${nomeParticipanteSanitizado}-${codigoVerificador}`;
            const tempFile = templateFile.makeCopy(tempFileName, targetFolder);
            const tempDoc = DocumentApp.openById(tempFile.getId());
            const body = tempDoc.getBody();

            body.replaceText('{{NOME_PARTICIPANTE}}', participante.NomeParticipante);
            body.replaceText('{{NOME_EVENTO}}', config.NomeEvento);
            body.replaceText('{{DATA_EVENTO}}', new Date(config.DataEvento).toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo' }));
            body.replaceText('{{MENSAGEM_CORPO}}', config.MensagemCorpo);
            body.replaceText('{{CPF}}', participante.CPF || '');
            body.replaceText('{{Data de Emissão}}', dataEmissaoFormatada);
            
            tempDoc.saveAndClose();

            const pdfBlob = tempFile.getAs('application/pdf');
            const pdfFileName = `${configId}_${config.idParceiro}_${nomeParticipanteSanitizado}_${codigoVerificador}.pdf`;
            const pdfFile = targetFolder.createFile(pdfBlob).setName(pdfFileName);
            
            tempFile.setTrashed(true);
            
            sheetEmitidos.appendRow([
                configId,
                config.idParceiro,
                config.nomeParceiro,
                participante.NomeParticipante,
                participante.EmailParticipante,
                new Date(), // Data real da emissão (para log)
                pdfFile.getUrl(),
                codigoVerificador
            ]);
            
            if(statusColIndex > 0){
              participantesSheet.getRange(index + 2, statusColIndex).setValue('Emitido com Sucesso');
            }

            sucessos++;
            logAtividade('Emissão de Certificado', `Sucesso para ${participante.NomeParticipante}, Código: ${codigoVerificador}`);

        } catch (e) {
            const errorDetail = `Erro ao emitir para ${participante.NomeParticipante}: ${e.message}`;
            erros.push(errorDetail);
            logAtividade('ERRO de Emissão', errorDetail);
            if(statusColIndex > 0){
               participantesSheet.getRange(index + 2, statusColIndex).setValue(`Falha na emissão: ${e.message.substring(0, 200)}`);
            }
        }
    });

    if (erros.length === 0) {
        return JSON.stringify({ success: true, message: `${sucessos} certificado(s) emitido(s) com sucesso!` });
    } else {
        return JSON.stringify({ success: false, message: `Operação concluída com ${erros.length} erro(s).`, details: erros });
    }
}


/**
 * Função de log simples.
 * @param {string} acao A ação realizada.
 * @param {string} detalhes Os detalhes da ação.
 */
function logAtividade(acao, detalhes) {
  try {
    const logSheet = SpreadsheetApp.openById(SHEET_ID_EMISSOR_CERTIFICADOS).getSheetByName(SHEET_LOGS);
    logSheet.appendRow([new Date(), acao, detalhes]);
  } catch(e) {
    // Caso a escrita no log falhe, registra no log do próprio Apps Script para depuração.
    console.error(`Falha ao escrever no log da planilha: Ação="${acao}", Detalhes="${detalhes}". Erro: ${e.message}`);
  }
}

/**
 * Busca e processa os dados da aba 'CertificadosEmitidos' para o relatório.
 * @param {object} filtros Objeto contendo { dataInicio: string, dataFim: string }.
 * @return {string} JSON com os dados agregados e os dados brutos filtrados.
 */
function buscarDadosRelatorio(filtros) {
  try {
    const emissorSS = SpreadsheetApp.openById(SHEET_ID_EMISSOR_CERTIFICADOS);
    const sheetEmitidos = emissorSS.getSheetByName(SHEET_EMITIDOS);
    const sheetConfig = emissorSS.getSheetByName(SHEET_CONFIG);

    // Mapeia ConfigID para NomeEvento para consulta rápida
    const mapaConfigEvento = sheetConfig.getDataRange().getValues()
      .slice(1) // Pula o cabeçalho
      .reduce((mapa, linha) => {
        mapa[linha[0]] = linha[4]; // linha[0] é ConfigID, linha[4] é NomeEvento
        return mapa;
      }, {});

    const dadosCompletos = sheetEmitidos.getDataRange().getValues().slice(1);

    const dataInicio = filtros.dataInicio ? new Date(filtros.dataInicio + "T00:00:00") : null;
    const dataFim = filtros.dataFim ? new Date(filtros.dataFim + "T23:59:59") : null;

    const dadosFiltrados = dadosCompletos.filter(linha => {
      const dataEmissao = new Date(linha[5]); // Coluna F (índice 5) é a DataEmissao
      if (dataInicio && dataEmissao < dataInicio) return false;
      if (dataFim && dataEmissao > dataFim) return false;
      return true;
    });

    const stats = {
      porParceiro: {},
      porEvento: {},
      porParticipante: {},
      total: dadosFiltrados.length
    };

    const dadosParaExportar = [
      ['ConfigID', 'ID Parceiro', 'Nome Parceiro', 'Nome Participante', 'Email', 'Data Emissão', 'Nome Evento', 'Link PDF', 'Código Verificador']
    ];

    dadosFiltrados.forEach(linha => {
      const configId = linha[0];
      const nomeParceiro = linha[2];
      const nomeParticipante = linha[3];
      const nomeEvento = mapaConfigEvento[configId] || "Evento não encontrado";

      // Agrega estatísticas
      stats.porParceiro[nomeParceiro] = (stats.porParceiro[nomeParceiro] || 0) + 1;
      stats.porEvento[nomeEvento] = (stats.porEvento[nomeEvento] || 0) + 1;
      stats.porParticipante[nomeParticipante] = (stats.porParticipante[nomeParticipante] || 0) + 1;

      // Adiciona a linha formatada para exportação
      dadosParaExportar.push([linha[0], linha[1], linha[2], linha[3], linha[4], linha[5], nomeEvento, linha[6], linha[7]]);
    });

    return JSON.stringify({ success: true, stats: stats, dadosParaExportar: dadosParaExportar });

  } catch (e) {
    logAtividade("ERRO Relatório", e.message);
    return JSON.stringify({ success: false, message: e.message });
  }
}

/**
 * Exporta os dados do relatório para uma nova Planilha Google.
 * @param {Array<Array<string>>} dados Os dados a serem exportados (deve incluir cabeçalho).
 * @return {string} JSON com a URL da nova planilha ou mensagem de erro.
 */
function exportarParaPlanilha(dados) {
  try {
    const dataFormatada = Utilities.formatDate(new Date(), "GMT-3", "yyyy-MM-dd_HH-mm-ss");
    const nomeArquivo = `Relatorio_Certificados_${dataFormatada}`;
    const novaPlanilha = SpreadsheetApp.create(nomeArquivo);
    const folha = novaPlanilha.getSheets()[0];
    
    folha.getRange(1, 1, dados.length, dados[0].length).setValues(dados);
    folha.setFrozenRows(1);
    folha.autoResizeColumns(1, dados[0].length);

    logAtividade("Exportação Planilha", `Relatório gerado em: ${novaPlanilha.getUrl()}`);
    return JSON.stringify({ success: true, url: novaPlanilha.getUrl() });

  } catch (e) {
    logAtividade("ERRO Exportação Planilha", e.message);
    return JSON.stringify({ success: false, message: e.message });
  }
}

/**
 * --- Busca todas as configurações salvas na aba 'Config.Salvas' ---
 * @return {string} JSON com a lista de configurações.
 */
function buscarConfigsSalvas() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID_EMISSOR_CERTIFICADOS).getSheetByName(SHEET_CONFIG);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Remove cabeçalho

    const configs = data.map(row => {
      let configObj = {};
      headers.forEach((header, i) => {
        configObj[header] = row[i];
      });
      return configObj;
    }).reverse(); // Mostra as mais recentes primeiro

    return JSON.stringify({ success: true, data: configs });
  } catch (e) {
    logAtividade("ERRO Buscar Configs", e.message);
    return JSON.stringify({ success: false, message: e.message });
  }
}

/**
 * --- Processa a emissão de um lote salvo, emitindo apenas para os pendentes ---
 * @param {string} configId O ID da configuração a ser processada.
 * @return {string} JSON com o resultado da emissão.
 */
function emitirCertificadosLoteSalvo(configId) {
  try {
    // 1. Encontrar a configuração
    const configSheet = SpreadsheetApp.openById(SHEET_ID_EMISSOR_CERTIFICADOS).getSheetByName(SHEET_CONFIG);
    const configs = configSheet.getDataRange().getValues();
    const headers = configs.shift();
    const configRow = configs.find(row => row[0] === configId);
    
    if (!configRow) {
      throw new Error("Configuração não encontrada.");
    }

    const config = {};
    headers.forEach((header, i) => {
      config[header] = configRow[i];
    });

    // 2. Buscar participantes e filtrar os pendentes
    const pSheet = SpreadsheetApp.openById(config.idSheetParticipantes).getSheets()[0];
    const pData = pSheet.getDataRange().getValues();
    const pHeaders = pData.shift();
    const statusColIndex = pHeaders.indexOf('StatusEmissao');

    if (statusColIndex === -1) {
      throw new Error(`A coluna 'StatusEmissao' não foi encontrada na planilha de participantes: ${config.idSheetParticipantes}`);
    }

    const participantesPendentes = pData.filter(row => row[statusColIndex] !== 'Emitido com Sucesso');

    if (participantesPendentes.length === 0) {
      return JSON.stringify({ success: true, message: "Nenhum certificado novo para emitir. Todos os participantes já foram processados." });
    }

    // 3. Montar os dados para a função de emissão existente
    const participantesParaEmitir = participantesPendentes.map(row => {
      let pObj = {};
      pHeaders.forEach((header, i) => { pObj[header] = row[i]; });
      return pObj;
    });

    const emissaoData = {
      config: config,
      participantes: participantesParaEmitir,
      templateId: config.TemplateDocID,
      folderId: config.TargetFolderID,
      configId: config.ConfigID
    };

    // 4. Chamar a função de emissão original (REUTILIZAÇÃO DE CÓDIGO!)
    return emitirCertificados(emissaoData);

  } catch (e) {
    logAtividade("ERRO Emissão Lote Salvo", e.message);
    return JSON.stringify({ success: false, message: e.message });
  }
}