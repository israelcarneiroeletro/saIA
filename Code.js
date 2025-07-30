// =================== CONFIGURAÇÕES GLOBAIS ===================
// AVISO: Mantenha sua chave de API segura. Use PropertiesService em produção.
// =================== CONFIGURAÇÕES ===================
const GEMINI_API_KEY = 'AIzaSyBsw7ylNp0YJfLER7pbtQaRqtgkYAXIQWY'; // <<< IMPORTANTE: SUBSTITUA PELA SUA CHAVE

const USE_FEW_SHOT_LEARNING = false; 

// ID da pasta que contém os exemplos para o "few-shot learning" da IA.
const DRIVE_FOLDER_ID_EXEMPLOS = "1lyBwonhg_yRaFh4VRewweM_cN2fJgr1B"; 
const FINAL_ANALYSIS_FILENAME = 'analise_final.json';
const DRAFT_ANALYSIS_FILENAME = 'analise_rascunho.json';

// ========== NOVAS CONSTANTES PARA PLANILHAS ==========
const REPORT_SPREADSHEET_TEMPLATE_ID = '15xj5R1RAk80qKdyNac-TFSXXF_ld3rL3DiAD8jQ3cQY';
const MODELS_DATABASE_SPREADSHEET_ID = '1CSaTYECwakM07PV_rMSQ-gIF6zuBj4YUfRlUKGs64lU';
const REPORT_FILENAME = 'Relatório de Ocorrências saIA.xlsx';
// =========================================================


// =================== FUNÇÃO PRINCIPAL DO WEBAPP ===================
function doGet(e) {
  // Verifica se a URL está pedindo a página do service worker
  if (e.parameter.page === 'sw') {
    return ContentService.createTextOutput(HtmlService.createTemplateFromFile('sw').evaluate().getContent())
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  // Verifica se a URL está pedindo o arquivo de manifesto
  if (e.parameter.page === 'manifest') {
    return ContentService.createTextOutput(HtmlService.createTemplateFromFile('manifest').evaluate().getContent())
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Se não for nenhum dos acima, serve a página principal do aplicativo
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('saIA | Eletromidia')
    .setFaviconUrl('https://i.ibb.co/WpHG5C25/OPEC-T1-OPEC-T.png')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =================== LÓGICA DE CACHE DE ENDEREÇOS ===================
function atualizarCacheDeEnderecos() {
  const SPREADSHEET_ID = '1CSaTYECwakM07PV_rMSQ-gIF6zuBj4YUfRlUKGs64lU'; 
  const SHEET_NAME = '8-Parada com linhas'; 
  const CHUNK_SIZE = 2000;

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    data.shift(); // Remove a linha do cabeçalho

    const addressMap = {};
    data.forEach(row => {
      const code = String(row[0]).trim();
      const street = String(row[1]).trim();
      const number = String(row[2]).trim();

      if (code && street) {
        let fullAddress = street;
        if (number && number !== '') {
          fullAddress += `, ${number}`;
        }
        addressMap[code] = fullAddress;
      }
    });

    const cache = CacheService.getScriptCache();
    const allKeys = Object.keys(addressMap);
    let chunkIndex = 0;

    for (let i = 0; i < allKeys.length; i += CHUNK_SIZE) {
      const chunk = {};
      const chunkKeys = allKeys.slice(i, i + CHUNK_SIZE);
      chunkKeys.forEach(key => {
        chunk[key] = addressMap[key];
      });
      
      const chunkKey = `address_chunk_${chunkIndex}`;
      cache.put(chunkKey, JSON.stringify(chunk), 21600); 
      Logger.log(`Chunk ${chunkKey} com ${Object.keys(chunk).length} itens salvo no cache.`);
      chunkIndex++;
    }
    
    cache.put('address_chunk_count', String(chunkIndex), 21600);
    Logger.log(`Processo concluído. Total de ${chunkIndex} chunks salvos.`);

  } catch (e) {
    Logger.log(`Falha crítica ao atualizar o cache de endereços: ${e.toString()}`);
  }
}

function getMapFromCache() {
  const cache = CacheService.getScriptCache();
  const chunkCountStr = cache.get('address_chunk_count');
  
  if (!chunkCountStr) {
    Logger.log('Cache vazio ou expirado. Tentando popular agora...');
    atualizarCacheDeEnderecos();
    const newChunkCountStr = cache.get('address_chunk_count');
    if (!newChunkCountStr) {
      Logger.log('Não foi possível popular o cache. Retornando mapa vazio.');
      return {};
    }
    return getMapFromCache();
  }

  const chunkCount = parseInt(chunkCountStr, 10);
  const finalMap = {};

  for (let i = 0; i < chunkCount; i++) {
    const chunkKey = `address_chunk_${i}`;
    const chunkJson = cache.get(chunkKey);
    if (chunkJson) {
      const chunk = JSON.parse(chunkJson);
      Object.assign(finalMap, chunk);
    } else {
       Logger.log(`AVISO: O chunk ${chunkKey} não foi encontrado no cache.`);
    }
  }
  
  Logger.log(`Mapa de endereços carregado do cache com ${Object.keys(finalMap).length} itens.`);
  return finalMap;
}

// =================== LÓGICA DE CACHE DE MODELOS (NOVO) ===================

/**
 * Dicionário para traduzir o modelo bruto da planilha para o modelo final.
 */
const MODEL_TRANSLATION_DICT = {
    'Abrigo Vidro Modelo Caos Leve': 'CAOS LEVE',
    'Abrigo Vidro Modelo Minimalista Leve': 'MINIMALISTA LEVE',
    'Abrigo Vidro Modelo Brutalista Leve': 'BRUTALISTA LEVE',
    'Abrigo Vidro Modelo CAOS': 'CAOS ESTRUTURADO',
    'Abrigo Concreto - Padrão SPTrans': 'ORFÃO',
    'Abrigo Vidro Modelo Brutalista': 'BRUTALISTA',
    'Abrigo USP': 'USP',
    'Abrigo EMTU': 'EMTU',
    'Abrigo Fora do Padrão SPTrans': 'TOTEM',
    'Abrigo Metálico Corbucci': 'CORBUCCI',
    'Abrigo Concessionária Rodovia': 'ORFÃO',
    'Abrigo Vidro Modelo Minimalista': 'MINIMALISTA',
    'Abrigo Metálico Barcelona': 'TOTEM',
    'Abrigo Metálico': 'ORFÃO',
    'Abrigo Metálico SP 450 Antigo': 'TOTEM',
    'Abrigo/cobertura Terminal SPTrans': 'ORFÃO',
    'Abrigo Metálico Corbucci Bidirecional': 'CORBUCCI',
    'Abrigo Concreto Antigos': 'TOTEM',
    'Abrigo outro município': 'ORFÃO',
    'Abrigo/cobertura Metrô/CPTM': 'TOTEM',
    'Abrigo Vidro Modelo Minimalista Bidirecional': 'MINIMALISTA BI',
    'Abrigo Monumento': 'ORFÃO',
    'Abrigo Vidro Modelo High Tech': 'MINIMALISTA TOP'
};


/**
 * Lê a planilha de modelos e salva um mapa (SPTrans Code -> Modelo) no cache.
 */
function atualizarCacheDeModelos() {
  const SHEET_NAME = '8-Parada com linhas';
  const CHUNK_SIZE = 2000;

  try {
    const sheet = SpreadsheetApp.openById(MODELS_DATABASE_SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues(); // Ler da linha 2 até a última, colunas A até I

    const modelMap = {};
    data.forEach(row => {
      const code = String(row[0]).trim(); // Coluna A: Parada (SPTrans)
      const rawModel = String(row[8]).trim(); // Coluna I: Abrigo

      if (code) {
        if (rawModel && MODEL_TRANSLATION_DICT[rawModel]) {
          modelMap[code] = MODEL_TRANSLATION_DICT[rawModel];
        } else if (!rawModel) {
           modelMap[code] = 'TOTEM'; // Se a célula do modelo estiver vazia
        } else {
           modelMap[code] = MODEL_TRANSLATION_DICT[rawModel] || 'TOTEM'; // Pega do dicionário ou usa um padrão
        }
      }
    });

    const cache = CacheService.getScriptCache();
    const allKeys = Object.keys(modelMap);
    let chunkIndex = 0;

    for (let i = 0; i < allKeys.length; i += CHUNK_SIZE) {
      const chunk = {};
      const chunkKeys = allKeys.slice(i, i + CHUNK_SIZE);
      chunkKeys.forEach(key => {
        chunk[key] = modelMap[key];
      });
      
      const chunkKey = `model_chunk_${chunkIndex}`;
      cache.put(chunkKey, JSON.stringify(chunk), 21600); // 6 horas de cache
      Logger.log(`Chunk de Modelo ${chunkKey} com ${Object.keys(chunk).length} itens salvo.`);
      chunkIndex++;
    }
    
    cache.put('model_chunk_count', String(chunkIndex), 21600);
    Logger.log(`Cache de modelos atualizado. Total de ${chunkIndex} chunks.`);

  } catch (e) {
    Logger.log(`Falha crítica ao atualizar o cache de modelos: ${e.toString()}`);
  }
}

/**
 * Busca um modelo de abrigo no cache a partir do código SPTrans.
 * @param {string} spTransCode O código a ser procurado.
 * @returns {string|null} O modelo do abrigo ou null se não for encontrado.
 */
function getModeloFromCache(spTransCode) {
    const cache = CacheService.getScriptCache();
    const chunkCountStr = cache.get('model_chunk_count');

    if (!chunkCountStr) {
        Logger.log('Cache de modelos vazio. Populando agora...');
        atualizarCacheDeModelos();
        // Após tentar popular, verifica novamente. Se ainda não houver, retorna null.
        const newChunkCountStr = cache.get('model_chunk_count');
        if (!newChunkCountStr) {
            Logger.log('Não foi possível popular o cache de modelos.');
            return null; 
        }
    }

    const chunkCount = parseInt(cache.get('model_chunk_count'), 10);
    for (let i = 0; i < chunkCount; i++) {
        const chunkKey = `model_chunk_${i}`;
        const chunkJson = cache.get(chunkKey);
        if (chunkJson) {
            const chunk = JSON.parse(chunkJson);
            if (chunk[spTransCode]) {
                return chunk[spTransCode]; // Retorna o modelo encontrado
            }
        }
    }
    
    Logger.log(`Modelo para o código ${spTransCode} não encontrado no cache.`);
    return null; // <-- ALTERAÇÃO PRINCIPAL: Retorna null para indicar "não encontrado"
}


// =================== LÓGICA DE SALVAMENTO EM PLANILHA (NOVO) ===================

/**
 * Verifica se a planilha de relatório existe na pasta raiz. Se não, cria a partir do template.
 * @param {GoogleAppsScript.Drive.Folder} rootFolder A pasta mãe do projeto.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} O objeto da aba 'CONFERE' da planilha de relatório.
 */
function getOrCreateReportSheet(rootFolder) {
  const existingFiles = rootFolder.getFilesByName(REPORT_FILENAME);
  
  let spreadsheet;
  if (existingFiles.hasNext()) {
    // Arquivo já existe, apenas abre
    const file = existingFiles.next();
    spreadsheet = SpreadsheetApp.openById(file.getId());
    Logger.log(`Planilha de relatório encontrada: ${file.getName()}`);
  } else {
    // Arquivo não existe, copia do template
    const templateFile = DriveApp.getFileById(REPORT_SPREADSHEET_TEMPLATE_ID);
    const newFile = templateFile.makeCopy(REPORT_FILENAME, rootFolder);
    spreadsheet = SpreadsheetApp.openById(newFile.getId());
    Logger.log(`Planilha de relatório criada a partir do template na pasta ${rootFolder.getName()}`);
  }
  
  return spreadsheet.getSheetByName('CONFERE');
}


/**
 * Salva ou atualiza as ocorrências de uma análise na planilha de relatório.
 * Esta versão remove e reaplica filtros para garantir a integridade dos dados.
 * @param {string} folderId O ID da pasta da análise (ponto de ônibus).
 * @param {string} rootFolderId O ID da pasta mãe do projeto.
 * @param {string} jsonString O JSON da análise.
 */
function saveDataToSpreadsheet(folderId, rootFolderId, jsonString) {
  let sheet;
  let filterRangeA1 = null;

  try {
    const analysisData = JSON.parse(jsonString);
    const folder = DriveApp.getFolderById(folderId);
    const rootFolder = DriveApp.getFolderById(rootFolderId);

    // 1. Obter a planilha de destino
    sheet = getOrCreateReportSheet(rootFolder);
    if (!sheet) {
      throw new Error("A aba 'CONFERE' não foi encontrada na planilha de relatório.");
    }

    // ===== INÍCIO DA CORREÇÃO DE FILTRO =====
    // Verifica se há um filtro ativo, salva seu range e o remove temporariamente.
    const existingFilter = sheet.getFilter();
    if (existingFilter) {
      Logger.log("Filtro ativo detectado. Removendo temporariamente.");
      filterRangeA1 = existingFilter.getRange().getA1Notation();
      existingFilter.remove();
    }
    // ===== FIM DA CORREÇÃO DE FILTRO =====

    // 2. Extrair dados comuns
    const folderName = folder.getName();
    const nameParts = folderName.match(/^(\d+)/);
    const spTransCode = nameParts ? nameParts[1] : "N/A";
    
    const addressMap = getMapFromCache();
    const formattedName = formatFolderName(folderName, addressMap);
    const finalAddressString = formattedName.split(' | ')[1] || "Endereço não encontrado";
    
    const analysisDate = new Date();
    const model = getModeloFromCache(spTransCode);

    // 3. Lógica de Atualização: Excluir linhas antigas para este SPTransCode
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const rowsToDelete = [];
    
    for (let i = values.length - 1; i >= 1; i--) { // Começa da última linha de dados
      const rowSpTransCode = values[i][1]; // Coluna B é o N° SPTrans
      if (String(rowSpTransCode).trim().toUpperCase() === String(spTransCode).trim().toUpperCase()) {
        rowsToDelete.push(i + 1);
      }
    }

    if (rowsToDelete.length > 0) {
      Logger.log(`Encontradas ${rowsToDelete.length} linhas para o código ${spTransCode}. Excluindo...`);
      rowsToDelete.sort((a, b) => b - a).forEach(rowIndex => {
        sheet.deleteRow(rowIndex);
      });
      Logger.log("Linhas antigas excluídas com sucesso.");
    }

    // 4. Preparar e Inserir Novas Linhas
    const rowsToAdd = [];
    const diag = analysisData.diagnostico_manutencao;

    if (diag) {
      (diag.servicos_executados || []).forEach(occ => {
        rowsToAdd.push([
          analysisDate, String(spTransCode).toUpperCase(), String(finalAddressString).toUpperCase(),
          String(model).toUpperCase(), String(occ.item_resolvido).toUpperCase(),
          String(findOccurrenceType(occ.item_resolvido, diag.ocorrencias_identificadas_antes)).toUpperCase(),
          'CORRIGIDO', occ.acao_realizada
        ]);
      });

      (diag.pendencias_remanescentes_depois || []).forEach(occ => {
        rowsToAdd.push([
          analysisDate, String(spTransCode).toUpperCase(), String(finalAddressString).toUpperCase(),
          String(model).toUpperCase(), String(occ.item).toUpperCase(),
          String(occ.tipo).toUpperCase(), 'PENDENTE', occ.descricao
        ]);
      });

      // Ocorrências marcadas como "inexistentes" (ocorrencias_excluidas) são intencionalmente ignoradas
      // e não são adicionadas à planilha, pois, como o nome sugere, elas não existem.
    }

    if (rowsToAdd.length > 0) {
      // Como não há filtro, getLastRow() é seguro, mas usamos getDataRange() por robustez.
      const lastRow = sheet.getDataRange().getLastRow();
      sheet.getRange(lastRow + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
      Logger.log(`${rowsToAdd.length} novas linhas adicionadas para o código ${spTransCode}.`);
    } else {
      Logger.log(`Nenhuma ocorrência para registrar para o código ${spTransCode}.`);
    }

  } catch (e) {
    Logger.log(`ERRO CRÍTICO ao salvar dados na planilha para a pasta ${folderId}: ${e.message} \nStack: ${e.stack}`);
    throw e; // Re-lança o erro para ser pego pelo frontend
  } finally {
    // ===== REAPLICAR O FILTRO =====
    // Este bloco SEMPRE será executado, garantindo que o filtro seja restaurado.
    if (sheet && filterRangeA1) {
      sheet.getRange(filterRangeA1).createFilter();
      Logger.log("Filtro re-aplicado com sucesso.");
    }
  }
}

/**
 * Função auxiliar para encontrar o tipo (Preventiva/Corretiva) de um serviço executado,
 * buscando na lista original de ocorrências.
 * @param {string} itemResolvido - O nome do item resolvido.
 * @param {Array} ocorrenciasAntes - O array de 'ocorrencias_identificadas_antes'.
 * @returns {string} - 'Preventiva', 'Corretiva' ou 'N/D'.
 */
function findOccurrenceType(itemResolvido, ocorrenciasAntes) {
    const original = (ocorrenciasAntes || []).find(o => o.item === itemResolvido);
    return original ? original.tipo : 'PREVENTIVA';
}

/**
 * Extrai o primeiro objeto JSON completo de uma string que pode conter lixo ou duplicações.
 * @param {string} text A string bruta retornada pela API.
 * @returns {string} Uma string contendo apenas o primeiro objeto JSON válido.
 */
function extractFirstJsonFromString(text) {
  const cleanedText = text.replace(/```json/g, '').replace(/```/g, '').trim();
  const startIndex = cleanedText.indexOf('{');
  
  if (startIndex === -1) {
    return ""; // Retorna string vazia se nenhum JSON for encontrado
  }

  let braceCount = 0;
  for (let i = startIndex; i < cleanedText.length; i++) {
    if (cleanedText[i] === '{') {
      braceCount++;
    } else if (cleanedText[i] === '}') {
      braceCount--;
    }
    
    if (braceCount === 0) {
      // Encontrou o final do primeiro objeto JSON completo
      return cleanedText.substring(startIndex, i + 1);
    }
  }
  
  return ""; // Retorna string vazia se o JSON estiver malformado/incompleto
}

// =================== FUNÇÕES EXPOSTAS PARA O FRONTEND ===================

function gerarConclusaoComGemma(jsonRascunhoString) {
  const apiKey = GEMINI_API_KEY; 
  const model = 'gemma-3-12b-it'; // Modelo original restaurado
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const promptsConfig = getPromptsConfig();
  const promptStaging = promptsConfig['GEMMA_CONCLUSION'].prompt;

  const requestData = {
    "contents": [{"parts": [{ "text": promptStaging + "\n\n" + jsonRascunhoString }]}],
    "generationConfig": { 
      "temperature": 0.2
      // response_mime_type não é especificado para obter texto simples
    }
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(requestData),
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 200) {
      const apiResponse = JSON.parse(responseText);
      
      if (apiResponse.candidates && apiResponse.candidates[0] && 
          apiResponse.candidates[0].finishReason === "STOP" &&
          apiResponse.candidates[0].content && apiResponse.candidates[0].content.parts && 
          apiResponse.candidates[0].content.parts[0].text) {
        
        let conclusaoGerada = apiResponse.candidates[0].content.parts[0].text;
        conclusaoGerada = conclusaoGerada.replace(/```json/g, '').replace(/```/g, '').trim();
        Logger.log("Conclusão gerada pelo Gemma: " + conclusaoGerada);
        return conclusaoGerada;

      } else {
        Logger.log(`Resposta da API não continha uma conclusão válida. Resposta completa: ${responseText}`);
        throw new Error("A IA não retornou uma conclusão válida. Motivo: " + (apiResponse.candidates ? apiResponse.candidates[0].finishReason : "Resposta inesperada"));
      }
    } else {
      Logger.log(`Erro na API do Gemma: Código ${responseCode}. Resposta: ${responseText}`);
      throw new Error(`O servidor da IA retornou um erro (${responseCode}).`);
    }
  } catch (e) {
    Logger.log('Uma exceção ocorreu ao chamar o Gemma: ' + e.toString());
    throw new Error('Falha ao comunicar com a IA para gerar a conclusão. Detalhes: ' + e.message);
  }
}

function checkInitialStatuses(folderIds) {
  const statuses = {};
  folderIds.forEach(id => {
    try {
      const folder = DriveApp.getFolderById(id);
      if (getFileIfExists(folder, FINAL_ANALYSIS_FILENAME)) {
        statuses[id] = 'reviewed';
      } else if (getFileIfExists(folder, DRAFT_ANALYSIS_FILENAME)) {
        statuses[id] = 'analyzed';
      } else {
        statuses[id] = 'pending';
      }
    } catch (e) { statuses[id] = 'error'; }
  });
  return statuses;
}

function formatAddressString(str) {
  if (!str) return '';
  const abbreviations = {
    'rua': 'R.', 'avenida': 'Av.', 'alameda': 'Al.', 'praça': 'Pça.', 'rodovia': 'Rod.', 'terminal': 'Term.'
  };
  let formattedStr = str.toLowerCase();
  for (const key in abbreviations) {
    const regex = new RegExp('^' + key + '\\b', 'i');
    if (regex.test(formattedStr)) {
      formattedStr = formattedStr.replace(regex, abbreviations[key]);
      break;
    }
  }
  return formattedStr.split(' ').map(word => {
    return word.charAt(0).toUpperCase() + word.slice(1);
  }).join(' ');
}

function formatFolderName(originalName, addressMap) {
  const match = originalName.match(/^(\d+)\s*(.*)/);
  if (!match) {
    return originalName; 
  }
  const spTransCode = match[1];
  let finalAddressString;
  if (addressMap && addressMap[spTransCode]) {
    finalAddressString = addressMap[spTransCode];
  } else {
    let restOfName = match[2].trim();
    finalAddressString = restOfName.replace(/(.*)\s+(\d+)$/, '$1, $2');
  }
  return `${spTransCode} | ${formatAddressString(finalAddressString)}`;
}

function getFolderTree(rootFolderId) {
  try {
    const addressMap = getMapFromCache();
    const rootFolder = DriveApp.getFolderById(rootFolderId);
    return buildTree(rootFolder, addressMap);
  } catch (e) {
    Logger.log(`Erro ao buscar árvore de pastas para ID ${rootFolderId}: ${e.toString()}`);
    return { error: true, message: `Pasta não encontrada ou acesso negado. Verifique o ID/Link: ${e.message}` };
  }
}

/**
 * Procura (ou cria) a planilha de relatório e retorna sua URL.
 * @param {string} rootFolderId O ID da pasta raiz da análise.
 * @returns {{url: string}|{error: boolean, message: string}}
 */
function getReportSheetUrl(rootFolderId) {
  try {
    const rootFolder = DriveApp.getFolderById(rootFolderId);
    // A função getOrCreateReportSheet já contém a lógica para encontrar ou criar o arquivo.
    // Ela retorna a *aba* da planilha, então precisamos pegar o spreadsheet pai e depois a URL.
    const sheet = getOrCreateReportSheet(rootFolder);
    const spreadsheet = sheet.getParent();
    const url = spreadsheet.getUrl();
    Logger.log(`URL da planilha de relatório encontrada: ${url}`);
    return { url: url };
  } catch (e) {
    Logger.log(`Erro ao obter URL da planilha de relatório: ${e.toString()}`);
    return { error: true, message: e.message };
  }
}


function getFolderListForMapping(rootFolderId) {
  try {
    const rootFolder = DriveApp.getFolderById(rootFolderId);
    const subfolders = rootFolder.getFolders();
    const folderList = [];
    const addressMap = getMapFromCache();

    while (subfolders.hasNext()) {
      const folder = subfolders.next();
      const folderId = folder.getId();
      const folderName = folder.getName();
      const nameParts = folderName.match(/^(\d+)/);
      const spTransCode = nameParts ? nameParts[1] : null;
      
      let status = 'pending';
      if (getFileIfExists(folder, FINAL_ANALYSIS_FILENAME)) {
        status = 'reviewed';
      } else if (getFileIfExists(folder, DRAFT_ANALYSIS_FILENAME)) {
        status = 'analyzed';
      }

      const model = spTransCode ? getModeloFromCache(spTransCode) : 'N/A';
      const equipmentType = (String(model).toUpperCase() === 'TOTEM') ? 'Totem' : 'Abrigo';

      folderList.push({
        id: folderId,
        name: formatFolderName(folderName, addressMap),
        status: status,
        equipmentType: equipmentType
      });
    }
    
    // Ordena a lista de pastas pelo nome formatado
    folderList.sort((a, b) => a.name.localeCompare(b.name));

    return { success: true, folders: folderList };

  } catch (e) {
    Logger.log(`Erro ao listar pastas para mapeamento (ID: ${rootFolderId}): ${e.toString()}`);
    return { error: true, message: `Falha ao mapear pastas: ${e.message}` };
  }
}


// Substitua a função saveAnalysis original por esta versão
function saveAnalysis(folderId, jsonString, isFinal) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const dataToSave = JSON.parse(jsonString);
    // ... (A lógica interna de saveAnalysis, como estava, continua aqui) ...
     // A lógica de limpar o source, checar houve_ocorrencia_previa etc.
     delete dataToSave.source;
     if (dataToSave.diagnostico_manutencao) {
       const diag = dataToSave.diagnostico_manutencao;
       const antesVazio = !diag.ocorrencias_identificadas_antes || diag.ocorrencias_identificadas_antes.length === 0;
       const executadosVazio = !diag.servicos_executados || diag.servicos_executados.length === 0;
       const pendentesVazio = !diag.pendencias_remanescentes_depois || diag.pendencias_remanescentes_depois.length === 0;
       diag.houve_ocorrencia_previa = (antesVazio && executadosVazio && pendentesVazio) ? "Não" : "Sim";
     }
     const finalJsonString = JSON.stringify(dataToSave, null, 2);
    
    // Lógica de apagar arquivos
    deleteFileIfExists(folder, DRAFT_ANALYSIS_FILENAME);
    deleteFileIfExists(folder, FINAL_ANALYSIS_FILENAME);

    const fileName = isFinal ? FINAL_ANALYSIS_FILENAME : DRAFT_ANALYSIS_FILENAME;
    folder.createFile(fileName, finalJsonString, MimeType.PLAIN_TEXT);
    Logger.log(`Análise salva como ${fileName} na pasta ${folder.getName()}.`);

    return { success: true, message: `Análise salva como ${fileName}.` };

  } catch (e) {
    Logger.log(`Erro ao salvar análise na pasta ${folderId}: ${e.toString()}`);
    return { success: false, message: e.toString() };
  }
}

// ========== NOVA FUNÇÃO DE FACHADA PARA O FRONTEND ==========
/**
 * Orquestra o salvamento do arquivo JSON e a atualização da planilha.
 * @param {string} folderId O ID da pasta da análise (ponto de ônibus).
 * @param {string} rootFolderId O ID da pasta mãe do projeto.
 * @param {string} jsonString O JSON da análise.
 * @param {boolean} isFinal Indica se é um salvamento final ('Concluir e Avançar').
 */
function finalizeAndSave(folderId, rootFolderId, jsonString, isFinal) {
  // Primeiro, salva o arquivo JSON (lógica que já existia em saveAnalysis)
  const jsonSaveResult = saveAnalysis(folderId, jsonString, isFinal);
  if (!jsonSaveResult.success) {
    // Se falhar ao salvar o JSON, interrompe e retorna o erro.
    return jsonSaveResult;
  }
  
  // Se o salvamento for FINAL, também salva na planilha
  if (isFinal) {
    try {
      saveDataToSpreadsheet(folderId, rootFolderId, jsonString);
      return { success: true, message: 'Análise final salva e relatório atualizado.' };
    } catch (e) {
      return { success: false, message: `JSON salvo, mas falha ao atualizar planilha: ${e.message}` };
    }
  }

  // Se for apenas rascunho, retorna o sucesso do salvamento do JSON.
  return jsonSaveResult;
}

function getFolderDataBundle(folderId) {
  let analysisResult = {
    analysisData: null,
    imageData: []
  };

  try {
    const folder = DriveApp.getFolderById(folderId);
    let analysisData;
    let source;

    const finalFile = getFileIfExists(folder, FINAL_ANALYSIS_FILENAME);
    if (finalFile) {
      analysisData = finalFile.getBlob().getDataAsString();
      source = 'final';
    } else {
      const draftFile = getFileIfExists(folder, DRAFT_ANALYSIS_FILENAME);
      if (draftFile) {
        analysisData = draftFile.getBlob().getDataAsString();
        source = 'draft';
      }
    }

    if (!analysisData) {
      Logger.log(`Nenhuma análise encontrada. Iniciando nova análise com IA para a pasta ${folder.getName()}.`);
      // ================= INÍCIO DA MODIFICAÇÃO CENTRAL =================
      // A chamada para generateNewIAAnalysis agora dispara a nova lógica de múltiplos passos.
      const newAnalysis = generateNewIAAnalysis(folder); 
      // ================= FIM DA MODIFICAÇÃO CENTRAL =================
      const resultObj = JSON.parse(newAnalysis);
      resultObj.source = 'new';
      analysisResult.analysisData = resultObj;
    } else {
      Logger.log(`Análise existente (source: ${source}) encontrada para a pasta ${folder.getName()}.`);
      const resultObj = JSON.parse(analysisData);
      resultObj.source = source;
      analysisResult.analysisData = resultObj;
    }

    // A lógica de carregar imagens para o frontend permanece a mesma.
    const files = folder.getFiles();
    const images = [];

    while (files.hasNext()) {
      const file = files.next();
      const mimeType = file.getMimeType();

      if ((mimeType === MimeType.JPEG || mimeType === MimeType.PNG) && file.getSize() > 0) {
        const thumbnailUrl = `https://drive.google.com/thumbnail?sz=w720&id=${file.getId()}`;
        const options = { headers: { 'Authorization': `Bearer ${ScriptApp.getOAuthToken()}` } };
        const thumbnailBlob = UrlFetchApp.fetch(thumbnailUrl, options).getBlob();
        const base64Data = Utilities.base64Encode(thumbnailBlob.getBytes());
        const name = file.getName();
        const match = name.match(/_(\d{8})_(\d{6})/);
        const sortKey = match ? match[1] + match[2] : name;

        images.push({
          name: name,
          sortKey: sortKey,
          data: `data:${thumbnailBlob.getContentType()};base64,${base64Data}`
        });
      }
    }

    images.sort((a, b) => a.sortKey.localeCompare(b.sortKey));
    analysisResult.imageData = images;

    return JSON.stringify(analysisResult);

  } catch (e) {
    Logger.log(`Erro ao processar o pacote de dados da pasta ${folderId}: ${e.toString()}`);
    return JSON.stringify({ analysisData: { error: true, message: e.toString() }, imageData: [] });
  }
}


function criarPayloadDeExemplos(exemplosFolder) {
  const exemplosConfig = [
    {
      antes: "EXE1-ANTES-12:48AM.jpg", depois: "EXE1-DEPOIS-01:16AM.jpg",
      laudo: `{ "contexto_analise": "Auditoria técnica comparativa de um abrigo de ônibus (antes vs. depois).", "quantidade_equipamentos_identificados": 1, "analise_temporal": { "timestamp_antes": "00:48:10", "timestamp_depois": "01:16:59", "duracao_servico_minutos": 28, "consistencia_temporal": "Duração de 28 min é consistente com uma manutenção preventiva civil (limpeza e pintura leve)." }, "diagnostico_manutencao": { "houve_ocorrencia_previa": "Sim", "ocorrencias_identificadas_antes": [ { "item": "Pintura Desgastada/Desbotada", "tipo": "Preventiva", "descricao": "A pintura da estrutura e do banco apresentava desgaste e pichações." }, { "item": "Calçada Suja", "tipo": "Preventiva", "descricao": "A calçada na área do abrigo estava suja e com objetos sob o banco." }, { "item": "Itinerário Danificado/Ausente", "tipo": "Preventiva", "descricao": "O adesivo do itinerário estava danificado e ilegível." } ], "servicos_executados": [ { "item_resolvido": "Pintura Desgastada/Desbotada", "acao_realizada": "Estrutura e banco foram pintados. A fita zebrada no banco indica tinta fresca." }, { "item_resolvido": "Calçada Suja", "acao_realizada": "Limpeza da calçada e remoção dos objetos realizada, evidenciada pelo piso molhado." }, { "item_resolvido": "Itinerário Danificado/Ausente", "acao_realizada": "O itinerário danificado foi removido." } ], "pendencias_remanescentes_depois": [ { "item": "Itinerário Danificado/Ausente", "tipo": "Preventiva", "status": "Pendente", "descricao": "O itinerário danificado foi retirado, mas um novo não foi aplicado. O espaço está ausente." } ] }, "conclusao_auditoria": "Manutenção preventiva civil realizada com sucesso, incluindo pintura e limpeza. Contudo, persiste a pendência da instalação de um novo adesivo de itinerário." }`
    },
    {
      antes: "EXE2-ANTES-10:30AM.jpg", depois: "EXE2-DEPOIS-10:42AM.jpg",
      laudo: `{"contexto_analise": "Auditoria técnica comparativa de um abrigo de ônibus (antes vs. depois).","quantidade_equipamentos_identificados": 1,"analise_temporal": {"timestamp_antes": "10:30:06","timestamp_depois": "10:42:14","duracao_servico_minutos": 12,"consistencia_temporal": "Duração de 12 min é consistente com a remoção de cartazes e limpeza rápida, mas incompatível com um reparo estrutural."},"diagnostico_manutencao": {"houve_ocorrencia_previa": "Sim","ocorrencias_identificadas_antes": [{"item": "Pichação/Colagem/Sujeira/Ferrugem Leve","tipo": "Preventiva","descricao": "Presença de múltiplos cartazes irregulares colados nos vidros do abrigo."},{"item": "Calçada Suja","tipo": "Preventiva","descricao": "Acúmulo de lixo e sujeira na calçada, principalmente atrás e sob o banco do abrigo."},{"item": "Estrutura Abalroada/Danificada","tipo": "Corretiva","descricao": "A cobertura (teto) do abrigo está visivelmente danificada e torta em sua estrutura de sustentação."}],"servicos_executados": [{"item_resolvido": "Pichação/Colagem/Sujeira/Ferrugem Leve","acao_realizada": "Todos os cartazes irregulares foram removidos dos vidros."},{"item_resolvido": "Calçada Suja","acao_realizada": "A calçada foi lavada e limpa, como evidenciado pelo novo padrão do piso visível e ausência de lixo."}],"pendencias_remanescentes_depois": [{"item": "Estrutura Abalroada/Danificada","tipo": "Corretiva","status": "Pendente","descricao": "A cobertura continua danificada/abalroada. O problema estrutural não foi corrigido."}]},"conclusao_auditoria": "Manutenção preventiva de limpeza foi executada, com remoção de cartazes e lavagem da calçada. A principal pendência corretiva, o dano estrutural na cobertura, não foi resolvida e permanece."}`
    },
    {
      antes: "EXE3-ANTES-03:08AM.jpg", depois: "EXE3-DEPOIS-03:32AM.jpg",
      laudo: `{"contexto_analise": "Auditoria técnica comparativa de um abrigo de ônibus (antes vs. depois).","quantidade_equipamentos_identificados": 1,"analise_temporal": {"timestamp_antes": "03:08:40","timestamp_depois": "03:32:02","duracao_servico_minutos": 24,"consistencia_temporal": "Duração de 24 min é consistente com uma manutenção preventiva de lavagem."},"diagnostico_manutencao": {"houve_ocorrencia_previa": "Não","ocorrencias_identificadas_antes": [],"servicos_executados": [{"item_resolvido": "Manutenção Preventiva de Rotina","acao_realizada": "Foi realizada a lavagem da calçada na área do abrigo, evidenciada pelo aspecto molhado e manchas de sabão na imagem posterior."}],"pendencias_remanescentes_depois": []},"conclusao_auditoria": "Nenhuma ocorrência grave foi identificada na imagem 'antes', indicando bom estado de conservação. Foi executada uma lavagem preventiva da calçada. Manutenção concluída com 100% de sucesso."}`
    }
  ];
  let conversationTurns = [];
  function getImageAsPart(fileName) {
    const files = exemplosFolder.getFilesByName(fileName);
    if (files.hasNext()) {
      const file = files.next();
      return { inline_data: { mime_type: file.getMimeType(), data: Utilities.base64Encode(file.getBlob().getBytes()) } };
    } else {
      Logger.log(`ERRO: Arquivo de exemplo "${fileName}" não encontrado.`);
      return null;
    }
  }
  for (const exemplo of exemplosConfig) {
    const partAntes = getImageAsPart(exemplo.antes);
    const partDepois = getImageAsPart(exemplo.depois);
    if (!partAntes || !partDepois) return null;
    conversationTurns.push({ "role": "user", "parts": [ { "text": `--- ANTES ---\nNome: ${exemplo.antes}` }, partAntes, { "text": `--- DEPOIS ---\nNome: ${exemplo.depois}` }, partDepois ] });
    conversationTurns.push({ "role": "model", "parts": [{ "text": exemplo.laudo }] });
  }
  return conversationTurns;
}


// =================== FUNÇÕES DE APOIO (BACKEND) - MODIFICADAS E NOVAS ===================

/**
 * Lê e analisa o arquivo prompts.html, usando cache para otimização.
 * @returns {Object} O objeto JSON com todas as configurações de prompt.
 */
function getPromptsConfig() {
  const cache = CacheService.getScriptCache();
  const cachedPrompts = cache.get('prompts_config');

  if (cachedPrompts) {
    Logger.log("Configuração de prompts carregada do cache.");
    return JSON.parse(cachedPrompts);
  }

  try {
    // Método correto para ler arquivos de projeto no Apps Script
    const content = HtmlService.createHtmlOutputFromFile('prompts').getContent();
    const promptsConfig = JSON.parse(content);
    
    // Armazena no cache por 1 hora para evitar leituras repetidas do arquivo
    cache.put('prompts_config', JSON.stringify(promptsConfig), 3600);
    Logger.log("Configuração de prompts lida do arquivo 'prompts.html' e salva no cache.");
    
    return promptsConfig;
  } catch (e) {
    Logger.log(`ERRO CRÍTICO: Não foi possível ler ou analisar 'prompts.html'. Verifique se o arquivo existe e se o JSON é válido. Erro: ${e.toString()}`);
    // Retorna um objeto de fallback para evitar que a aplicação quebre totalmente
    return { "FALLBACK": { "prompt": "ERRO: PROMPT DE FALLBACK NÃO CARREGADO.", "fewShotLearning": [] } };
  }
}

/**
 * Seleciona o prompt e os exemplos corretos com base no modelo do equipamento.
 * @param {string} modelName O nome do modelo (ex: 'TOTEM', 'MINIMALISTA').
 * @returns {{prompt: string, fewShotLearning: Array}} O prompt e os exemplos para o modelo.
 */
function getPromptForModel(modelName) {
  const promptsConfig = getPromptsConfig();
  const modelKey = String(modelName).toUpperCase();

  // Regra para TOTEM
  if (modelKey === 'TOTEM' || modelKey === 'TOTEM_GEMINI_FLASH') {
      Logger.log(`Modelo '${modelKey}' mapeado para o prompt 'TOTEM_GEMINI_FLASH'.`);
      return promptsConfig['TOTEM_GEMINI_FLASH'];
  }

  // Regra para o grupo CAOS
  if (['CAOS', 'CAOS LEVE', 'CAOS ESTRUTURADO'].includes(modelKey)) {
    Logger.log(`Modelo '${modelKey}' mapeado para o prompt 'CAOS'.`);
    return promptsConfig['CAOS'];
  }

  // Regra para o grupo MINIMALISTA e fallback para Abrigo
  if (['MINIMALISTA', 'MINIMALISTA LEVE', 'MINIMALISTA BI', 'MINIMALISTA TOP', 'ABRIGO'].includes(modelKey)) {
    Logger.log(`Modelo '${modelKey}' mapeado para o prompt 'MINIMALISTA'.`);
    return promptsConfig['MINIMALISTA'];
  }
  
  // Tenta um match direto para outros modelos que possam existir
  if (promptsConfig[modelKey]) {
    Logger.log(`Prompt específico encontrado para o modelo: ${modelKey}`);
    return promptsConfig[modelKey];
  }

  // Fallback final e seguro para qualquer outro caso de abrigo não mapeado
  Logger.log(`AVISO: Nenhum prompt configurado para o modelo '${modelKey}'. Usando 'MINIMALISTA' como fallback final.`);
  return promptsConfig['MINIMALISTA'];
}


function getFileIfExists(folder, fileName) {
    const files = folder.getFilesByName(fileName);
    return files.hasNext() ? files.next() : null;
}

function deleteFileIfExists(folder, fileName) {
    const file = getFileIfExists(folder, fileName);
    if (file) {
        file.setTrashed(true); 
    }
}

/**
 * NOVA FUNÇÃO PRINCIPAL DE ANÁLISE
 * Orquestra o fluxo: 1. Busca no Cache, 2. Fallback para IA, 3. Análise Detalhada.
 */
function generateNewIAAnalysis(folderToAnalyze) {
    try {
        // ======================= ETAPA 1: EXTRAÇÃO DE DADOS INICIAIS =======================
        const folderName = folderToAnalyze.getName();
        const nameParts = folderName.match(/^(\d+)/);
        const spTransCode = nameParts ? nameParts[1] : null;

        if (!spTransCode) {
            throw new Error(`Não foi possível extrair o código SPTrans do nome da pasta: '${folderName}'`);
        }
        
        Logger.log(`Iniciando análise para a pasta ${folderName} (Cód. SPTrans: ${spTransCode})`);

        // ======================= ETAPA 2: IDENTIFICAÇÃO DO EQUIPAMENTO (CACHE-FIRST) =======================
        let tipoEquipamentoFinal; // Esta variável irá guardar o tipo, seja do cache ou da IA.
        
        // Tenta obter o modelo diretamente do cache
        const modeloDoCache = getModeloFromCache(spTransCode);

        if (modeloDoCache) {
            // SUCESSO NO CACHE: O modelo foi encontrado na base de dados.
            Logger.log(`Sucesso na consulta ao Cache. Modelo encontrado: "${modeloDoCache}" para o código ${spTransCode}.`);
            
            // Define o tipo final com base no resultado do cache.
            // Se o modelo for 'TOTEM', o tipo é 'Totem'. Qualquer outro modelo é considerado 'Abrigo'.
            tipoEquipamentoFinal = (String(modeloDoCache).toUpperCase() === 'TOTEM') ? 'Totem' : 'Abrigo';

        } else {
            // FALHA NO CACHE (FALLBACK): O modelo não foi encontrado, portanto, usar a IA.
            Logger.log(`Modelo para o código ${spTransCode} não encontrado no cache. Acionando fallback para IA (Gemma).`);

            const identificationImageParts = collectImagePartsFromFolder(folderToAnalyze, 2); // Pega as 2 primeiras imagens para identificação
            if (identificationImageParts.length === 0) {
                throw new Error(`Nenhuma imagem encontrada na pasta '${folderName}' para a análise de fallback.`);
            }

            const respostaIdBruta = chamarApiGemma(
                identificationImageParts,
                PROMPT_IDENTIFICACAO_TIPO,
                'gemini-2.5-flash' // Mantém o uso do Gemma 3 12b para o fallback
            );

            const respostaIdJsonLimpo = extractFirstJsonFromString(respostaIdBruta);
            if (!respostaIdJsonLimpo) {
                throw new Error(`A IA de identificação (fallback) não retornou um JSON válido. Resposta recebida: ${respostaIdBruta}`);
            }

            const parsedIdResponse = JSON.parse(respostaIdJsonLimpo);
            if (!parsedIdResponse || !parsedIdResponse.resposta) {
                throw new Error(`O JSON de identificação (fallback) não contém a chave 'resposta'. Resposta: ${respostaIdJsonLimpo}`);
            }
            
            // O resultado da IA agora define o nosso tipo final.
            tipoEquipamentoFinal = parsedIdResponse.resposta.trim();
            Logger.log(`RACIOCÍNIO DA IA (Fallback): "${parsedIdResponse.pensamento || 'Não fornecido'}"`);
        }

        Logger.log(`Identificação finalizada. Tipo de equipamento definido como: "${tipoEquipamentoFinal}"`);

        // Validação final do tipo antes de prosseguir para a análise cara
        if (tipoEquipamentoFinal !== 'Abrigo' && tipoEquipamentoFinal !== 'Totem') {
            throw new Error(`O tipo de equipamento final determinado é inválido: "${tipoEquipamentoFinal}". A análise foi interrompida.`);
        }

        // ======================= ETAPA 3: ANÁLISE DETALHADA COM BASE NO TIPO FINAL =======================
        Logger.log(`Iniciando análise detalhada para um "${tipoEquipamentoFinal}"`);
        const allImageParts = collectImagePartsFromFolder(folderToAnalyze);
        
        if (allImageParts.length === 0) {
            throw new Error(`Nenhuma imagem foi encontrada na pasta '${folderName}' para processamento. A análise não pode continuar.`);
        }

        // --- LÓGICA DE SELEÇÃO DE PROMPT RESTAURADA ---
        const promptKey = (tipoEquipamentoFinal === 'Totem') ? 'TOTEM_GEMINI_FLASH' : (modeloDoCache || 'MINIMALISTA');
        const { prompt: dynamicPrompt, fewShotLearning: dynamicFewShot } = getPromptForModel(promptKey);
        // --- FIM DA LÓGICA RESTAURADA ---

        let analysisResult;

        // A lógica agora diferencia se o prompt é para Totem ou um tipo genérico de Abrigo
        if (tipoEquipamentoFinal === 'Totem') {
            // Para Totens, usamos a nova função que pede JSON diretamente.
            analysisResult = chamarApiTotemFlash(
                allImageParts,
                dynamicPrompt // Usa o prompt dinâmico para Totem
            );
            // O resultado já é um JSON limpo, não precisa de 'extract' ou checagem de 'laudo'.
            Logger.log(`Análise de Totem recebida com sucesso como JSON.`);

        } else { // tipoEquipamentoFinal === 'Abrigo'
            // Para abrigos, usamos a chamada que suporta few-shot learning
            analysisResult = chamarApiGeminiFlash(allImageParts, dynamicPrompt, dynamicFewShot);
        }
        
        // Retorna o resultado final da análise detalhada
        return analysisResult.replace(/^```json\s*/, '').replace(/```$/, '');

    } catch (e) {
        Logger.log(`Falha crítica no processo de análise da IA: ${e.toString()}\nStack: ${e.stack}`);
        // Retorna um objeto de erro padronizado para o frontend
        return JSON.stringify({ error: true, message: e.toString(), stack: e.stack });
    }
}


/**
 * Coleta blobs de imagem de uma pasta para o payload da API.
 * A função agora garante a ordenação cronológica ANTES de limitar o número de imagens.
 * @param {GoogleAppsScript.Drive.Folder} folder A pasta do Drive.
 * @param {number} [limit=0] O número máximo de imagens a serem coletadas (0 para ilimitado).
 * @returns {Array} Um array de 'parts' de imagem para a API Gemini.
 */
function collectImagePartsFromFolder(folder, limit = 0) {
  const files = folder.getFiles();
  const allImages = [];

  // 1. Primeiro, coletar referências de todos os arquivos de imagem e seus sortKeys
  while (files.hasNext()) {
    const file = files.next();
    const mimeType = file.getMimeType();
    if ((mimeType === MimeType.JPEG || mimeType === MimeType.PNG) && file.getSize() > 0) {
      const name = file.getName();
      // A chave de ordenação prioriza timestamps no nome do arquivo, senão usa o nome completo.
      // Ex: foto_20240720_153000.jpg vem antes de foto_20240720_153500.jpg
      const match = name.match(/_(\d{8})_(\d{6})/);
      const sortKey = match ? match[1] + match[2] : name;
      allImages.push({ file, sortKey });
    }
  }

  // 2. Ordenar a lista de imagens cronologicamente
  allImages.sort((a, b) => a.sortKey.localeCompare(b.sortKey));

  // 3. Aplicar o limite para pegar apenas as primeiras N imagens (se o limite for > 0)
  const imagesToProcess = limit > 0 ? allImages.slice(0, limit) : allImages;

  // 4. Agora, processar apenas as imagens selecionadas (já ordenadas e limitadas)
  const imageParts = [];
  for (const imageData of imagesToProcess) {
    const file = imageData.file;
    const imageBlob = file.getThumbnail();
    const imageBytes = imageBlob.getBytes();

    imageParts.push({ "text": `--- INÍCIO DA IMAGEM ---\nNome do Arquivo: ${file.getName()}` });
    imageParts.push({
      inline_data: { mime_type: imageBlob.getContentType(), data: Utilities.base64Encode(imageBytes) }
    });
  }

  Logger.log(`Coletadas ${imageParts.length / 2} imagens para processamento (ordenadas cronologicamente).`);
  return imageParts;
}

function buildTree(folder, addressMap) {
  const subfolders = folder.getFolders();
  const tree = {
    id: folder.getId(),
    name: formatFolderName(folder.getName(), addressMap),
    children: []
  };
  while (subfolders.hasNext()) {
    tree.children.push(buildTree(subfolders.next(), addressMap));
  }
  return tree;
}

/**
 * NOVA FUNÇÃO DE CHAMADA PARA MODELOS GEMMA
 * Não espera JSON nativo e usa um prompt para forçar a saída de texto (que será JSON ou texto simples).
 * @param {Array} imageParts As partes da imagem para a solicitação.
 * @param {string} prompt O prompt de texto a ser usado.
 * @param {string} model O nome do modelo Gemma a ser usado (ex: 'gemma-2-9b-it').
 * @returns {string} A resposta de texto da API.
 */
function chamarApiGemma(imageParts, prompt, model) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${GEMINI_API_KEY}`;
  
  const payload = {
    "contents": [
      { "role": "user", "parts": [ { "text": prompt }, ...imageParts ] }
    ],
    "generationConfig": {
      "temperature": 0
    },
    "safetySettings": [
      { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE" },
      { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE" },
      { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE" },
      { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE" }
    ]
  };
  
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  Logger.log(`Chamando a API Gemma com o modelo: ${model}`);
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (responseCode === 200) {
    const data = JSON.parse(responseBody);
    if (data.candidates && data.candidates[0].content && data.candidates[0].content.parts && data.candidates[0].content.parts[0].text) {
      return data.candidates[0].content.parts[0].text;
    } else {
      Logger.log(`Resposta da API Gemma em formato inesperado (200). Corpo: ${responseBody}`);
      throw new Error(`Resposta da API Gemma em formato inesperado. Detalhes: ${JSON.stringify(data)}`);
    }
  } else {
    Logger.log(`A API Gemma retornou um erro. Código: ${responseCode}. Corpo: ${responseBody}`);
    throw new Error(`A API Gemma retornou um erro (${responseCode}). Detalhes: ${responseBody}`);
  }
}

/**
 * FUNÇÃO DE CHAMADA ESPECÍFICA PARA GEMINI FLASH (ANÁLISE DE ABRIGOS)
 * Mantida para o fluxo de análise de abrigos, esperando uma resposta JSON.
 * @param {Array} imageParts As partes da imagem para a solicitação.
 * @param {string} prompt O prompt de sistema a ser usado.
 * @param {Array} fewShotExamples O payload de "few-shot learning" vindo do JSON.
 * @returns {string} A resposta JSON da API.
 */
function chamarApiGeminiFlash(imageParts, prompt, fewShotExamples) {
  const model = 'gemini-2.5-flash'; // Modelo original restaurado
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${GEMINI_API_KEY}`;
  
  // Constrói o payload de exemplos dinamicamente
  const exemplosPayload = [];
  if (fewShotExamples && fewShotExamples.length > 0) {
    fewShotExamples.forEach(ex => {
      // Aqui, estamos assumindo que os exemplos no JSON são texto.
      // Se eles precisarem referenciar imagens, a lógica precisará ser expandida.
      exemplosPayload.push({ "role": "user", "parts": [{ "text": ex.input }] });
      exemplosPayload.push({ "role": "model", "parts": [{ "text": ex.output }] });
    });
  }

  const payload = {
    "contents": [
      { "role": "user", "parts": [{ "text": prompt }] },
      { "role": "model", "parts": [{ "text": "Entendido. Estou pronto para iniciar a auditoria técnica. Por favor, forneça as imagens." }] },
      ...exemplosPayload,
      { "role": "user", "parts": imageParts }
    ],
    "generationConfig": {
      "response_mime_type": "application/json",
      "temperature": 0.2
    },
    "safetySettings": [
      { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE" },
      { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE" },
      { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE" },
      { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE" }
    ]
  };
  
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  Logger.log(`Chamando a API Gemini Flash para análise de Abrigo.`);
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (responseCode === 200) {
    const data = JSON.parse(responseBody);
    if (data.candidates && data.candidates[0].content && data.candidates[0].content.parts && data.candidates[0].content.parts[0].text) {
      return data.candidates[0].content.parts[0].text;
    } else {
      Logger.log(`Resposta da API Gemini Flash em formato inesperado (200). Corpo: ${responseBody}`);
      throw new Error(`Resposta da API Gemini Flash em formato inesperado. Detalhes: ${JSON.stringify(data)}`);
    }
  } else {
    Logger.log(`A API Gemini Flash retornou um erro. Código: ${responseCode}. Corpo: ${responseBody}`);
    throw new Error(`A API Gemini Flash retornou um erro (${responseCode}). Detalhes: ${responseBody}`);
// This project is a prototype
  }
}

/**
 * FUNÇÃO DE CHAMADA ESPECÍFICA PARA ANÁLISE DE TOTENS (GEMINI FLASH)
 * Solicita uma resposta JSON e não depende de few-shot learning.
 * @param {Array} imageParts As partes da imagem para a solicitação.
 * @param {string} prompt O prompt de sistema a ser usado.
 * @returns {string} A resposta JSON da API.
 */
function chamarApiTotemFlash(imageParts, prompt) {
  const model = 'gemini-2.5-flash-lite-preview-06-17'; 
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${GEMINI_API_KEY}`;
  
  const payload = {
    "contents": [
      { "role": "user", "parts": [{ "text": prompt }, ...imageParts] }
    ],
    "generationConfig": {
      "response_mime_type": "application/json",
      "temperature": 0.2
    },
    "safetySettings": [
      { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE" },
      { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE" },
      { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE" },
      { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE" }
    ]
  };
  
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  Logger.log(`Chamando a API Gemini Flash para análise de Totem (JSON output).`);
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (responseCode === 200) {
    const data = JSON.parse(responseBody);
    if (data.candidates && data.candidates[0].content && data.candidates[0].content.parts && data.candidates[0].content.parts[0].text) {
      return data.candidates[0].content.parts[0].text;
    } else {
      Logger.log(`Resposta da API para Totem em formato inesperado (200). Corpo: ${responseBody}`);
      throw new Error(`Resposta da API para Totem em formato inesperado. Detalhes: ${JSON.stringify(data)}`);
    }
  } else {
    Logger.log(`A API para Totem retornou um erro. Código: ${responseCode}. Corpo: ${responseBody}`);
    throw new Error(`A API para Totem retornou um erro (${responseCode}). Detalhes: ${responseBody}`);
  }
}
