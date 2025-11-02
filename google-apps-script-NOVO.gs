/**
 * ============================================================================
 * GOOGLE APPS SCRIPT - API MONITOR CONTÁBIL
 * ============================================================================
 * 
 * INSTRUÇÕES:
 * 1. Abra sua planilha no Google Sheets
 * 2. Vá em: Extensões → Apps Script
 * 3. Delete TODO o código que estiver lá
 * 4. Cole este código
 * 5. Salve (Ctrl+S)
 * 6. Execute a função "configurarPlanilha" UMA VEZ para criar os cabeçalhos
 * 7. Implantar → Nova implantação → Aplicativo da Web
 *    - Executar como: Eu
 *    - Quem tem acesso: Qualquer pessoa
 * 8. Copie a URL e cole no index.html
 */

// ============================================================================
// CONFIGURAÇÃO - AJUSTE O NOME DA ABA AQUI
// ============================================================================
const SHEET_NAME = 'Página1'; // Mude para o nome da sua aba (Plan1, Sheet1, etc.)

// ============================================================================
// FUNÇÃO GET - Retorna todos os dados em formato JSON
// ============================================================================
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return criarResposta({
        error: `Planilha "${SHEET_NAME}" não encontrada. Verifique o nome da aba.`
      });
    }
    
    const lastRow = sheet.getLastRow();
    
    // Se só tem cabeçalho ou está vazia, retorna array vazio
    if (lastRow <= 1) {
      return criarResposta([]);
    }
    
    // Pegar todos os dados
    const range = sheet.getRange(1, 1, lastRow, 7);
    const values = range.getValues();
    
    // Cabeçalhos da primeira linha
    const headers = values[0];
    
    // Converter para array de objetos
    const dados = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      // Pular linhas completamente vazias
      if (!row[0] || row[0].toString().trim() === '') {
        continue;
      }
      
      const obj = {
        Nome: row[0] ? row[0].toString().trim() : '',
        CNPJ: row[1] ? row[1].toString().trim() : '',
        MRR: row[2] ? parseFloat(row[2]) || 0 : 0,
        Setup: row[3] ? parseFloat(row[3]) || 0 : 0,
        Mes: row[4] ? parseInt(row[4]) || '' : '',
        Pagamento: row[5] ? row[5].toString().trim() : '',
        InfoPagamento: row[6] ? row[6].toString().trim() : ''
      };
      
      dados.push(obj);
    }
    
    return criarResposta(dados);
    
  } catch (error) {
    Logger.log('Erro em doGet: ' + error.toString());
    return criarResposta({
      error: 'Erro ao buscar dados: ' + error.toString()
    });
  }
}

// ============================================================================
// FUNÇÃO POST - Adiciona, atualiza ou exclui dados
// ============================================================================
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return criarResposta({
        success: false,
        message: `Planilha "${SHEET_NAME}" não encontrada`
      });
    }
    
    // Parse dos dados recebidos
    const dados = JSON.parse(e.postData.contents);
    const action = dados.action;
    
    let resultado;
    
    switch(action) {
      case 'add':
        resultado = adicionarCliente(sheet, dados);
        break;
        
      case 'update':
        resultado = atualizarCliente(sheet, dados);
        break;
        
      case 'delete':
        resultado = excluirCliente(sheet, dados);
        break;
        
      case 'clear':
        resultado = limparTodosDados(sheet);
        break;
        
      default:
        return criarResposta({
          success: false,
          message: 'Ação não reconhecida: ' + action
        });
    }
    
    return criarResposta(resultado);
    
  } catch (error) {
    Logger.log('Erro em doPost: ' + error.toString());
    return criarResposta({
      success: false,
      message: 'Erro: ' + error.toString()
    });
  }
}

// ============================================================================
// FUNÇÃO ADICIONAR CLIENTE
// ============================================================================
function adicionarCliente(sheet, dados) {
  try {
    const novaLinha = [
      dados.Nome || '',
      dados.CNPJ || '',
      parseFloat(dados.MRR) || 0,
      parseFloat(dados.Setup) || 0,
      parseInt(dados.Mes) || '',
      dados.Pagamento || '',
      dados.InfoPagamento || ''
    ];
    
    sheet.appendRow(novaLinha);
    
    return {
      success: true,
      message: 'Cliente adicionado com sucesso'
    };
  } catch (error) {
    return {
      success: false,
      message: 'Erro ao adicionar: ' + error.toString()
    };
  }
}

// ============================================================================
// FUNÇÃO ATUALIZAR CLIENTE
// ============================================================================
function atualizarCliente(sheet, dados) {
  try {
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return {
        success: false,
        message: 'Nenhum dado para atualizar'
      };
    }
    
    // Buscar a linha que contém o cliente (por Nome e CNPJ originais)
    const range = sheet.getRange(2, 1, lastRow - 1, 2);
    const values = range.getValues();
    
    for (let i = 0; i < values.length; i++) {
      const nome = values[i][0].toString().trim();
      const cnpj = values[i][1].toString().trim();
      
      if (nome === dados.oldNome && cnpj === dados.oldCNPJ) {
        // Encontrou! Atualizar essa linha
        const rowNumber = i + 2; // +2 porque: array começa em 0, e pulamos o cabeçalho
        
        sheet.getRange(rowNumber, 1).setValue(dados.Nome || '');
        sheet.getRange(rowNumber, 2).setValue(dados.CNPJ || '');
        sheet.getRange(rowNumber, 3).setValue(parseFloat(dados.MRR) || 0);
        sheet.getRange(rowNumber, 4).setValue(parseFloat(dados.Setup) || 0);
        sheet.getRange(rowNumber, 5).setValue(parseInt(dados.Mes) || '');
        sheet.getRange(rowNumber, 6).setValue(dados.Pagamento || '');
        sheet.getRange(rowNumber, 7).setValue(dados.InfoPagamento || '');
        
        return {
          success: true,
          message: 'Cliente atualizado com sucesso'
        };
      }
    }
    
    return {
      success: false,
      message: 'Cliente não encontrado'
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'Erro ao atualizar: ' + error.toString()
    };
  }
}

// ============================================================================
// FUNÇÃO EXCLUIR CLIENTE
// ============================================================================
function excluirCliente(sheet, dados) {
  try {
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return {
        success: false,
        message: 'Nenhum dado para excluir'
      };
    }
    
    // Buscar a linha que contém o cliente
    const range = sheet.getRange(2, 1, lastRow - 1, 2);
    const values = range.getValues();
    
    for (let i = 0; i < values.length; i++) {
      const nome = values[i][0].toString().trim();
      const cnpj = values[i][1].toString().trim();
      
      if (nome === dados.Nome && cnpj === dados.CNPJ) {
        const rowNumber = i + 2;
        sheet.deleteRow(rowNumber);
        
        return {
          success: true,
          message: 'Cliente excluído com sucesso'
        };
      }
    }
    
    return {
      success: false,
      message: 'Cliente não encontrado'
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'Erro ao excluir: ' + error.toString()
    };
  }
}

// ============================================================================
// FUNÇÃO LIMPAR TODOS OS DADOS
// ============================================================================
function limparTodosDados(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
      // Deletar todas as linhas exceto o cabeçalho
      sheet.deleteRows(2, lastRow - 1);
    }
    
    return {
      success: true,
      message: 'Todos os dados foram limpos'
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'Erro ao limpar dados: ' + error.toString()
    };
  }
}

// ============================================================================
// FUNÇÃO AUXILIAR - Criar resposta JSON
// ============================================================================
function criarResposta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================================
// FUNÇÃO PARA TESTAR A CONEXÃO
// Execute esta função manualmente no Apps Script para verificar se tudo está OK
// ============================================================================
function testar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    Logger.log('❌ ERRO: Planilha "' + SHEET_NAME + '" não encontrada!');
    Logger.log('📋 Abas disponíveis:');
    ss.getSheets().forEach(s => Logger.log('  - ' + s.getName()));
    Logger.log('');
    Logger.log('💡 Solução: Altere a constante SHEET_NAME no início do código');
    return;
  }
  
  Logger.log('✅ Conexão OK!');
  Logger.log('📊 Planilha: ' + sheet.getName());
  Logger.log('📏 Total de linhas: ' + sheet.getLastRow());
  Logger.log('📏 Total de colunas: ' + sheet.getLastColumn());
  
  if (sheet.getLastRow() > 0) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log('📋 Cabeçalhos: ' + headers.join(', '));
  }
  
  Logger.log('');
  Logger.log('✅ Tudo pronto! Agora faça a implantação.');
}

// ============================================================================
// FUNÇÃO PARA CONFIGURAR A PLANILHA AUTOMATICAMENTE
// Execute esta função UMA VEZ para criar os cabeçalhos corretos
// ============================================================================
function configurarPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  // Se a aba não existe, criar
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    Logger.log('✅ Aba "' + SHEET_NAME + '" criada');
  }
  
  // Limpar tudo
  sheet.clear();
  
  // Adicionar cabeçalhos
  const headers = ['Nome', 'CNPJ', 'MRR', 'Setup', 'Mes', 'Pagamento', 'InfoPagamento'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Formatar cabeçalhos
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#FF6600');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  
  // Ajustar largura das colunas
  sheet.setColumnWidth(1, 250); // Nome
  sheet.setColumnWidth(2, 150); // CNPJ
  sheet.setColumnWidth(3, 100); // MRR
  sheet.setColumnWidth(4, 100); // Setup
  sheet.setColumnWidth(5, 80);  // Mes
  sheet.setColumnWidth(6, 120); // Pagamento
  sheet.setColumnWidth(7, 200); // InfoPagamento
  
  // Congelar primeira linha
  sheet.setFrozenRows(1);
  
  // Adicionar bordas
  headerRange.setBorder(true, true, true, true, true, true);
  
  Logger.log('✅ Planilha configurada com sucesso!');
  Logger.log('📋 Cabeçalhos: ' + headers.join(', '));
  Logger.log('');
  Logger.log('🎯 Próximos passos:');
  Logger.log('1. Execute a função "testar" para verificar');
  Logger.log('2. Faça a implantação do Web App');
  Logger.log('3. Copie a URL e cole no index.html');
}

// ============================================================================
// FUNÇÃO PARA ADICIONAR DADOS DE TESTE
// Execute esta função se quiser adicionar alguns dados de exemplo
// ============================================================================
function adicionarDadosTeste() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    Logger.log('❌ Execute primeiro a função "configurarPlanilha"');
    return;
  }
  
  const dadosTeste = [
    ['Empresa ABC Ltda', '12.345.678/0001-90', 1200, 5000, 11, 'PIX', 'Cliente premium'],
    ['Tech Solutions', '98.765.432/0001-10', 2000, 3000, 11, 'CARTÃO', 'Pagamento recorrente'],
    ['Contabilidade XYZ', '11.222.333/0001-44', 1500, 2000, 10, 'BOLETO', 'Vence dia 10']
  ];
  
  dadosTeste.forEach(linha => {
    sheet.appendRow(linha);
  });
  
  Logger.log('✅ Dados de teste adicionados!');
  Logger.log('📊 Total de clientes: ' + dadosTeste.length);
}
