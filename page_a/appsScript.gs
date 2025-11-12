function doPost(e) {
  try {
    // Verificar se existem dados
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Nenhum dado recebido do formulário');
    }
    
    // Obter os dados enviados
    const data = JSON.parse(e.postData.contents);
    
    // ID da SUA planilha
    const SHEET_ID = '1YyEolTrQJWjReyJzKTnDZlV1j54MZfYPpxltQohBt-Q';
    
    // Abrir a planilha
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    
    // Obter ou criar a aba "Leads"
    let sheet = spreadsheet.getSheetByName('Leads');
    if (!sheet) {
      // Se não existir a aba "Leads", cria uma nova
      sheet = spreadsheet.insertSheet('Leads');
    }
    
    // Verificar se a primeira linha está vazia (sem cabeçalho)
    const firstRow = sheet.getRange(1, 1, 1, 7).getValues()[0];
    const isFirstRowEmpty = firstRow.every(cell => cell === '');
    
    // Se a primeira linha estiver vazia, adicionar cabeçalhos
    if (isFirstRowEmpty) {
      sheet.getRange(1, 1, 1, 7).setValues([[
        'Nome', 'Celular', 'Email', 'Cidade', 'Empresa', 'Mensagem', 'Timestamp'
      ]]);
      
      // Formatar cabeçalho
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold')
                                .setBackground('#2563eb')
                                .setFontColor('white');
    }
    
    // Encontrar a próxima linha vazia (após a última linha com dados)
    const lastRow = sheet.getLastRow();
    const nextRow = lastRow + 1;
    
    // Adicionar nova linha na próxima posição disponível
    sheet.getRange(nextRow, 1, 1, 7).setValues([[
      data.nome || '',
      data.celular || '',
      data.email || '',
      data.cidade || '',
      data.empresa || '',
      data.mensagem || '',
      data.timestamp || new Date().toLocaleString('pt-BR')
    ]]);
    
    // Ajustar largura das colunas para melhor visualização
    sheet.autoResizeColumns(1, 7);
    
    // Retornar sucesso
    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: 'success', 
        message: 'Lead salvo com sucesso na aba Leads!',
        row: nextRow
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    // Retornar erro detalhado
    console.error('Erro completo:', error);
    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: 'error', 
        message: 'Erro ao salvar: ' + error.toString() 
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Função para testar se o script está funcionando
function doGet(e) {
  try {
    // Testar acesso à planilha
    const SHEET_ID = '1YyEolTrQJWjReyJzKTnDZlV1j54MZfYPpxltQohBt-Q';
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    let sheet = spreadsheet.getSheetByName('Leads');
    
    let sheetStatus = 'Aba Leads não existe';
    if (sheet) {
      const firstRow = sheet.getRange(1, 1, 1, 7).getValues()[0];
      const isFirstRowEmpty = firstRow.every(cell => cell === '');
      sheetStatus = {
        sheetExists: true,
        hasHeader: !isFirstRowEmpty,
        totalRows: sheet.getLastRow(),
        firstRow: firstRow
      };
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: 'active', 
        message: 'Script de captura de leads funcionando!',
        sheetStatus: sheetStatus,
        instructions: 'Use POST para enviar dados do formulário'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: 'error', 
        message: 'Erro no teste: ' + error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Função para testar manualmente
function testDoPost() {
  // Simular dados de teste
  const testData = {
    nome: 'Teste Silva',
    celular: '(11) 99999-9999',
    email: 'teste@email.com',
    cidade: 'São Paulo',
    empresa: 'Empresa Teste',
    mensagem: 'Mensagem de teste',
    timestamp: new Date().toLocaleString('pt-BR')
  };
  
  // Simular o objeto e que o doPost recebe
  const mockE = {
    postData: {
      contents: JSON.stringify(testData)
    }
  };
  
  // Executar doPost com dados de teste
  const result = doPost(mockE);
  Logger.log('Resultado do teste: ' + result.getContent());
}