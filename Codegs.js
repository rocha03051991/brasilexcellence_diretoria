/**
 * O arquivo principal do servidor (backend) da Intranet Brasil Excellence.
 * Versão: 6.6
 * Responsável por:
 * 1. Servir a interface do usuário (HTML/CSS/JS).
 * 2. Gerenciar o fluxo de recuperação de senha.
 * 3. Gerenciar a base de salários convenção.
 * 4. Gerar propostas e relatórios em PDF a partir de TEMPLATES em uma pasta específica do Drive.
 * 5. Fornecer dados para os KPIs com lógicas de cálculo e filtro revisadas.
 */

// Constantes Globais
const SPREADSHEET_ID = '1kz70LnOhEPM-9lMJBXYiWSZDfWWrFdhKSx8mynnZGto'; 
const SS = SpreadsheetApp.openById(SPREADSHEET_ID);

// --- CONFIGURAÇÃO PARA PDF ---
const PDF_GERADOS_FOLDER_ID = '12HDGX0gBrvoHfd5kRnbOsgbG6Jzj2jjS';
const DIR_RELATORIOS_FOLDER_ID = '1fl9HFf_6rMV2eCn2SZ0FJK3Dxub8p1Rj';
const TEMPLATES_FOLDER_ID = '1qnbhOHozWMJPiBG--xjea2tccGAQBGtM';

// --- SERVIDOR WEB ---

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.page = e.parameter.page || 'login';
  template.code = e.parameter.code || '';
  template.message = e.parameter.message || '';
  
  return template.evaluate()
    .setTitle('Intranet - Brasil Excellence')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// --- AUTENTICAÇÃO E SENHA ---

function checkUserCredentials(email, password) {
  try {
    const usersSheet = SS.getSheetByName('UTILIZADORES');
    if (!usersSheet) throw new Error("A aba 'UTILIZADORES' não foi encontrada.");
    const data = usersSheet.getDataRange().getValues();
    const headers = data.shift();
    const emailCol = headers.indexOf('Email'), nameCol = headers.indexOf('Nome'), roleCol = headers.indexOf('Perfil'), passwordCol = headers.indexOf('Senha');
    if ([emailCol, nameCol, roleCol, passwordCol].includes(-1)) throw new Error("A planilha 'UTILIZADORES' deve conter as colunas 'Email', 'Nome', 'Perfil' e 'Senha'.");
    for (const row of data) {
      if (row[emailCol] && row[emailCol].toLowerCase() === email.toLowerCase() && row[passwordCol] == password) {
        return { name: row[nameCol], email: row[emailCol], role: row[roleCol], avatar: 'https://i.pravatar.cc/40?u=' + row[emailCol] };
      }
    }
    return null;
  } catch (e) {
    Logger.log('Erro em checkUserCredentials: ' + e.message);
    return { error: e.message };
  }
}

function requestPasswordReset(email) {
  try {
    const usersSheet = SS.getSheetByName('UTILIZADORES');
    if (!usersSheet) return { success: false, message: "Aba de utilizadores não encontrada." };
    const data = usersSheet.getDataRange().getValues();
    const headers = data.shift();
    const emailCol = headers.indexOf('Email');
    const userRow = data.find(row => row[emailCol] && row[emailCol].toLowerCase() === email.toLowerCase());

    if (!userRow) {
      return { success: true, message: 'Se o e-mail estiver cadastrado, um link para redefinição de senha será enviado.' };
    }

    let resetSheet = SS.getSheetByName('PASSWORD_RESETS');
    if (!resetSheet) {
      resetSheet = SS.insertSheet('PASSWORD_RESETS');
      resetSheet.appendRow(['Email', 'Code', 'Expiration']);
    }

    const code = Math.random().toString(36).substring(2, 8).toUpperCase();
    const expiration = new Date(new Date().getTime() + 30 * 60000); // 30 minutos a partir de agora
    resetSheet.appendRow([email, code, expiration]);

    const resetUrl = `${getScriptUrl()}?page=reset&code=${code}`;
    const subject = "Redefinição de Senha - Intranet Brasil Excellence";
    const body = `Olá,\n\nRecebemos uma solicitação para redefinir sua senha.\n\nClique no link a seguir para criar uma nova senha: ${resetUrl}\n\nEste link é válido por 30 minutos.\n\nSe você não solicitou isso, pode ignorar este e-mail.\n\nAtenciosamente,\nSistema de Intranet Brasil Excellence.`;
    
    MailApp.sendEmail(email, subject, body);

    return { success: true, message: 'Se o e-mail estiver cadastrado, um link para redefinição de senha será enviado.' };
  } catch (e) {
    Logger.log('Erro em requestPasswordReset: ' + e.message);
    return { success: false, message: 'Ocorreu um erro ao processar sua solicitação.' };
  }
}

function resetPasswordWithCode(code, newPassword) {
  try {
    const resetSheet = SS.getSheetByName('PASSWORD_RESETS');
    if (!resetSheet) throw new Error("Aba de redefinição não encontrada.");

    const data = resetSheet.getDataRange().getValues();
    const headers = data.shift();
    const codeCol = headers.indexOf('Code'), emailCol = headers.indexOf('Email'), expCol = headers.indexOf('Expiration');
    
    let found = false;
    let rowIndex = -1;

    for (let i = data.length - 1; i >= 0; i--) { 
        if (data[i][codeCol] === code) {
            rowIndex = i;
            const expirationDate = new Date(data[i][expCol]);
            if (expirationDate < new Date()) {
                resetSheet.deleteRow(i + 2);
                return { success: false, message: 'Código de redefinição expirado. Por favor, solicite novamente.' };
            }
            found = true;
            break;
        }
    }

    if (!found) {
        return { success: false, message: 'Código de redefinição inválido.' };
    }

    const email = data[rowIndex][emailCol];
    
    const usersSheet = SS.getSheetByName('UTILIZADORES');
    const usersData = usersSheet.getDataRange().getValues();
    const usersHeaders = usersData.shift();
    const userEmailCol = usersHeaders.indexOf('Email'), passwordCol = usersHeaders.indexOf('Senha');

    for (let i = 0; i < usersData.length; i++) {
        if (usersData[i][userEmailCol].toLowerCase() === email.toLowerCase()) {
            usersSheet.getRange(i + 2, passwordCol + 1).setValue(newPassword);
            resetSheet.deleteRow(rowIndex + 2);
            return { success: true, message: 'Senha redefinida com sucesso! Pode fazer o login.' };
        }
    }

    return { success: false, message: 'Utilizador não encontrado.' };
  } catch (e) {
    Logger.log('Erro em resetPasswordWithCode: ' + e.message);
    return { success: false, message: 'Ocorreu um erro crítico: ' + e.message };
  }
}


// --- FUNÇÕES DE ESCRITA E MANIPULAÇÃO DE DADOS ---

function addFaturamentoPrevisao(previsaoData, userName) {
  try {
    const previsoesSheet = SS.getSheetByName('FINANCEIRO_PREVISOES');
    if (!previsoesSheet) throw new Error("A aba 'FINANCEIRO_PREVISOES' não foi encontrada.");
    
    const newRow = [
      'PREV-' + new Date().getTime(),
      new Date(),
      previsaoData.mes,
      previsaoData.ano,
      previsaoData.postoNome,
      previsaoData.valorBruto,
      previsaoData.descricao,
      userName
    ];
    previsoesSheet.appendRow(newRow);
    
    return { success: true, message: 'Previsão de faturamento lançada com sucesso!' };
  } catch (e) {
    Logger.log('Erro em addFaturamentoPrevisao: ' + e.message);
    return { success: false, message: e.message };
  }
}

function addNewCliente(clientData) {
  try {
    const clientesSheet = SS.getSheetByName('CLIENTES');
    if (!clientesSheet) throw new Error("A aba 'CLIENTES' não foi encontrada.");
    const newId = 'CLI-' + new Date().getTime();
    clientesSheet.appendRow([newId, clientData.postoId || '', clientData.cnpj, clientData.razaoSocial, clientData.nomeContato, clientData.celular, clientData.telefone, clientData.email, clientData.endereco]);
    return { success: true, message: 'Cliente cadastrado com sucesso!', newClient: {ID: newId, RAZAO_SOCIAL: clientData.razaoSocial, CNPJ: clientData.cnpj } };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function savePropostaGerada(propostaData) {
    try {
        const propostasSheet = SS.getSheetByName('PROPOSTAS_GERADAS');
        if (!propostasSheet) throw new Error("A aba 'PROPOSTAS_GERADAS' não foi encontrada.");
        const newId = 'PROP-' + new Date().getTime();
        propostasSheet.appendRow([newId, new Date(), propostaData.clienteNome, propostaData.valorTotal, 'Gerada', propostaData.responsavel]);
        return { success: true, message: 'Registro da proposta salvo com sucesso!' };
    } catch (e) {
        return { success: false, message: e.message };
    }
}

function addNewBaseSalario(cargoData) {
  try {
    const sheet = SS.getSheetByName('BASE SALARIOS CONVENCAO');
    if (!sheet) throw new Error("Aba 'BASE SALARIOS CONVENCAO' não encontrada.");
    const newId = 'CONV-' + new Date().getTime();
    const newRowData = [ newId, cargoData.nome, cargoData.salario ];
    sheet.appendRow(newRowData);
    return { success: true, message: 'Novo cargo de convenção adicionado!', newCargo: { id: newId, funcao: cargoData.nome, salario: cargoData.salario }};
  } catch (e) {
    return { success: false, message: 'Falha ao cadastrar o novo cargo de convenção: ' + e.message };
  }
}

function updateBaseSalariosConvencao(cargos) {
  try {
    const sheet = SS.getSheetByName('BASE SALARIOS CONVENCAO');
    if (!sheet) throw new Error("Aba 'BASE SALARIOS CONVENCAO' não foi encontrada.");
    
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    const headers = data.shift();
    const idCol = headers.indexOf('ID');
    const salarioCol = headers.indexOf('Salario_Convenção');
    
    if (idCol === -1 || salarioCol === -1) {
      throw new Error("Colunas 'ID' ou 'Salario_Convenção' não encontradas na aba.");
    }

    const dataMap = data.reduce((map, row, index) => {
      map[row[idCol]] = index + 2; // +2 para compensar o header e o 0-index
      return map;
    }, {});

    cargos.forEach(cargo => {
      const rowIndex = dataMap[cargo.id];
      if (rowIndex) {
        sheet.getRange(rowIndex, salarioCol + 1).setValue(cargo.salario);
      }
    });
    
    return { success: true, message: 'Base de salários convenção atualizada com sucesso!' };
  } catch (e) {
    Logger.log("Erro em updateBaseSalariosConvencao: " + e.message);
    return { success: false, message: e.message };
  }
}

// --- FUNÇÕES DE BUSCA DE DADOS ---

function getInitialData() {
  try {
    const postosSheet = SS.getSheetByName('POSTOS');
    const postosData = postosSheet ? postosSheet.getDataRange().getValues() : [['ID', 'NOME']];
    const postosHeaders = postosData.shift();
    const postos = postosData.map(row => ({ ID: row[0], NOME: row[1] })).filter(p => p.ID && p.NOME);

    const clientesSheet = SS.getSheetByName('CLIENTES');
    const clientesData = clientesSheet ? clientesSheet.getDataRange().getValues() : [['ID', 'POSTO_ID', 'Razao_Social', 'CNPJ']];
    const clientesHeaders = clientesData.shift();
    const clientes = clientesData.map(row => ({
        ID: row[clientesHeaders.indexOf('ID')], POSTO_ID: row[clientesHeaders.indexOf('POSTO_ID')], RAZAO_SOCIAL: row[clientesHeaders.indexOf('Razao_Social')], CNPJ: row[clientesHeaders.indexOf('CNPJ')],
    })).filter(c => c.RAZAO_SOCIAL);

    const orcamentoSheet = SS.getSheetByName('ORCAMENTO_PARA_PROPOSTA');
    const orcamentoData = orcamentoSheet ? orcamentoSheet.getDataRange().getValues() : [];
    if(orcamentoData.length > 1) orcamentoData.shift();
    const orcamentoBase = orcamentoData.map(row => ({ id: row[0], funcao: row[1], salario: row[2] })).filter(r => r.funcao && r.salario > 0);

    const salariosConvencaoSheet = SS.getSheetByName('BASE SALARIOS CONVENCAO');
    const salariosConvencaoData = salariosConvencaoSheet ? salariosConvencaoSheet.getDataRange().getValues() : [];
    if (salariosConvencaoData.length > 1) salariosConvencaoData.shift();
    const baseSalariosConvencao = salariosConvencaoData.map(row => ({ id: row[0], funcao: row[1], salario: row[2] })).filter(r => r.funcao);

    return JSON.parse(JSON.stringify({ postos, clientes, orcamentoBase, baseSalariosConvencao }));
  } catch (e) {
    return { error: e.message };
  }
}

function getDashboardData(periodo) {
  try {
    const data = {};
    const safeSum = (sheet, valueCol, filterCol, filterValue) => {
        const values = sheet.getDataRange().getValues();
        if (values.length < 2) return 0;
        const headers = values.shift();
        const valueColIndex = headers.indexOf(valueCol);
        const filterColIndex = filterCol ? headers.indexOf(filterCol) : -1;
        const mesColIndex = headers.indexOf('MES');
        const anoColIndex = headers.indexOf('ANO');
        
        if (valueColIndex === -1) return 0;

        return values.reduce((acc, row) => {
            let matchesFilter = true;
            if (filterColIndex !== -1 && filterValue && row[filterColIndex] != null) {
                matchesFilter = String(row[filterColIndex]).trim().toLowerCase() == String(filterValue).trim().toLowerCase();
            }
            
            let matchesPeriod = true;
            if(mesColIndex !== -1 && anoColIndex !== -1) {
              matchesPeriod = (String(row[mesColIndex]).trim().toLowerCase() === String(periodo.mes).trim().toLowerCase() && 
                               String(row[anoColIndex]).trim() === String(periodo.ano).trim());
            }

            if (matchesFilter && matchesPeriod) {
                const cellValue = row[valueColIndex];
                let numericValue = 0;
                if (typeof cellValue === 'number' && !isNaN(cellValue)) {
                    numericValue = cellValue;
                } else if (typeof cellValue === 'string' && cellValue.trim() !== '') {
                    const onlyNumbersAndComma = cellValue.replace(/[^0-9,]/g, '');
                    const dotDecimal = onlyNumbersAndComma.replace(',', '.');
                    numericValue = parseFloat(dotDecimal) || 0;
                }
                return acc + numericValue;
            }
            return acc;
        }, 0);
    };

    const previsoesSheet = SS.getSheetByName('FINANCEIRO_PREVISOES');
    const faturamentoSheet = SS.getSheetByName('FATURAMENTO');
    const aReceberSheet = SS.getSheetByName('CONTAS_A_RECEBER');
    const aPagarSheet = SS.getSheetByName('CONTAS_A_PAGAR');
    
    data.previaFaturamento = previsoesSheet ? safeSum(previsoesSheet, 'Valor_Bruto') : 0;
    data.faturamentoBruto = faturamentoSheet ? safeSum(faturamentoSheet, 'Valor_Bruto') : 0;
    data.faturamentoLiquido = faturamentoSheet ? safeSum(faturamentoSheet, 'Valor_Liquido') : 0; 
    data.totalPago = aPagarSheet ? safeSum(aPagarSheet, 'VALOR_TOTAL', 'Status', 'Pago') : 0;
    data.totalAPagar = aPagarSheet ? safeSum(aPagarSheet, 'VALOR_TOTAL') : 0;

    // Lógica corrigida para Total a Receber
    let totalAReceber = 0;
    let totalRecebido = 0;
    if (aReceberSheet) {
        const values = aReceberSheet.getDataRange().getValues();
        if (values.length > 1) {
            const headers = values.shift();
            const valorIndex = headers.indexOf('VALOR');
            const statusIndex = headers.indexOf('STATUS_PAGAMENTO');
            const mesIndex = headers.indexOf('MES');
            const anoIndex = headers.indexOf('ANO');

            if (valorIndex !== -1 && statusIndex !== -1 && mesIndex !== -1 && anoIndex !== -1) {
                values.forEach(row => {
                    const matchesPeriod = String(row[mesIndex]).trim().toLowerCase() === String(periodo.mes).trim().toLowerCase() && 
                                          String(row[anoIndex]).trim() === String(periodo.ano).trim();
                    if (matchesPeriod) {
                        const status = String(row[statusIndex]).trim().toLowerCase();
                        let numericValue = 0;
                        const cellValue = row[valorIndex];
                        if (typeof cellValue === 'number' && !isNaN(cellValue)) {
                            numericValue = cellValue;
                        } else if (typeof cellValue === 'string' && cellValue.trim() !== '') {
                            const onlyNumbersAndComma = cellValue.replace(/[^0-9,]/g, '');
                            numericValue = parseFloat(onlyNumbersAndComma.replace(',', '.')) || 0;
                        }

                        if (status === 'recebido') {
                            totalRecebido += numericValue;
                        } else if (status === 'pendente' || status === 'atrasado') {
                            totalAReceber += numericValue;
                        }
                    }
                });
            }
        }
    }
    data.totalAReceber = totalAReceber;
    data.totalRecebido = totalRecebido;

    const formatToBRL = (value) => (value || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });

    return JSON.parse(JSON.stringify({ 
        previaFaturamento: formatToBRL(data.previaFaturamento),
        faturamentoBruto: formatToBRL(data.faturamentoBruto),
        faturamentoLiquido: formatToBRL(data.faturamentoLiquido),
        totalAReceber: formatToBRL(data.totalAReceber), 
        totalRecebido: formatToBRL(data.totalRecebido),
        totalAPagar: formatToBRL(data.totalAPagar), 
        totalPago: formatToBRL(data.totalPago)
    }));
  } catch (e) {
    Logger.log("Erro em getDashboardData: " + e.message);
    return { error: 'Ocorreu um erro ao buscar os dados: ' + e.message };
  }
}

function getKpiDetails(kpiName, periodo) {
  try {
      let sheetName = '';
      switch(kpiName) {
          case 'previaFaturamento': sheetName = 'FINANCEIRO_PREVISOES'; break;
          case 'faturamento': sheetName = 'FATURAMENTO'; break;
          case 'contasAReceber': sheetName = 'CONTAS_A_RECEBER'; break;
          case 'totalAPagar': sheetName = 'CONTAS_A_PAGAR'; break;
          default: throw new Error('KPI desconhecido.');
      }
      const sheet = SS.getSheetByName(sheetName);
      if (!sheet) return { headers: [], data: [], error: `A aba '${sheetName}' não foi encontrada.` };
      
      const allData = sheet.getDataRange().getValues();
      const headers = allData.shift() || [];
      const mesColIndex = headers.indexOf('MES');
      const anoColIndex = headers.indexOf('ANO');

      if (!periodo || !periodo.mes || !periodo.ano || mesColIndex === -1 || anoColIndex === -1) {
        return JSON.parse(JSON.stringify({ headers, data: allData, sheetName }));
      }
      
      const filteredData = allData.filter(row => {
        return (String(row[mesColIndex]).trim().toLowerCase() === String(periodo.mes).trim().toLowerCase() && 
                String(row[anoColIndex]).trim() === String(periodo.ano).trim());
      });

      return JSON.parse(JSON.stringify({ headers, data: filteredData, sheetName }));
  } catch (e) {
      return { error: e.message };
  }
}

// --- GERAÇÃO DE PDF ---

function generatePdfFromTemplate(propostaData) {
  try {
    const { cliente, items, total } = propostaData;
    const today = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy");
    
    const templatesFolder = DriveApp.getFolderById(TEMPLATES_FOLDER_ID);
    const templateFiles = templatesFolder.getFilesByName('Proposta Comercial Template');
    if (!templateFiles.hasNext()) {
      throw new Error("Template 'Proposta Comercial Template' não encontrado na pasta de templates.");
    }
    const templateFile = templateFiles.next();
    const outputFolder = DriveApp.getFolderById(PDF_GERADOS_FOLDER_ID);
    
    const newFileName = `Proposta - ${cliente.RAZAO_SOCIAL} - ${today}`;
    const newFile = templateFile.makeCopy(newFileName, outputFolder);
    
    const doc = DocumentApp.openById(newFile.getId());
    const body = doc.getBody();
    
    body.replaceText('{{CLIENTE}}', cliente.RAZAO_SOCIAL || '');
    body.replaceText('{{CNPJ}}', cliente.CNPJ || 'Não informado');
    body.replaceText('{{DATA}}', today);
    body.replaceText('{{TOTAL_SERVICOS}}', total);
    body.replaceText('{{TOTAL_BENEFICIOS}}', 'R$ 0,00');
    body.replaceText('{{TOTAL_GLOBAL}}', total);

    const tables = body.getTables();
    if (tables.length > 0) {
      let servicesTable = null;
      for (let i = 0; i < tables.length; i++) {
        if (tables[i].getRow(0) && tables[i].getRow(0).getNumCells() > 1 && tables[i].getRow(0).getCell(1).getText().toUpperCase().includes('FUNÇÃO')) {
          servicesTable = tables[i];
          break;
        }
      }
      
      if (servicesTable && servicesTable.getNumRows() > 1) {
        const templateRow = servicesTable.getRow(1);
        
        items.forEach((item, index) => {
          const newRow = servicesTable.appendTableRow();
          const cellContents = [index + 1, item.cargo, item.escala, item.quantidade, item.valorUnitario, item.totalLinha];
          cellContents.forEach((content, i) => {
            const cell = newRow.appendTableCell();
            cell.setText(String(content));
            if (templateRow.getNumCells() > i) {
              cell.setAttributes(templateRow.getCell(i).getAttributes());
            }
          });
        });

        servicesTable.removeRow(1);
      }
    }

    doc.saveAndClose();
    const pdfBlob = doc.getAs('application/pdf');
    const pdfFile = outputFolder.createFile(pdfBlob).setName(newFileName + ".pdf");
    newFile.setTrashed(true);
    
    return { success: true, url: pdfFile.getUrl() };
  } catch (e) {
    Logger.log("Erro em generatePdfFromTemplate: " + e.message + " Stack: " + e.stack);
    return { success: false, message: 'Ocorreu um erro ao gerar o PDF: ' + e.message };
  }
}

function generateMiniReportPdf(reportData) {
  try {
    const { title, period, headers, rows, userRole } = reportData;
    const outputFolderId = userRole === 'Diretor' ? DIR_RELATORIOS_FOLDER_ID : PDF_GERADOS_FOLDER_ID;
    const outputFolder = DriveApp.getFolderById(outputFolderId);

    const templatesFolder = DriveApp.getFolderById(TEMPLATES_FOLDER_ID);
    const templateFiles = templatesFolder.getFilesByName('Relatorios Template');
     if (!templateFiles.hasNext()) {
      throw new Error("Template 'Relatorios Template' não encontrado na pasta de templates.");
    }
    const templateFile = templateFiles.next();
    
    const today = Utilities.formatDate(new Date(), "GMT-3", "dd_MM_yyyy");
    const docName = `Relatório - ${title} - ${today}`;

    const newFile = templateFile.makeCopy(docName, outputFolder);
    const doc = DocumentApp.openById(newFile.getId());
    const body = doc.getBody();

    body.replaceText('{{TITULO_RELATORIO}}', title.toUpperCase());
    body.replaceText('{{PERIODO}}', `Período do Filtro: ${period}`);

    const tables = body.getTables();
    if (tables.length > 0) {
        const reportTable = tables[0];
        if (reportTable.getNumRows() > 1) {
            const templateRow = reportTable.getRow(1);
            
            const header = reportTable.insertTableRow(1);
            headers.forEach(h => header.appendTableCell(h));
            reportTable.getRow(1).editAsText().setBold(true);

            rows.forEach(rowData => {
                const newRow = reportTable.appendTableRow();
                rowData.forEach((cellData, i) => {
                    const cell = newRow.appendTableCell();
                    cell.setText(String(cellData));
                    if(templateRow.getNumCells() > i) {
                      cell.setAttributes(templateRow.getCell(i).getAttributes());
                    }
                });
            });
            reportTable.removeRow(2);
        }
    }
    
    doc.saveAndClose();
    const pdfBlob = doc.getAs('application/pdf');
    const pdfFile = outputFolder.createFile(pdfBlob).setName(docName + '.pdf');
    newFile.setTrashed(true);

    return { success: true, url: pdfFile.getUrl() };
  } catch (e) {
    Logger.log("Erro em generateMiniReportPdf: " + e.message + " " + e.stack);
    return { success: false, message: 'Ocorreu um erro ao gerar o PDF do relatório.' };
  }
}