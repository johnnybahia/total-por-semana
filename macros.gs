
// --- SUA FUNÇÃO ORIGINAL ---
// Esta função NÃO é chamada pelo Web App.
// Ela deve ser executada por um "Acionador" (Trigger) de tempo ou de edição.
function monitorarIMPORTRANGE() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Esta função ainda olha para "DADOS" e "RELATÓRIO GERAL DA PRODUÇÃO"
  // O Web App (doGet) olha para "RELATÓRIO GERAL DA PRODUÇÃO1"
  // Se precisar que esta função monitore a outra aba, você deve alterar os nomes aqui.
  var sheetBanco = ss.getSheetByName("DADOS"); 
  var sheetRelatorio = ss.getSheetByName("RELATÓRIO GERAL DA PRODUÇÃO");

  if (!sheetBanco || !sheetRelatorio) {
    Logger.log("Uma ou ambas as abas não foram encontradas (monitorarIMPORTRANGE).");
    return;
  }

  var range = sheetBanco.getRange("A1:l5000");
  var valoresAtuais = range.getValues();

  var cache = PropertiesService.getScriptProperties();
  var valoresAntigos = cache.getProperty("dados_antigos");

  if (valoresAntigos) {
    valoresAntigos = JSON.parse(valoresAntigos);
    var atualStr = JSON.stringify(valoresAtuais);
    var antigoStr = JSON.stringify(valoresAntigos);

    if (atualStr !== antigoStr) {
      var now = Utilities.formatDate(new Date(), "America/Fortaleza", "dd/MM/yyyy HH:mm:ss");
      sheetRelatorio.getRange("g2").setValue(now);
      cache.setProperty("dados_antigos", JSON.stringify(valoresAtuais));
      Logger.log("Alteração detectada! Horário atualizado.");
    } else {
      Logger.log("Nenhuma alteração detectada. G2 permanece o mesmo.");
    }
  } else {
    cache.setProperty("dados_antigos", JSON.stringify(valoresAtuais));
    Logger.log("Primeira execução: cache inicializado.");
  }
}
/**
 * MARFIM - Total de Pares por Semana
 *
 * Lê a aba "ESPELHO PARA CONSULTA":
 *   - Col A (A2:A5000): Clientes
 *   - Col D (D2:D5000): Tipo (filtra apenas registros que contêm "CM", ex: "125CM", "90CM")
 *   - Col E (E2:E5000): Quantidade de Pares
 *   - Col G (G2:G5000): Data
 *
 * Gera na aba "TOTAL DE PARES POR SEMANA" um resumo com:
 *   - Uma tabela por semana (Segunda a Domingo)
 *   - Total de pares por cliente dentro de cada semana
 *   - Total geral por semana
 */

function gerarTotalParesPorSemana() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Aba de origem ──────────────────────────────────────────────────────────
  const abaOrigem = ss.getSheetByName("ESPELHO PARA CONSULTA");
  if (!abaOrigem) {
    SpreadsheetApp.getUi().alert('Aba "ESPELHO PARA CONSULTA" não encontrada!');
    return;
  }

  // ── Aba de destino ─────────────────────────────────────────────────────────
  let abaDestino = ss.getSheetByName("TOTAL DE PARES POR SEMANA");
  if (!abaDestino) {
    abaDestino = ss.insertSheet("TOTAL DE PARES POR SEMANA");
  }

  // ── Leitura dos dados de origem ────────────────────────────────────────────
  const ultimaLinha = 5000;
  const primeiraLinha = 2;
  const totalLinhas = ultimaLinha - primeiraLinha + 1;

  const colCliente = abaOrigem.getRange(primeiraLinha, 1, totalLinhas, 1).getValues(); // Col A
  const colTipo    = abaOrigem.getRange(primeiraLinha, 4, totalLinhas, 1).getValues(); // Col D
  const colPares   = abaOrigem.getRange(primeiraLinha, 5, totalLinhas, 1).getValues(); // Col E
  const colData    = abaOrigem.getRange(primeiraLinha, 7, totalLinhas, 1).getValues(); // Col G

  // ── Estrutura: { "YYYY-MM-DD": { cliente: total, ... } } ──────────────────
  // Chave da semana = data da segunda-feira da semana
  const semanas = {};

  for (let i = 0; i < totalLinhas; i++) {
    const cliente = String(colCliente[i][0] || "").trim();
    const tipo    = String(colTipo[i][0]    || "").trim().toUpperCase();
    const pares   = Number(colPares[i][0]   || 0);
    const data    = colData[i][0];

    // Ignora linhas vazias ou sem CM (busca "CM" em qualquer posição: "125CM", "90CM", etc.)
    if (!cliente || !tipo.includes("CM") || !data || pares <= 0) continue;

    // Valida que é uma data válida
    // Usa getFullYear/Month/Date direto para evitar deslocamento por fuso horário
    let dataObj;
    if (data instanceof Date) {
      // Reconstrói a data usando os valores locais (ano/mês/dia) sem horário
      dataObj = new Date(data.getFullYear(), data.getMonth(), data.getDate());
    } else {
      dataObj = new Date(data);
      dataObj = new Date(dataObj.getFullYear(), dataObj.getMonth(), dataObj.getDate());
    }
    if (isNaN(dataObj.getTime())) continue;

    // Encontra a segunda-feira da semana
    const segunda = getSegundaDaSemana(dataObj);
    const chave   = formatarData(segunda); // "YYYY-MM-DD"

    if (!semanas[chave]) semanas[chave] = {};
    if (!semanas[chave][cliente]) semanas[chave][cliente] = 0;
    semanas[chave][cliente] += pares;
  }

  // ── Ordena semanas cronologicamente ───────────────────────────────────────
  const chavesOrdenadas = Object.keys(semanas).sort();

  if (chavesOrdenadas.length === 0) {
    SpreadsheetApp.getUi().alert("Nenhum dado com CM encontrado para gerar o relatório.");
    return;
  }

  // ── Monta o conteúdo para a aba de destino ────────────────────────────────
  abaDestino.clearContents();
  abaDestino.clearFormats();

  // Coleta todos os clientes únicos (para ordenação consistente)
  const todosClientes = new Set();
  chavesOrdenadas.forEach(chave => {
    Object.keys(semanas[chave]).forEach(c => todosClientes.add(c));
  });
  const clientesOrdenados = Array.from(todosClientes).sort();

  let linhaAtual = 1;

  // Título geral
  abaDestino.getRange(linhaAtual, 1).setValue("TOTAL DE PARES POR SEMANA (apenas CM)");
  abaDestino.getRange(linhaAtual, 1).setFontSize(13).setFontWeight("bold")
    .setBackground("#1a3a5c").setFontColor("#ffffff");
  abaDestino.getRange(linhaAtual, 1, 1, clientesOrdenados.length + 2)
    .merge().setHorizontalAlignment("center");
  linhaAtual += 2;

  // ── Uma tabela por semana ──────────────────────────────────────────────────
  chavesOrdenadas.forEach(chave => {
    const dadosSemana = semanas[chave];
    const segunda = new Date(chave + "T00:00:00");
    const domingo = new Date(segunda);
    domingo.setDate(domingo.getDate() + 6);

    // Cabeçalho da semana
    const labelSemana = `Semana: ${formatarDataBR(segunda)} a ${formatarDataBR(domingo)}`;
    const numColunas = clientesOrdenados.length + 2; // cliente + pares + total

    abaDestino.getRange(linhaAtual, 1).setValue(labelSemana);
    abaDestino.getRange(linhaAtual, 1, 1, numColunas)
      .merge()
      .setBackground("#2e6da4")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setFontSize(11)
      .setHorizontalAlignment("left");
    linhaAtual++;

    // Cabeçalho da tabela: "CLIENTE" | "TOTAL DE PARES"
    abaDestino.getRange(linhaAtual, 1).setValue("CLIENTE");
    abaDestino.getRange(linhaAtual, 2).setValue("TOTAL DE PARES");
    abaDestino.getRange(linhaAtual, 1, 1, 2)
      .setBackground("#d0e4f7")
      .setFontWeight("bold")
      .setBorder(true, true, true, true, true, true);
    linhaAtual++;

    // Linhas de cada cliente
    let totalGeral = 0;
    const clientesDaSemana = Object.keys(dadosSemana).sort();

    clientesDaSemana.forEach((cliente, idx) => {
      const pares = dadosSemana[cliente];
      totalGeral += pares;

      abaDestino.getRange(linhaAtual, 1).setValue(cliente);
      abaDestino.getRange(linhaAtual, 2).setValue(pares);

      // Zebra striping
      const bg = idx % 2 === 0 ? "#f5f9ff" : "#ffffff";
      abaDestino.getRange(linhaAtual, 1, 1, 2).setBackground(bg)
        .setBorder(false, true, false, true, false, false,
          "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
      linhaAtual++;
    });

    // Linha de total geral da semana
    abaDestino.getRange(linhaAtual, 1).setValue("TOTAL GERAL DA SEMANA");
    abaDestino.getRange(linhaAtual, 2).setValue(totalGeral);
    abaDestino.getRange(linhaAtual, 1, 1, 2)
      .setBackground("#1a3a5c")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setBorder(true, true, true, true, true, true);
    linhaAtual += 2; // Espaço entre semanas
  });

  // Ajusta largura das colunas
  abaDestino.setColumnWidth(1, 260);
  abaDestino.setColumnWidth(2, 160);

  // Congela primeira linha
  abaDestino.setFrozenRows(1);

  // Timestamp de atualização
  abaDestino.getRange(linhaAtual, 1)
    .setValue("Atualizado em: " + new Date().toLocaleString("pt-BR"));
  abaDestino.getRange(linhaAtual, 1).setFontColor("#888888").setFontStyle("italic");

  SpreadsheetApp.getUi().alert("✅ Relatório gerado com sucesso na aba \"TOTAL DE PARES POR SEMANA\"!");
}

// ── Helpers ────────────────────────────────────────────────────────────────

/** Retorna a segunda-feira da semana de uma data */
function getSegundaDaSemana(data) {
  const d = new Date(data);
  d.setHours(0, 0, 0, 0);
  const diaSemana = d.getDay(); // 0=Dom, 1=Seg, ..., 6=Sab
  // Ajuste: se domingo (0), volta 6 dias; senão, volta (diaSemana - 1) dias
  const diff = diaSemana === 0 ? -6 : 1 - diaSemana;
  d.setDate(d.getDate() + diff);
  return d;
}

/** Formata Date para "YYYY-MM-DD" (chave de ordenação) */
function formatarData(d) {
  const ano = d.getFullYear();
  const mes = String(d.getMonth() + 1).padStart(2, "0");
  const dia = String(d.getDate()).padStart(2, "0");
  return `${ano}-${mes}-${dia}`;
}

/** Formata Date para "DD/MM/YYYY" (exibição) */
function formatarDataBR(d) {
  const dia = String(d.getDate()).padStart(2, "0");
  const mes = String(d.getMonth() + 1).padStart(2, "0");
  const ano = d.getFullYear();
  return `${dia}/${mes}/${ano}`;
}

// ── Menu personalizado ─────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📦 Marfim")
    .addItem("Gerar Total de Pares por Semana", "gerarTotalParesPorSemana")
    .addToUi();
}
