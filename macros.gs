
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

// ── Configuração da planilha de produção diária ────────────────────────────
// ID extraído da URL da planilha de produção
const ID_PLANILHA_PRODUCAO = "1ahJCmhBEWYZhJsA-m7g_q_v-eJORDu4SXbYb2yrpHVA";

// Mapeamento de nomes de meses em PT para índice (0-11)
const MESES_PT = {
  "JANEIRO": 0, "FEVEREIRO": 1, "MARÇO": 2, "MARCO": 2,
  "ABRIL": 3, "MAIO": 4, "JUNHO": 5,
  "JULHO": 6, "AGOSTO": 7, "SETEMBRO": 8,
  "OUTUBRO": 9, "NOVEMBRO": 10, "DEZEMBRO": 11
};

/**
 * Lê a planilha de produção diária e retorna um objeto
 * { "YYYY-MM-DD": totalProduzido } onde a chave é a segunda-feira da semana.
 */
function lerProducaoPorSemana() {
  const producaoPorSemana = {};

  let ssProducao;
  try {
    ssProducao = SpreadsheetApp.openById(ID_PLANILHA_PRODUCAO);
  } catch (e) {
    Logger.log("Erro ao abrir planilha de produção: " + e.message);
    return producaoPorSemana;
  }

  const abas = ssProducao.getSheets();

  abas.forEach(aba => {
    const nomeAba = aba.getName().trim().toUpperCase();

    // Espera formato "MÊS AAAA" ex: "ABRIL 2026", "MARÇO 2026"
    const match = nomeAba.match(/^([A-ZÇÃÕÁÉÍÓÚÂÊÎÔÛÀÈÌÒÙÜ]+)\s+(\d{4})$/);
    if (!match) return;

    const nomeMes = match[1];
    const ano = parseInt(match[2]);
    const mesIdx = MESES_PT[nomeMes];
    if (mesIdx === undefined) return;

    const lastRow = aba.getLastRow();
    if (lastRow < 3) return;

    const numLinhas = lastRow - 2; // dados a partir da linha 3
    const colDias   = aba.getRange(3, 1,  numLinhas, 1).getValues(); // Col A
    const colTotais = aba.getRange(3, 22, numLinhas, 1).getValues(); // Col V

    for (let i = 0; i < numLinhas; i++) {
      const diaVal   = colDias[i][0];
      const totalDia = Number(colTotais[i][0] || 0);

      if (!diaVal || isNaN(totalDia) || totalDia <= 0) continue;

      // Col A pode ser objeto Date ou número/texto com o dia
      let dia;
      if (diaVal instanceof Date) {
        dia = diaVal.getDate();
      } else {
        dia = parseInt(String(diaVal).trim());
      }
      if (isNaN(dia) || dia < 1 || dia > 31) continue;

      const dataObj = new Date(ano, mesIdx, dia);
      if (isNaN(dataObj.getTime())) continue;

      const segunda = getSegundaDaSemana(dataObj);
      const chave   = formatarData(segunda);

      producaoPorSemana[chave] = (producaoPorSemana[chave] || 0) + totalDia;
    }
  });

  return producaoPorSemana;
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
 * Lê também a planilha de produção diária para comparar produção x pedidos.
 *
 * Gera na aba "TOTAL DE PARES POR SEMANA" um resumo com:
 *   - Uma tabela por semana (Segunda a Domingo)
 *   - Total de pares por cliente dentro de cada semana
 *   - Total de pedidos, produção realizada e saldo por semana
 *
 * @param {boolean} silencioso - se true, não exibe alertas (usado pelo acionador automático)
 */
function gerarTotalParesPorSemana(silencioso) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Aba de origem ──────────────────────────────────────────────────────────
  const abaOrigem = ss.getSheetByName("ESPELHO PARA CONSULTA");
  if (!abaOrigem) {
    if (!silencioso) SpreadsheetApp.getUi().alert('Aba "ESPELHO PARA CONSULTA" não encontrada!');
    return;
  }

  // ── Aba de destino ─────────────────────────────────────────────────────────
  let abaDestino = ss.getSheetByName("TOTAL DE PARES POR SEMANA");
  if (!abaDestino) {
    abaDestino = ss.insertSheet("TOTAL DE PARES POR SEMANA");
  }

  // ── Leitura dos dados de origem ────────────────────────────────────────────
  const ultimaLinha  = 5000;
  const primeiraLinha = 2;
  const totalLinhas  = ultimaLinha - primeiraLinha + 1;

  const colCliente = abaOrigem.getRange(primeiraLinha, 1, totalLinhas, 1).getValues(); // Col A
  const colTipo    = abaOrigem.getRange(primeiraLinha, 4, totalLinhas, 1).getValues(); // Col D
  const colPares   = abaOrigem.getRange(primeiraLinha, 5, totalLinhas, 1).getValues(); // Col E
  const colData    = abaOrigem.getRange(primeiraLinha, 7, totalLinhas, 1).getValues(); // Col G

  // ── Estrutura: { "YYYY-MM-DD": { cliente: total, ... } } ──────────────────
  const semanas = {};

  for (let i = 0; i < totalLinhas; i++) {
    const cliente = String(colCliente[i][0] || "").trim();
    const tipo    = String(colTipo[i][0]    || "").trim().toUpperCase();
    const pares   = Number(colPares[i][0]   || 0);
    const data    = colData[i][0];

    if (!cliente || !tipo.includes("CM") || !data || isNaN(pares) || pares <= 0) continue;

    let dataObj;
    if (data instanceof Date) {
      dataObj = new Date(data.getFullYear(), data.getMonth(), data.getDate());
    } else {
      const str = String(data).trim();
      const partesBR = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (partesBR) {
        dataObj = new Date(Number(partesBR[3]), Number(partesBR[2]) - 1, Number(partesBR[1]));
      } else {
        dataObj = new Date(str);
        dataObj = new Date(dataObj.getFullYear(), dataObj.getMonth(), dataObj.getDate());
      }
    }
    if (isNaN(dataObj.getTime())) continue;

    const segunda = getSegundaDaSemana(dataObj);
    const chave   = formatarData(segunda);

    if (!semanas[chave]) semanas[chave] = {};
    if (!semanas[chave][cliente]) semanas[chave][cliente] = 0;
    semanas[chave][cliente] += pares;
  }

  // ── Ordena semanas cronologicamente ───────────────────────────────────────
  const chavesOrdenadas = Object.keys(semanas).sort();

  if (chavesOrdenadas.length === 0) {
    if (!silencioso) SpreadsheetApp.getUi().alert("Nenhum dado com CM encontrado para gerar o relatório.");
    return;
  }

  // ── Lê produção diária da outra planilha ──────────────────────────────────
  const producaoPorSemana = lerProducaoPorSemana();

  // ── Cache acumulado de pedidos por semana ─────────────────────────────────
  // O total de pedidos nunca decresce: se um pedido sair da origem, o maior
  // valor já registrado é mantido. Novos pedidos são somados normalmente.
  const props = PropertiesService.getScriptProperties();
  const cachePedidos = JSON.parse(props.getProperty("cache_pedidos") || "{}");

  // Para cada semana calculada, aplica Math.max(cache, novo)
  Object.keys(semanas).forEach(chave => {
    const totalNovo = Object.values(semanas[chave]).reduce((s, v) => s + v, 0);
    const totalCached = cachePedidos[chave] || 0;
    if (totalNovo > totalCached) {
      cachePedidos[chave] = totalNovo;
    }
    // Substitui os valores individuais dos clientes pelo proporcional ao total final,
    // mas mantém os clientes como estão — apenas o total da linha final é protegido.
  });
  props.setProperty("cache_pedidos", JSON.stringify(cachePedidos));

  // ── Monta o conteúdo para a aba de destino ────────────────────────────────
  abaDestino.clearContents();
  abaDestino.clearFormats();

  let linhaAtual = 1;

  // Título geral
  abaDestino.getRange(linhaAtual, 1).setValue("TOTAL DE PARES POR SEMANA (apenas CM)");
  abaDestino.getRange(linhaAtual, 1)
    .setFontSize(20).setFontWeight("bold")
    .setBackground("#1a3a5c").setFontColor("#ffffff");
  abaDestino.getRange(linhaAtual, 1, 1, 3)
    .merge().setHorizontalAlignment("center");
  linhaAtual += 2;

  // ── Uma tabela por semana ──────────────────────────────────────────────────
  chavesOrdenadas.forEach(chave => {
    const dadosSemana = semanas[chave];
    const segunda = new Date(chave + "T00:00:00");
    const domingo = new Date(segunda);
    domingo.setDate(domingo.getDate() + 6);

    const labelSemana = `Semana: ${formatarDataBR(segunda)} a ${formatarDataBR(domingo)}`;

    // Cabeçalho da semana
    abaDestino.getRange(linhaAtual, 1).setValue(labelSemana);
    abaDestino.getRange(linhaAtual, 1, 1, 3)
      .merge()
      .setBackground("#2e6da4")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setFontSize(20)
      .setHorizontalAlignment("left");
    linhaAtual++;

    // Cabeçalho da tabela
    abaDestino.getRange(linhaAtual, 1).setValue("CLIENTE");
    abaDestino.getRange(linhaAtual, 2).setValue("TOTAL DE PARES");
    abaDestino.getRange(linhaAtual, 1, 1, 2)
      .setBackground("#d0e4f7")
      .setFontWeight("bold")
      .setBorder(true, true, true, true, true, true);
    linhaAtual++;

    // Linhas de cada cliente
    let totalPedidos = 0;
    const clientesDaSemana = Object.keys(dadosSemana).sort();

    clientesDaSemana.forEach((cliente, idx) => {
      const pares = dadosSemana[cliente];
      totalPedidos += pares;

      abaDestino.getRange(linhaAtual, 1).setValue(cliente);
      abaDestino.getRange(linhaAtual, 2).setValue(pares);

      const bg = idx % 2 === 0 ? "#f5f9ff" : "#ffffff";
      abaDestino.getRange(linhaAtual, 1, 1, 2).setBackground(bg)
        .setBorder(false, true, false, true, false, false,
          "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
      linhaAtual++;
    });

    // ── Linha: TOTAL DE PEDIDOS (valor protegido: nunca decresce) ───────────
    const totalPedidosFinal = cachePedidos[chave] || totalPedidos;
    abaDestino.getRange(linhaAtual, 1).setValue("TOTAL DE PEDIDOS");
    abaDestino.getRange(linhaAtual, 2).setValue(totalPedidosFinal);
    abaDestino.getRange(linhaAtual, 1, 1, 2)
      .setBackground("#1a3a5c")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setBorder(true, true, true, true, true, true);
    linhaAtual++;

    // ── Linha: PRODUÇÃO REALIZADA ────────────────────────────────────────────
    const producaoSemana = producaoPorSemana[chave] || 0;
    abaDestino.getRange(linhaAtual, 1).setValue("PRODUÇÃO REALIZADA");
    abaDestino.getRange(linhaAtual, 2).setValue(producaoSemana);
    abaDestino.getRange(linhaAtual, 1, 1, 2)
      .setBackground("#1e6b3c")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setBorder(true, true, true, true, true, true);
    linhaAtual++;

    // ── Linha: SALDO (produção - pedidos) ────────────────────────────────────
    const saldo = producaoSemana - totalPedidosFinal;
    const bgSaldo  = saldo >= 0 ? "#27ae60" : "#c0392b"; // verde ou vermelho
    const labelSaldo = saldo >= 0
      ? `SALDO: +${saldo} (em dia)`
      : `SALDO: ${saldo} (em atraso)`;
    abaDestino.getRange(linhaAtual, 1).setValue(labelSaldo);
    abaDestino.getRange(linhaAtual, 2).setValue(saldo);
    abaDestino.getRange(linhaAtual, 1, 1, 2)
      .setBackground(bgSaldo)
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setBorder(true, true, true, true, true, true);
    linhaAtual += 2; // Espaço entre semanas
  });

  // Ajusta largura das colunas
  abaDestino.setColumnWidth(1, 280);
  abaDestino.setColumnWidth(2, 180);

  // Aplica fonte 20 em todo o conteúdo gerado
  abaDestino.getDataRange().setFontSize(20);

  // Congela primeira linha
  abaDestino.setFrozenRows(1);

  // Timestamp de atualização
  abaDestino.getRange(linhaAtual, 1)
    .setValue("Atualizado em: " + new Date().toLocaleString("pt-BR"));
  abaDestino.getRange(linhaAtual, 1).setFontColor("#888888").setFontStyle("italic");

  if (!silencioso) {
    SpreadsheetApp.getUi().alert("✅ Relatório gerado com sucesso na aba \"TOTAL DE PARES POR SEMANA\"!");
  }
}

/**
 * Função chamada pelo acionador automático diário.
 * Chama gerarTotalParesPorSemana em modo silencioso (sem alertas).
 */
function atualizarRelatorioAutomatico() {
  gerarTotalParesPorSemana(true);
}

/**
 * Cria (ou recria) o acionador diário automático às 6h (horário de Fortaleza).
 * Execute esta função UMA VEZ manualmente para ativar as atualizações automáticas.
 */
function criarAcionadorDiario() {
  // Remove acionadores anteriores para evitar duplicatas
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "atualizarRelatorioAutomatico") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("atualizarRelatorioAutomatico")
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .inTimezone("America/Fortaleza")
    .create();

  SpreadsheetApp.getUi().alert(
    "✅ Acionador diário criado!\n\nO relatório será atualizado automaticamente todos os dias às 6h (horário de Fortaleza)."
  );
}

/**
 * Remove o acionador diário automático.
 */
function removerAcionadorDiario() {
  let removidos = 0;
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "atualizarRelatorioAutomatico") {
      ScriptApp.deleteTrigger(trigger);
      removidos++;
    }
  });
  SpreadsheetApp.getUi().alert(
    removidos > 0
      ? "✅ Acionador diário removido com sucesso."
      : "ℹ️ Nenhum acionador diário encontrado."
  );
}

// ── Helpers ────────────────────────────────────────────────────────────────

/** Retorna a segunda-feira da semana de uma data */
function getSegundaDaSemana(data) {
  const d = new Date(data);
  d.setHours(0, 0, 0, 0);
  const diaSemana = d.getDay(); // 0=Dom, 1=Seg, ..., 6=Sab
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

/**
 * Zera o cache de totais de pedidos.
 * Use apenas se quiser recalcular tudo do zero a partir da origem.
 */
function zerarCachePedidos() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "Zerar cache de pedidos?",
    "Isso fará o relatório recalcular os totais de pedidos do zero na próxima execução.\n\nTotais que já saíram da origem serão perdidos. Deseja continuar?",
    ui.ButtonSet.YES_NO
  );
  if (resp === ui.Button.YES) {
    PropertiesService.getScriptProperties().deleteProperty("cache_pedidos");
    ui.alert("✅ Cache zerado. Na próxima geração os totais serão recalculados.");
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📦 Marfim")
    .addItem("Gerar Total de Pares por Semana", "gerarTotalParesPorSemana")
    .addSeparator()
    .addItem("⏰ Ativar atualização diária automática", "criarAcionadorDiario")
    .addItem("🗑️ Desativar atualização diária", "removerAcionadorDiario")
    .addSeparator()
    .addItem("🔄 Zerar cache de pedidos (recalcular do zero)", "zerarCachePedidos")
    .addToUi();
}
