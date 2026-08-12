/* ===================== Utilidades ===================== */
async function carregarLogoBuffer() {
  const resp = await fetch("assets/logo.png");
  const buf = await resp.arrayBuffer();
  return buf;
}

async function carregarLogoDataUrl() {
  const buf = await carregarLogoBuffer();
  let binary = "";
  const bytes = new Uint8Array(buf);
  for (let i = 0; i < bytes.byteLength; i++) binary += String.fromCharCode(bytes[i]);
  return "data:image/png;base64," + btoa(binary);
}

function dataStrParaData(str) {
  const d = parseDataLocal(str);
  return d;
}

function nomeArquivoBase() {
  const dep = (state.departamento || "").replace(/[^\w\-]+/g, "_");
  const hoje = new Date().toISOString().slice(0, 10);
  return `Plano_de_Acao_${dep}_${hoje}`;
}

/* ===================== Excel ===================== */
async function gerarExcel() {
  const ExcelJSLib = window.ExcelJS;
  const wb = new ExcelJSLib.Workbook();
  wb.creator = "Sistema Plano de Ação";
  wb.created = new Date();

  const ws = wb.addWorksheet("Plano de Ação", {
    views: [{ state: "frozen", ySplit: 9 }]
  });

  // Larguras de coluna (iguais à planilha original)
  ws.columns = [
    { width: 2.71 },  // A
    { width: 16.29 }, // B
    { width: 18.29 }, // C
    { width: 58.57 }, // D
    { width: 78.14 }, // E
    { width: 42.86 }, // F
    { width: 30 },    // G
    { width: 17.86 }, // H
    { width: 17.86 }, // I
    { width: 16.86 }, // J
    { width: 24.71 }, // K
    { width: 20, hidden: true } // L (não usada)
  ];

  ws.getRow(1).height = 9.75;
  [2, 3, 4, 5].forEach(r => ws.getRow(r).height = 19.5);
  ws.getRow(6).height = 16.15;
  ws.getRow(7).height = 38.45;
  ws.getRow(8).height = 20.45;
  ws.getRow(9).height = 64.5;

  const AZUL_ESCURO = "FF0C5EA5";
  const AZUL_CLARO = "FF07B1E5";
  const AMARELO = "FFFFFF00";
  const VERMELHO = "FFFF0000";
  const VERDE = "FF00B050";
  const FONTE = "Calibri";

  function estiloCaixa(cell, { size = 12, bold = false, color = "FF000000", bg = null, align = "center", wrap = false, border = "thin" } = {}) {
    cell.font = { name: FONTE, size, bold, color: { argb: color } };
    if (bg) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
    cell.alignment = { horizontal: align, vertical: "middle", wrapText: wrap };
    if (border) {
      cell.border = {
        top: { style: border }, left: { style: border },
        bottom: { style: border }, right: { style: border }
      };
    }
  }

  // ===== Cabeçalho: logo + título + caixa de código =====
  ws.mergeCells("B2:E5");
  ws.mergeCells("F2:J5");
  estiloCaixa(ws.getCell("F2"), { size: 24, bold: true, wrap: true, border: "medium" });
  ws.getCell("F2").value = "PLANO DE AÇÃO\n(Gestão de Desempenho)";
  estiloCaixa(ws.getCell("B2"), { border: "medium" });

  ["K2", "K3", "K4", "K5"].forEach(c => estiloCaixa(ws.getCell(c), { size: 12, bold: c === "K2" || c === "K4", border: "medium" }));
  ws.getCell("K2").value = "CÓDIGO";
  ws.getCell("K3").value = state.formCodigo || "FORM.";
  ws.getCell("K4").value = "VERSÃO";
  ws.getCell("K5").value = state.versao || "00";

  try {
    const logoBuffer = await carregarLogoBuffer();
    const imageId = wb.addImage({ buffer: logoBuffer, extension: "png" });
    ws.addImage(imageId, {
      tl: { col: 1.1, row: 1.15 },
      ext: { width: 300, height: 82 }
    });
  } catch (e) {
    console.warn("Não foi possível inserir a logo no Excel:", e);
  }

  // ===== Linha 6: rótulos REUNIÃO / DIRETORIA / DEPARTAMENTO =====
  ws.mergeCells("B6:D6");
  ws.mergeCells("F6:F6");
  ws.mergeCells("G6:I6");
  ["B6", "F6", "G6"].forEach(c => estiloCaixa(ws.getCell(c), { size: 16, bold: true, color: "FFFFFFFF", bg: AZUL_ESCURO, border: "medium" }));
  ws.getCell("B6").value = "REUNIÃO";
  ws.getCell("F6").value = "DIRETORIA";
  ws.getCell("G6").value = "DEPARTAMENTO";
  ["A6", "E6", "J6", "K6"].forEach(c => estiloCaixa(ws.getCell(c), { bg: AZUL_ESCURO, border: "medium" }));

  // ===== Linha 7: valores selecionados =====
  ws.mergeCells("B7:D7");
  ws.mergeCells("F7:F7");
  ws.mergeCells("G7:I7");
  ["B7", "F7", "G7"].forEach(c => estiloCaixa(ws.getCell(c), { size: 14, bold: false, color: "FFFFFFFF", bg: AZUL_CLARO }));
  ws.getCell("B7").value = state.tipoReuniao;
  ws.getCell("F7").value = nomeDiretoriaCompleto(state.diretoria);
  ws.getCell("G7").value = state.departamento;
  ["A7", "E7", "J7", "K7"].forEach(c => estiloCaixa(ws.getCell(c), { bg: AZUL_CLARO }));

  // ===== Linha 8: título "PLANO DE AÇÃO" =====
  ws.mergeCells("B8:K8");
  estiloCaixa(ws.getCell("B8"), { size: 16, bold: true, color: "FFFFFFFF", bg: AZUL_ESCURO, border: "medium" });
  ws.getCell("B8").value = "PLANO DE AÇÃO";
  estiloCaixa(ws.getCell("A8"), { bg: AZUL_ESCURO, border: "medium" });

  // ===== Linha 9: cabeçalho da tabela =====
  const cabecalhos = {
    B9: "DATA REGISTRO", C9: "ÁREA/OBRA", D9: "ASSUNTO:", E9: "DETALHAMENTO AÇÃO",
    F9: "OBSERVAÇÕES \n(COMENTÁRIOS SOBRE O ANDAMENTO)", G9: "RESPONSÁVEIS",
    H9: "DATA PREVISTA TÉRMINO", I9: "DATA REPROGRAMADA", J9: "DATA CONCLUSÃO", K9: "STATUS"
  };
  Object.entries(cabecalhos).forEach(([coord, texto]) => {
    const cell = ws.getCell(coord);
    estiloCaixa(cell, { size: 12, bold: true, color: "FFFFFFFF", bg: AZUL_CLARO, wrap: true, border: "medium" });
    cell.value = texto;
  });
  estiloCaixa(ws.getCell("A9"), { bg: AZUL_CLARO, border: "medium" });

  // ===== Aba "Listas" (apoio aos menus suspensos) =====
  const wsListas = wb.addWorksheet("Listas");
  wsListas.getCell("B1").value = "Nome Diretoria";
  wsListas.getCell("C1").value = "Abreviação";
  DIRETORIAS.forEach((d, i) => {
    wsListas.getCell(`B${i + 2}`).value = d.nome;
    wsListas.getCell(`C${i + 2}`).value = d.abrev;
  });
  const colunaPorAbrev = { GG: "F", Engenharia: "G", "D.I": "H", Comercial: "I", Controladoria: "J", "J.C": "K" };
  Object.entries(colunaPorAbrev).forEach(([abrev, coluna]) => {
    wsListas.getCell(`${coluna}1`).value = abrev;
    (DEPARTAMENTOS_POR_DIRETORIA[abrev] || []).forEach((dep, i) => {
      wsListas.getCell(`${coluna}${i + 2}`).value = dep;
    });
  });
  wsListas.columns.forEach(c => c.width = 24);

  // Nomes definidos (listas nomeadas) para o menu suspenso em cascata de Departamento
  const nomesSanitizados = { GG: "DEPT_GG", Engenharia: "DEPT_Engenharia", "D.I": "DEPT_DI", Comercial: "DEPT_Comercial", Controladoria: "DEPT_Controladoria", "J.C": "DEPT_JC" };
  Object.entries(colunaPorAbrev).forEach(([abrev, coluna]) => {
    const qtd = (DEPARTAMENTOS_POR_DIRETORIA[abrev] || []).length;
    wb.definedNames.add(`Listas!$${coluna}$2:$${coluna}$${Math.max(qtd + 1, 2)}`, nomesSanitizados[abrev]);
  });

  // Célula auxiliar oculta: resolve o nome da lista de departamentos da diretoria escolhida
  ws.getCell("M7").value = { formula: 'CONCATENATE("DEPT_",SUBSTITUTE(VLOOKUP($F$7,Listas!$B$2:$C$7,2,FALSE),".",""))' };
  ws.getColumn("M").hidden = true;

  // ===== Validações (menus suspensos) =====
  ws.getCell("B7").dataValidation = {
    type: "list", allowBlank: true, formulae: ['"Alinhamento,Desempenho"']
  };
  ws.getCell("F7").dataValidation = {
    type: "list", allowBlank: true,
    formulae: [`"${DIRETORIAS.map(d => d.nome).join(",")}"`]
  };
  ws.getCell("G7").dataValidation = {
    type: "list", allowBlank: true, formulae: ["INDIRECT($M$7)"]
  };

  // ===== Linhas de dados =====
  const primeiraLinha = 10;
  state.linhas.forEach((linha, idx) => {
    const r = primeiraLinha + idx;
    const row = ws.getRow(r);
    row.height = 30;

    const setTexto = (col, valor, wrap = true) => {
      const cell = ws.getCell(`${col}${r}`);
      cell.value = valor || null;
      cell.font = { name: FONTE, size: 11 };
      cell.alignment = { horizontal: "left", vertical: "top", wrapText: wrap };
      cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    };
    const setData = (col, valorStr) => {
      const cell = ws.getCell(`${col}${r}`);
      const d = dataStrParaData(valorStr);
      cell.value = d || null;
      if (d) cell.numFmt = "dd/mm/yyyy";
      cell.font = { name: FONTE, size: 11 };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    };

    setData("B", linha.dataRegistro);
    setTexto("C", linha.areaObra, false);
    setTexto("D", linha.assunto);
    setTexto("E", linha.detalhamento);
    setTexto("F", linha.observacoes);
    setTexto("G", linha.responsaveis, false);
    setData("H", linha.dataPrevista);
    setData("I", linha.dataReprogramada);
    setData("J", linha.dataConclusao);

    const kCell = ws.getCell(`K${r}`);
    kCell.value = { formula: `IF(ISBLANK(E${r}),"",IF(ISBLANK(J${r}),IF(H${r}-TODAY()>=0,"Em Andamento","Atrasado"),"Concluído"))`, result: calcularStatus(linha) };
    kCell.font = { name: FONTE, size: 11, bold: true };
    kCell.alignment = { horizontal: "center", vertical: "middle" };
    kCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
  });

  const ultimaLinha = primeiraLinha + state.linhas.length - 1;

  // Menu suspenso de status (igual ao original, coexiste com a fórmula)
  ws.dataValidations.add(`K${primeiraLinha}:K${ultimaLinha}`, {
    type: "list", allowBlank: true, formulae: ['"Não Iniciado,Em andamento,Concluído"']
  });

  // Formatação condicional (mesmas cores do arquivo original)
  ws.addConditionalFormatting({
    ref: `K${primeiraLinha}:K${ultimaLinha}`,
    rules: [
      { type: "cellIs", operator: "equal", formulae: ['"Em Andamento"'], priority: 1, style: { font: { bold: true }, fill: { type: "pattern", pattern: "solid", fgColor: { argb: AMARELO } } } },
      { type: "cellIs", operator: "equal", formulae: ['"Atrasado"'], priority: 2, style: { font: { bold: true, color: { argb: "FFFFFFFF" } }, fill: { type: "pattern", pattern: "solid", fgColor: { argb: VERMELHO } } } },
      { type: "cellIs", operator: "equal", formulae: ['"Concluído"'], priority: 3, style: { font: { bold: true, color: { argb: "FFFFFFFF" } }, fill: { type: "pattern", pattern: "solid", fgColor: { argb: VERDE } } } }
    ]
  });

  ws.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 8 };

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  baixarBlob(blob, `${nomeArquivoBase()}.xlsx`);
}

/* ===================== PDF ===================== */
async function gerarPdf() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a3" });

  const margemEsq = 24;
  const margemDir = 24;
  const larguraUtil = doc.internal.pageSize.getWidth() - margemEsq - margemDir;

  const AZUL_ESCURO = [12, 94, 165];
  const AZUL_CLARO = [7, 177, 229];
  const AMARELO = [255, 255, 0];
  const VERMELHO = [255, 0, 0];
  const VERDE = [0, 176, 80];

  function desenharCabecalho() {
    let y = 20;
    const alturaCabecalho = 60;

    // logo
    try {
      doc.addImage(window.__logoDataUrl, "PNG", margemEsq + 4, y, 140, 38);
    } catch (e) { /* segue sem logo */ }

    // título
    doc.setFillColor(...AZUL_ESCURO);
    doc.rect(margemEsq + 160, y, larguraUtil - 160 - 130, alturaCabecalho, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(18);
    doc.text("PLANO DE AÇÃO", margemEsq + 160 + (larguraUtil - 160 - 130) / 2, y + 26, { align: "center" });
    doc.setFontSize(11);
    doc.setFont("helvetica", "normal");
    doc.text("(Gestão de Desempenho)", margemEsq + 160 + (larguraUtil - 160 - 130) / 2, y + 44, { align: "center" });

    // caixa código
    const xCodigo = margemEsq + larguraUtil - 130;
    doc.setDrawColor(0, 0, 0);
    doc.setTextColor(0, 0, 0);
    doc.setFontSize(9);
    const linhasCodigo = [["CÓDIGO", ""], [state.formCodigo || "FORM.", ""], ["VERSÃO", ""], [state.versao || "00", ""]];
    const alturaLinha = alturaCabecalho / 4;
    linhasCodigo.forEach((l, i) => {
      doc.rect(xCodigo, y + i * alturaLinha, 130, alturaLinha);
      doc.setFont("helvetica", i % 2 === 0 ? "bold" : "normal");
      doc.text(l[0], xCodigo + 65, y + i * alturaLinha + alturaLinha / 2 + 3, { align: "center" });
    });

    y += alturaCabecalho + 4;

    // barra reunião/diretoria/departamento
    const largTerco = larguraUtil / 3;
    doc.setFillColor(...AZUL_ESCURO);
    doc.rect(margemEsq, y, larguraUtil, 16, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(9);
    doc.text("REUNIÃO", margemEsq + largTerco / 2, y + 11, { align: "center" });
    doc.text("DIRETORIA", margemEsq + largTerco * 1.5, y + 11, { align: "center" });
    doc.text("DEPARTAMENTO", margemEsq + largTerco * 2.5, y + 11, { align: "center" });
    y += 16;

    doc.setFillColor(...AZUL_CLARO);
    doc.rect(margemEsq, y, larguraUtil, 18, "F");
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.text(state.tipoReuniao || "", margemEsq + largTerco / 2, y + 13, { align: "center" });
    doc.text(nomeDiretoriaCompleto(state.diretoria) || "", margemEsq + largTerco * 1.5, y + 13, { align: "center" });
    doc.text(state.departamento || "", margemEsq + largTerco * 2.5, y + 13, { align: "center" });
    y += 18 + 4;

    doc.setFillColor(...AZUL_ESCURO);
    doc.rect(margemEsq, y, larguraUtil, 18, "F");
    doc.setFontSize(11);
    doc.setFont("helvetica", "bold");
    doc.text("PLANO DE AÇÃO", margemEsq + larguraUtil / 2, y + 13, { align: "center" });
    y += 18 + 6;

    doc.setTextColor(0, 0, 0);
    return y;
  }

  const inicioTabelaY = desenharCabecalho();

  const colunas = [
    "DATA\nREGISTRO", "ÁREA/OBRA", "ASSUNTO", "DETALHAMENTO AÇÃO",
    "OBSERVAÇÕES", "RESPONSÁVEIS", "DATA PREVISTA\nTÉRMINO", "DATA\nREPROGRAMADA",
    "DATA\nCONCLUSÃO", "STATUS"
  ];

  function fmtData(str) {
    const d = parseDataLocal(str);
    if (!d) return "";
    return d.toLocaleDateString("pt-BR");
  }

  const linhas = state.linhas.map(l => {
    const status = calcularStatus(l);
    return [
      fmtData(l.dataRegistro), l.areaObra || "", l.assunto || "", l.detalhamento || "",
      l.observacoes || "", l.responsaveis || "", fmtData(l.dataPrevista), fmtData(l.dataReprogramada),
      fmtData(l.dataConclusao), status
    ];
  });

  doc.autoTable({
    startY: inicioTabelaY,
    margin: { left: margemEsq, right: margemDir, top: 40 },
    head: [colunas],
    body: linhas,
    styles: { font: "helvetica", fontSize: 7.5, cellPadding: 3, valign: "middle", lineColor: [90, 90, 90], lineWidth: 0.5, overflow: "linebreak" },
    headStyles: { fillColor: AZUL_CLARO, textColor: [255, 255, 255], fontStyle: "bold", halign: "center" },
    bodyStyles: { textColor: [0, 0, 0] },
    columnStyles: {
      0: { cellWidth: 55, halign: "center" },
      1: { cellWidth: 65 },
      2: { cellWidth: 95 },
      3: { cellWidth: "auto" },
      4: { cellWidth: "auto" },
      5: { cellWidth: 75 },
      6: { cellWidth: 62, halign: "center" },
      7: { cellWidth: 72, halign: "center" },
      8: { cellWidth: 62, halign: "center" },
      9: { cellWidth: 65, halign: "center", fontStyle: "bold" }
    },
    didParseCell: function (data) {
      if (data.section === "body" && data.column.index === 9) {
        const v = data.cell.raw;
        if (v === "Em Andamento") { data.cell.styles.fillColor = AMARELO; data.cell.styles.textColor = [0, 0, 0]; }
        else if (v === "Atrasado") { data.cell.styles.fillColor = VERMELHO; data.cell.styles.textColor = [255, 255, 255]; }
        else if (v === "Concluído") { data.cell.styles.fillColor = VERDE; data.cell.styles.textColor = [255, 255, 255]; }
      }
    },
    didDrawPage: function (data) {
      if (data.pageNumber > 1) {
        doc.setFillColor(...AZUL_ESCURO);
        doc.rect(margemEsq, 12, larguraUtil, 20, "F");
        doc.setTextColor(255, 255, 255);
        doc.setFont("helvetica", "bold");
        doc.setFontSize(11);
        doc.text(
          `PLANO DE AÇÃO — ${nomeDiretoriaCompleto(state.diretoria)} / ${state.departamento}`,
          margemEsq + larguraUtil / 2, 12 + 14, { align: "center" }
        );
      }
      const pageCount = doc.internal.getNumberOfPages();
      doc.setFontSize(8);
      doc.setTextColor(100);
      doc.text(
        `Página ${doc.internal.getCurrentPageInfo().pageNumber} de ${pageCount}`,
        doc.internal.pageSize.getWidth() - margemDir,
        doc.internal.pageSize.getHeight() - 10,
        { align: "right" }
      );
    }
  });

  doc.save(`${nomeArquivoBase()}.pdf`);
}

function baixarBlob(blob, nomeArquivo) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nomeArquivo;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

/* Pré-carrega a logo em base64 para uso no PDF */
carregarLogoDataUrl().then(url => { window.__logoDataUrl = url; }).catch(() => {});
