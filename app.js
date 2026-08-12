/* ===================== Estado ===================== */
const STORAGE_KEY = "planoDeAcao.v1";

const COLS = [
  "dataRegistro", "areaObra", "assunto", "detalhamento", "observacoes",
  "responsaveis", "dataPrevista", "dataReprogramada", "dataConclusao"
];

let state = {
  formCodigo: "",
  versao: "00",
  tipoReuniao: "Desempenho",
  diretoria: "Engenharia",
  departamento: "Obras",
  linhas: []
};

function novaLinha() {
  return {
    id: crypto.randomUUID ? crypto.randomUUID() : String(Date.now() + Math.random()),
    dataRegistro: "", areaObra: "", assunto: "", detalhamento: "", observacoes: "",
    responsaveis: "", dataPrevista: "", dataReprogramada: "", dataConclusao: ""
  };
}

function carregarEstado() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const salvo = JSON.parse(raw);
      state = Object.assign(state, salvo);
    }
  } catch (e) {
    console.warn("Não foi possível carregar dados salvos:", e);
  }
  if (!state.linhas || state.linhas.length === 0) {
    state.linhas = Array.from({ length: 15 }, novaLinha);
  }
}

function salvarEstado() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

/* ===================== Status (mesma lógica da fórmula do Excel) =====================
   =IF(ISBLANK(E),"",IF(ISBLANK(J),IF(H-TODAY()>=0,"Em Andamento","Atrasado"),"Concluído"))
   E = detalhamento | H = dataPrevista | J = dataConclusao
*/
function hojeSemHora() {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  return d;
}

function parseDataLocal(str) {
  if (!str) return null;
  const [y, m, d] = str.split("-").map(Number);
  if (!y || !m || !d) return null;
  return new Date(y, m - 1, d);
}

function calcularStatus(linha) {
  if (!linha.detalhamento || linha.detalhamento.trim() === "") return "";
  if (!linha.dataConclusao || linha.dataConclusao.trim() === "") {
    const prevista = parseDataLocal(linha.dataPrevista);
    const hoje = hojeSemHora();
    const diff = prevista ? (prevista - hoje) : -1; // sem data prevista = tratado como atrasado (igual à fórmula original)
    return diff >= 0 ? "Em Andamento" : "Atrasado";
  }
  return "Concluído";
}

function statusClasse(status) {
  if (status === "Em Andamento") return "em-andamento";
  if (status === "Atrasado") return "atrasado";
  if (status === "Concluído") return "concluido";
  return "";
}

/* ===================== Render ===================== */
function popularDiretorias() {
  const sel = document.getElementById("diretoria");
  sel.innerHTML = "";
  DIRETORIAS.forEach(d => {
    const opt = document.createElement("option");
    opt.value = d.abrev;
    opt.textContent = d.nome;
    sel.appendChild(opt);
  });
  sel.value = state.diretoria || DIRETORIAS[0].abrev;
  popularDepartamentos();
}

function popularDepartamentos() {
  const diretoriaAbrev = document.getElementById("diretoria").value;
  const sel = document.getElementById("departamento");
  const lista = DEPARTAMENTOS_POR_DIRETORIA[diretoriaAbrev] || [];
  sel.innerHTML = "";
  lista.forEach(depName => {
    const opt = document.createElement("option");
    opt.value = depName;
    opt.textContent = depName;
    sel.appendChild(opt);
  });
  if (lista.includes(state.departamento)) {
    sel.value = state.departamento;
  } else {
    state.departamento = sel.value;
  }
}

function nomeDiretoriaCompleto(abrev) {
  const d = DIRETORIAS.find(d => d.abrev === abrev);
  return d ? d.nome : abrev;
}

function renderCabecalho() {
  document.getElementById("tipoReuniao").value = state.tipoReuniao;
  document.getElementById("diretoria").value = state.diretoria;
  popularDepartamentos();
}

function renderTabela() {
  const tbody = document.getElementById("corpoTabela");
  tbody.innerHTML = "";
  state.linhas.forEach((linha) => {
    tbody.appendChild(criarLinhaDom(linha));
  });
  atualizarContador();
}

function criarCampo(tag, tipo, valor, oninput, textarea) {
  const el = document.createElement(textarea ? "textarea" : "input");
  if (!textarea) el.type = tipo || "text";
  el.value = valor || "";
  el.addEventListener("input", oninput);
  return el;
}

function criarLinhaDom(linha) {
  const tr = document.createElement("tr");
  tr.dataset.id = linha.id;

  const tdData = document.createElement("td");
  tdData.appendChild(criarCampo(null, "date", linha.dataRegistro, e => { linha.dataRegistro = e.target.value; salvarEstado(); }));
  tr.appendChild(tdData);

  const tdArea = document.createElement("td");
  tdArea.appendChild(criarCampo(null, "text", linha.areaObra, e => { linha.areaObra = e.target.value; salvarEstado(); }));
  tr.appendChild(tdArea);

  const tdAssunto = document.createElement("td");
  tdAssunto.appendChild(criarCampo(null, null, linha.assunto, e => { linha.assunto = e.target.value; salvarEstado(); }, true));
  tr.appendChild(tdAssunto);

  const tdDetalhe = document.createElement("td");
  const inputDetalhe = criarCampo(null, null, linha.detalhamento, e => {
    linha.detalhamento = e.target.value;
    salvarEstado();
    atualizarStatusLinha(tr, linha);
  }, true);
  tdDetalhe.appendChild(inputDetalhe);
  tr.appendChild(tdDetalhe);

  const tdObs = document.createElement("td");
  tdObs.appendChild(criarCampo(null, null, linha.observacoes, e => { linha.observacoes = e.target.value; salvarEstado(); }, true));
  tr.appendChild(tdObs);

  const tdResp = document.createElement("td");
  tdResp.appendChild(criarCampo(null, "text", linha.responsaveis, e => { linha.responsaveis = e.target.value; salvarEstado(); }));
  tr.appendChild(tdResp);

  const tdPrevista = document.createElement("td");
  tdPrevista.appendChild(criarCampo(null, "date", linha.dataPrevista, e => {
    linha.dataPrevista = e.target.value;
    salvarEstado();
    atualizarStatusLinha(tr, linha);
  }));
  tr.appendChild(tdPrevista);

  const tdReprog = document.createElement("td");
  tdReprog.appendChild(criarCampo(null, "date", linha.dataReprogramada, e => { linha.dataReprogramada = e.target.value; salvarEstado(); }));
  tr.appendChild(tdReprog);

  const tdConclusao = document.createElement("td");
  tdConclusao.appendChild(criarCampo(null, "date", linha.dataConclusao, e => {
    linha.dataConclusao = e.target.value;
    salvarEstado();
    atualizarStatusLinha(tr, linha);
  }));
  tr.appendChild(tdConclusao);

  const tdStatus = document.createElement("td");
  tdStatus.className = "status-cell";
  tr.appendChild(tdStatus);

  const tdAcoes = document.createElement("td");
  tdAcoes.className = "col-acoes";
  const btnRemover = document.createElement("button");
  btnRemover.textContent = "Remover";
  btnRemover.addEventListener("click", () => {
    if (state.linhas.length <= 1) {
      alert("É necessário manter ao menos uma linha.");
      return;
    }
    if (!confirm("Remover esta ação da lista?")) return;
    state.linhas = state.linhas.filter(l => l.id !== linha.id);
    salvarEstado();
    renderTabela();
  });
  tdAcoes.appendChild(btnRemover);
  tr.appendChild(tdAcoes);

  atualizarStatusLinha(tr, linha);
  return tr;
}

function atualizarStatusLinha(tr, linha) {
  const status = calcularStatus(linha);
  const td = tr.querySelector("td.status-cell");
  td.textContent = status;
  td.className = "status-cell " + statusClasse(status);
}

function atualizarContador() {
  document.getElementById("contadorLinhas").textContent =
    `${state.linhas.length} ação(ões) na lista`;
}

/* ===================== Eventos de cabeçalho ===================== */
function ligarEventosCabecalho() {
  document.getElementById("tipoReuniao").addEventListener("change", e => {
    state.tipoReuniao = e.target.value;
    salvarEstado();
  });
  document.getElementById("diretoria").addEventListener("change", e => {
    state.diretoria = e.target.value;
    popularDepartamentos();
    salvarEstado();
  });
  document.getElementById("departamento").addEventListener("change", e => {
    state.departamento = e.target.value;
    salvarEstado();
  });
}

/* ===================== Adicionar / Limpar ===================== */
document.getElementById("btnAddLinha").addEventListener("click", () => {
  state.linhas.push(novaLinha());
  salvarEstado();
  renderTabela();
  const linhas = document.querySelectorAll("#corpoTabela tr");
  const ultima = linhas[linhas.length - 1];
  if (ultima) ultima.scrollIntoView({ behavior: "smooth", block: "center" });
});

document.getElementById("btnLimpar").addEventListener("click", () => {
  if (!confirm("Isso vai apagar TODAS as ações preenchidas e não pode ser desfeito. Continuar?")) return;
  state.linhas = Array.from({ length: 15 }, novaLinha);
  salvarEstado();
  renderTabela();
});

/* ===================== Inicialização ===================== */
carregarEstado();
popularDiretorias();
renderCabecalho();
renderTabela();
ligarEventosCabecalho();

/* ===================== Exportação ===================== */
document.getElementById("btnExportXlsx").addEventListener("click", () => {
  gerarExcel().catch(err => {
    console.error(err);
    alert("Erro ao gerar a planilha: " + err.message);
  });
});
document.getElementById("btnExportPdf").addEventListener("click", () => {
  gerarPdf().catch(err => {
    console.error(err);
    alert("Erro ao gerar o PDF: " + err.message);
  });
});
