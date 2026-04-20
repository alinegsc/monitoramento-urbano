let dadosGlobais = [];

// Função para converter número do Excel em data legível
function excelSerialParaData(serial) {
    if (!serial || isNaN(serial)) return serial;
    const base = new Date(1899, 11, 30);
    const data = new Date(base.getTime() + serial * 24 * 60 * 60 * 1000);
    return data.toLocaleDateString("pt-BR");
}

// Carregar Excel
async function carregarExcel() {
    const response = await fetch("demandas.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const dados = XLSX.utils.sheet_to_json(sheet);

    dadosGlobais = dados;
    preencherFiltros(dados);
    mostrarCards(dados);
}

// Preencher filtros
function preencherFiltros(dados) {
    const bairros = [...new Set(dados.map(d => d["Bairro/Região"]))];
    const problemas = [...new Set(dados.map(d => d["Problema"]))];

    const filtroBairro = document.getElementById("filtro-bairro");
    const filtroProblema = document.getElementById("filtro-problema");

    bairros.forEach(b => {
        let opt = document.createElement("option");
        opt.value = b;
        opt.textContent = b;
        filtroBairro.appendChild(opt);
    });

    problemas.forEach(p => {
        let opt = document.createElement("option");
        opt.value = p;
        opt.textContent = p;
        filtroProblema.appendChild(opt);
    });
}

// Aplicar filtro
function aplicarFiltro() {
    const bairro = document.getElementById("filtro-bairro").value;
    const problema = document.getElementById("filtro-problema").value;

    let filtrados = dadosGlobais.filter(d => {
        return (bairro === "Todos" || d["Bairro/Região"] === bairro) &&
               (problema === "Todas" || d["Problema"] === problema);
    });

    mostrarCards(filtrados);
}

// Mostrar cards
function mostrarCards(dados) {
    const container = document.getElementById("cards-container");
    container.innerHTML = "";

    dados.forEach(d => {
        let card = document.createElement("div");
        card.className = "card";
        card.innerHTML = `
            <h3>${d["Bairro/Região"]}</h3>
            <p><strong>Problema:</strong> ${d["Problema"]}</p>
            <p><strong>Data:</strong> ${excelSerialParaData(d["Data da demanda"])}</p>
            <p><strong>Dias sem solução:</strong> ${d["Dias sem solução"]}</p>
            <p><strong>Status:</strong> ${d["Status"]}</p>
        `;
        container.appendChild(card);
    });
}

carregarExcel();
