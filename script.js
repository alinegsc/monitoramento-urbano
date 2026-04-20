let dadosGlobais = []; // guarda todos os dados do Excel

// Função para carregar o Excel
async function carregarExcel() {
    const response = await fetch("demandas parelheiros.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    const primeiraAba = workbook.SheetNames[0];
    const dados = XLSX.utils.sheet_to_json(workbook.Sheets[primeiraAba]);

    dadosGlobais = dados; // salva para usar nos filtros

    preencherFiltros(dados);
    mostrarCards(dados);
}

// Preenche os filtros com opções únicas
function preencherFiltros(dados) {
    const bairros = [...new Set(dados.map(item => item["Bairro/Região"]))];
    const problemas = [...new Set(dados.map(item => item["Problema"]))];

    const filtroBairro = document.getElementById("filtro-bairro");
    const filtroProblema = document.getElementById("filtro-problema");

    bairros.forEach(b => {
        const opt = document.createElement("option");
        opt.value = b;
        opt.textContent = b;
        filtroBairro.appendChild(opt);
    });

    problemas.forEach(p => {
        const opt = document.createElement("option");
        opt.value = p;
        opt.textContent = p;
        filtroProblema.appendChild(opt);
    });
}

// Aplica os filtros
function aplicarFiltro() {
    const bairroSelecionado = document.getElementById("filtro-bairro").value;
    const problemaSelecionado = document.getElementById("filtro-problema").value;

    let filtrados = dadosGlobais;

    if (bairroSelecionado !== "Todos") {
        filtrados = filtrados.filter(item => item["Bairro/Região"] === bairroSelecionado);
    }

    if (problemaSelecionado !== "Todas") {
        filtrados = filtrados.filter(item => item["Problema"] === problemaSelecionado);
    }

    mostrarCards(filtrados);
}

// Função para criar os cards
function mostrarCards(dados) {
    const container = document.getElementById("cards-container");
    container.innerHTML = "";

    dados.forEach(item => {
        const card = document.createElement("div");
        card.className = "card";

        card.innerHTML = `
            <p><i class="fa-solid fa-location-dot"></i> Bairro: <strong>${item["Bairro/Região"]}</strong></p>
            <p><i class="fa-solid fa-triangle-exclamation"></i> Problema: ${item["Problema"]}</p>
            <p><i class="fa-solid fa-calendar"></i> Data: ${item["Data da demanda"]}</p>
            <p><i class="fa-solid fa-hourglass-half"></i> ${item["Dias sem solução"] || "N/A"} dias sem solução</p>
            <span class="status ${corStatus(item["Status"])}">Status: ${item["Status"]}</span>
        `;

        container.appendChild(card);
    });
}

// Define cor do status
function corStatus(status) {
    if (!status) return "vermelho";
    status = status.toLowerCase();
    if (status.includes("pendente")) return "vermelho";
    if (status.includes("andamento")) return "laranja";
    if (status.includes("resolvido")) return "verde";
    return "vermelho";
}

// Carrega ao abrir
carregarExcel();
