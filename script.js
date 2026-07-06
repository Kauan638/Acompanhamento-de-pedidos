// =====================================
// CONFIGURAÇÃO — SEÇÕES E CATEGORIAS
// =====================================
// A lógica de agrupamento foi validada contra
// o arquivo real de consulta da API:
// - PEDIDOS GERAIS = todas as linhas, sem filtro
// - Pavilhão 1 e 2 se dividem por Sorter (SIM/NÃO)
// - Pavilhão 3 = Bebidas / Pavilhão Pereciveis =
//   Perecíveis, sem divisão por sorter
// - Categoria = campo "Status Carga"
// - R$ = soma de Valor / Vol = soma de Quantidade

const SECOES = [

    { id:"geral", titulo:"Pedidos Gerais", filtro: null },

    { id:"pav1_nao_sorter", titulo:"Pavilhão 1 - Não Sorter",
      filtro: r => r.pavilhao === "Pavilhão 1" && r.sorter === "NÃO" },

    { id:"pav1_sorter", titulo:"Pavilhão - 1 Sorter",
      filtro: r => r.pavilhao === "Pavilhão 1" && r.sorter === "SIM" },

    { id:"pav2_nao_sorter", titulo:"Pavilhão 2 - Não Sorter",
      filtro: r => r.pavilhao === "Pavilhão 2" && r.sorter === "NÃO" },

    { id:"pav2_sorter", titulo:"Pavilhão 2 - Sorter",
      filtro: r => r.pavilhao === "Pavilhão 2" && r.sorter === "SIM" },

    { id:"pav3_bebidas", titulo:"Pavilhão 3 - Bebidas",
      filtro: r => r.pavilhao === "Pavilhão 3" },

    { id:"pav_pereciveis", titulo:"Pavilhão Perecíveis",
      filtro: r => r.pavilhao === "Pavilhão Pereciveis" }

];

const CATEGORIAS = [
    "Roterizado - Não listado",
    "Roterizado - Listado",
    "Não Roterizado"
];

const CHAVE_STORAGE = "acompanhamentoPedidos_dias";

// =====================================
// INICIALIZAÇÃO
// =====================================

window.addEventListener("load", () => {

    const hoje = new Date();

    const iso =
    hoje.toISOString().slice(0,10);

    const campoData =
    document.getElementById("dataArquivo");

    if(campoData){
        campoData.value = iso;
    }

    document
    .getElementById("arquivoPedidos")
    ?.addEventListener("change", function(){

        const arquivo = this.files[0];

        document.getElementById("nomePedidos").innerText =
        arquivo ? arquivo.name : "Nenhum arquivo selecionado";

    });

    renderizarChipsDias();

    renderizarTabela();

});

// =====================================
// LOADING
// =====================================

function mostrarLoading(){
    const el = document.getElementById("loading");
    if(el) el.style.display = "flex";
}

function ocultarLoading(){
    const el = document.getElementById("loading");
    if(el) el.style.display = "none";
}

// =====================================
// HELPERS DE STORAGE
// =====================================

function obterDiasSalvos(){

    try{

        const bruto = localStorage.getItem(CHAVE_STORAGE);

        return bruto ? JSON.parse(bruto) : {};

    }
    catch(erro){

        console.error("Não consegui ler os dias salvos:", erro);

        return {};

    }

}

function salvarDias(dias){

    try{

        localStorage.setItem(
            CHAVE_STORAGE,
            JSON.stringify(dias)
        );

    }
    catch(erro){

        console.error("Não consegui salvar os dias:", erro);

        alert(
            "Não consegui salvar no navegador. O dia foi processado mas pode não persistir ao recarregar a página."
        );

    }

}

// =====================================
// HELPERS DE FORMATAÇÃO
// =====================================

function formatarMoeda(valor){

    return valor.toLocaleString(
        "pt-BR",
        {style:"currency",currency:"BRL"}
    );

}

function formatarNumero(valor){

    return valor.toLocaleString(
        "pt-BR",
        {minimumFractionDigits:2,maximumFractionDigits:2}
    );

}

// data guardada como "YYYY-MM-DD" (input type=date) —
// exibida como "DD/MM" igual à planilha original

function formatarRotuloDia(isoData){

    const [ano,mes,dia] = isoData.split("-");

    return `${dia}/${mes}`;

}

// =====================================
// DETECÇÃO DE COLUNAS
// (mesmo padrão dos outros painéis — tolera
// pequenas variações de nome/acentuação no
// cabeçalho do export)
// =====================================

function detectarColuna(objeto, candidatos){

    if(!objeto){
        return null;
    }

    const chaves = Object.keys(objeto);

    for(const candidato of candidatos){

        const exato =
        chaves.find(k => k.toLowerCase().trim() === candidato.toLowerCase());

        if(exato){
            return exato;
        }

    }

    for(const candidato of candidatos){

        const parcial =
        chaves.find(k => k.toLowerCase().includes(candidato.toLowerCase()));

        if(parcial){
            return parcial;
        }

    }

    return null;

}

// =====================================
// LEITURA DO XLSX (SheetJS)
// =====================================

function lerXLSX(arquivo){

    return new Promise((resolve, reject) => {

        const leitor = new FileReader();

        leitor.onload = (evento) => {

            try{

                const dados = new Uint8Array(evento.target.result);

                const workbook = XLSX.read(dados, {type:"array"});

                const primeiraAba = workbook.SheetNames[0];

                const planilha = workbook.Sheets[primeiraAba];

                const linhas = XLSX.utils.sheet_to_json(planilha, {defval:""});

                resolve(linhas);

            }
            catch(erro){

                reject(erro);

            }

        };

        leitor.onerror = (erro) => reject(erro);

        leitor.readAsArrayBuffer(arquivo);

    });

}

// =====================================
// PROCESSAMENTO PRINCIPAL
// =====================================

async function processarEAdicionarDia(){

    try{

        mostrarLoading();

        const arquivo =
        document.getElementById("arquivoPedidos")?.files[0];

        const dataIso =
        document.getElementById("dataArquivo")?.value;

        if(!arquivo){

            alert("Selecione o arquivo .xlsx da consulta.");

            ocultarLoading();

            return;

        }

        if(!dataIso){

            alert("Selecione o dia desse arquivo.");

            ocultarLoading();

            return;

        }

        const linhasBrutas = await lerXLSX(arquivo);

        if(!linhasBrutas.length){

            alert("O arquivo está vazio.");

            ocultarLoading();

            return;

        }

        const colPavilhao =
        detectarColuna(linhasBrutas[0], ["pavilhão","pavilhao","pavilão"]);

        const colSorter =
        detectarColuna(linhasBrutas[0], ["sorter"]);

        const colStatusCarga =
        detectarColuna(linhasBrutas[0], ["status carga"]);

        const colValor =
        detectarColuna(linhasBrutas[0], ["valor"]);

        const colQuantidade =
        detectarColuna(linhasBrutas[0], ["quantidade"]);

        if(!colPavilhao || !colSorter || !colStatusCarga || !colValor || !colQuantidade){

            console.log(
                "Colunas disponíveis:",
                Object.keys(linhasBrutas[0])
            );

            alert(
                "Não consegui identificar todas as colunas necessárias (Pavilhão, Sorter, Status Carga, Valor, Quantidade). Abra o console (F12) e me manda os nomes das colunas."
            );

            ocultarLoading();

            return;

        }

        const linhas = linhasBrutas.map(r => ({

            pavilhao: String(r[colPavilhao] || "").trim(),
            sorter: String(r[colSorter] || "").trim().toUpperCase(),
            statusCarga: String(r[colStatusCarga] || "").trim(),
            valor: Number(r[colValor]) || 0,
            quantidade: Number(r[colQuantidade]) || 0

        }));

        const resumoDia = calcularResumo(linhas);

        const dias = obterDiasSalvos();

        dias[dataIso] = resumoDia;

        salvarDias(dias);

        renderizarChipsDias();

        renderizarTabela();

        ocultarLoading();

        alert(
            `Dia ${formatarRotuloDia(dataIso)} processado e adicionado à tabela.`
        );

    }
    catch(erro){

        console.error(erro);

        ocultarLoading();

        alert(
            "Erro ao processar o arquivo:\n\n" +
            erro.message +
            "\n\n(detalhe técnico no console, F12)"
        );

    }

}

// calcula o resumo (por seção/categoria) de um
// conjunto de linhas já normalizadas

function calcularResumo(linhas){

    const resumo = {};

    SECOES.forEach(secao => {

        const linhasSecao =
        secao.filtro
        ? linhas.filter(secao.filtro)
        : linhas;

        const categorias = {};

        let totalValor = 0;

        let totalQtd = 0;

        CATEGORIAS.forEach(cat => {

            const linhasCat =
            linhasSecao.filter(r => r.statusCarga === cat);

            const valor =
            linhasCat.reduce((s,r) => s + r.valor, 0);

            const qtd =
            linhasCat.reduce((s,r) => s + r.quantidade, 0);

            categorias[cat] = {
                valor,
                qtd,
                temDados: linhasCat.length > 0
            };

            totalValor += valor;

            totalQtd += qtd;

        });

        resumo[secao.id] = {
            categorias,
            totalValor,
            totalQtd,
            temDados: linhasSecao.length > 0
        };

    });

    return resumo;

}

// =====================================
// CHIPS DOS DIAS CARREGADOS
// =====================================

function renderizarChipsDias(){

    const container =
    document.getElementById("diasChips");

    if(!container){
        return;
    }

    const dias = obterDiasSalvos();

    const datas =
    Object.keys(dias).sort();

    if(!datas.length){

        container.innerHTML =
        `<p style="color:#6b7280;">Nenhum dia carregado ainda.</p>`;

        return;

    }

    container.innerHTML =
    datas.map(data => `
        <div class="dia-chip">
            📅 ${formatarRotuloDia(data)}
            <button onclick="removerDia('${data}')" title="Remover este dia">✕</button>
        </div>
    `).join("");

}

function removerDia(data){

    if(!confirm(`Remover o dia ${formatarRotuloDia(data)} da tabela?`)){

        return;

    }

    const dias = obterDiasSalvos();

    delete dias[data];

    salvarDias(dias);

    renderizarChipsDias();

    renderizarTabela();

}

function limparTudo(){

    if(!confirm("Isso vai apagar TODOS os dias carregados. Confirma?")){

        return;

    }

    localStorage.removeItem(CHAVE_STORAGE);

    renderizarChipsDias();

    renderizarTabela();

}

// =====================================
// TABELA PRINCIPAL
// =====================================

function montarHtmlTabela(){

    const dias = obterDiasSalvos();

    const datas =
    Object.keys(dias).sort();

    if(!datas.length){

        return `<p style="text-align:center;color:#6b7280;padding:40px;">
        Processe um arquivo pra gerar a tabela.
        </p>`;

    }

    let html = `
    <table class="pedidos-table">

        <thead>
            <tr>
                <th>Tipo</th>
                <th>Categoria</th>
                ${datas.map(d => `<th>${formatarRotuloDia(d)}</th>`).join("")}
            </tr>
        </thead>

        <tbody>
    `;

    SECOES.forEach(secao => {

        html += `
            <tr class="linha-secao">
                <td colspan="${2 + datas.length}">${secao.titulo}</td>
            </tr>
        `;

        // linhas de R$

        CATEGORIAS.forEach(cat => {

            html += `<tr class="linha-rs">
                <td class="col-tipo">R$</td>
                <td class="col-categoria">Totais para ${cat}</td>
                ${datas.map(d => celulaValor(dias[d], secao.id, cat, "valor")).join("")}
            </tr>`;

        });

        html += `<tr class="linha-rs linha-total">
            <td class="col-tipo">R$</td>
            <td class="col-categoria">📊 Total Geral</td>
            ${datas.map(d => celulaTotal(dias[d], secao.id, "totalValor")).join("")}
        </tr>`;

        // linhas de Vol

        CATEGORIAS.forEach(cat => {

            html += `<tr class="linha-vol">
                <td class="col-tipo">Vol</td>
                <td class="col-categoria">Totais para ${cat}</td>
                ${datas.map(d => celulaValor(dias[d], secao.id, cat, "qtd")).join("")}
            </tr>`;

        });

        html += `<tr class="linha-vol linha-total">
            <td class="col-tipo">Vol</td>
            <td class="col-categoria">📊 Total Geral</td>
            ${datas.map(d => celulaTotal(dias[d], secao.id, "totalQtd")).join("")}
        </tr>`;

    });

    html += `
        </tbody>

    </table>
    `;

    return html;

}

function celulaValor(resumoDia, secaoId, categoria, campo){

    const secao = resumoDia?.[secaoId];

    const dadoCategoria = secao?.categorias?.[categoria];

    if(!dadoCategoria || !dadoCategoria.temDados){

        return `<td class="valor-vazio">-</td>`;

    }

    const valor = dadoCategoria[campo === "valor" ? "valor" : "qtd"];

    return `<td>${campo === "valor" ? formatarMoeda(valor) : formatarNumero(valor)}</td>`;

}

function celulaTotal(resumoDia, secaoId, campo){

    const secao = resumoDia?.[secaoId];

    if(!secao || !secao.temDados){

        return `<td class="valor-vazio">-</td>`;

    }

    const valor = secao[campo];

    return `<td>${campo === "totalValor" ? formatarMoeda(valor) : formatarNumero(valor)}</td>`;

}

function renderizarTabela(){

    const container =
    document.getElementById("tabelaContainer");

    if(!container){
        return;
    }

    container.innerHTML = montarHtmlTabela();

}

// =====================================
// IMPRIMIR
// =====================================

function imprimirTabela(){

    const dias = obterDiasSalvos();

    if(!Object.keys(dias).length){

        alert("Nenhum dia carregado pra imprimir.");

        return;

    }

    const janela = window.open("", "_blank");

    if(!janela){

        alert("Permita pop-ups para este site.");

        return;

    }

    const tabelaHtml = montarHtmlTabela();

    const html = `
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Acompanhamento de Pedidos</title>
<style>
@page{ size:A4 landscape; margin:10mm; }
*{ box-sizing:border-box; }
body{ font-family:Arial,Helvetica,sans-serif; color:#222; margin:0; }
h1{ text-align:center; color:#1e3a8a; font-size:18px; margin-bottom:14px; }
table{ width:100%; border-collapse:collapse; }
th,td{ border:1px solid #ccc; padding:6px 8px; font-size:11px; text-align:right; }
th{ background:#1e3a8a; color:#fff; text-align:center; }
td.col-tipo, td.col-categoria{ text-align:left; }
tr.linha-secao td{ background:#dbe4f3; font-weight:bold; text-transform:uppercase; }
tr.linha-total td{ background:#fef3c7; font-weight:bold; }
</style>
</head>
<body>
<h1>📦 Acompanhamento de Pedidos — Distribuição por Pavilhão (Valor e Volume)</h1>
${tabelaHtml}
</body>
</html>
    `;

    janela.document.open();

    janela.document.write(html);

    janela.document.close();

    setTimeout(() => {

        janela.focus();

        janela.print();

    }, 500);

}

// =====================================
// GERAR IMAGEM PARA WHATSAPP
// =====================================

function montarRelatorioImagem(){

    const container =
    document.getElementById("relatorioImagem");

    if(!container){
        return;
    }

    const dias = obterDiasSalvos();

    const datas =
    Object.keys(dias).sort();

    const agora =
    new Date().toLocaleString("pt-BR");

    const rotuloDias =
    datas.length
    ? (
        datas.length === 1
        ? formatarRotuloDia(datas[0])
        : `${formatarRotuloDia(datas[0])} a ${formatarRotuloDia(datas[datas.length - 1])}`
      )
    : "—";

    // KPIs de topo: soma da seção "Pedidos Gerais" (que já
    // contempla todas as linhas, sem filtro) em todos os
    // dias carregados.

    let valorTotalGeral = 0;

    let volumeTotalGeral = 0;

    datas.forEach(d => {

        const secaoGeral =
        dias[d]?.["geral"];

        if(secaoGeral?.temDados){

            valorTotalGeral += secaoGeral.totalValor || 0;

            volumeTotalGeral += secaoGeral.totalQtd || 0;

        }

    });

    container.innerHTML = `

    <div class="ri-topo-faixa"></div>

    <div class="ri-cabecalho">

        <div class="ri-titulo-bloco">
            <div class="ri-titulo">
                📦 Acompanhamento de Pedidos
            </div>
            <div class="ri-subtitulo">
                Relatório Executivo · Distribuição por Pavilhão
            </div>
        </div>

        <div class="ri-periodo-badge">
            <div class="ri-periodo-label">Período</div>
            <div class="ri-periodo-valor">${rotuloDias}</div>
        </div>

    </div>

    <div class="ri-kpis">

        <div class="ri-kpi">
            <div class="ri-kpi-label">Dias no Relatório</div>
            <div class="ri-kpi-valor">${datas.length}</div>
        </div>

        <div class="ri-kpi">
            <div class="ri-kpi-label">Valor Total Geral</div>
            <div class="ri-kpi-valor">${formatarMoeda(valorTotalGeral)}</div>
        </div>

        <div class="ri-kpi">
            <div class="ri-kpi-label">Volume Total Geral</div>
            <div class="ri-kpi-valor">${formatarNumero(volumeTotalGeral)}</div>
        </div>

    </div>

    <div class="pedidos-table-wrap">
        ${montarHtmlTabela()}
    </div>

    <div class="ri-rodape">
        <span>Gerado pelo Acompanhamento de Pedidos</span>
        <span>${agora}</span>
    </div>

    `;

}

async function gerarImagemRelatorio(){

    const dias = obterDiasSalvos();

    if(!Object.keys(dias).length){

        alert("Nenhum dia carregado pra gerar o relatório.");

        return;

    }

    montarRelatorioImagem();

    if(document.fonts && document.fonts.ready){

        await document.fonts.ready;

    }

    const elemento =
    document.getElementById("relatorioImagem");

    let canvas;

    try{

        canvas = await html2canvas(elemento, {

            backgroundColor: "#14181C",

            scale: 2

        });

    }
    catch(erro){

        console.error(erro);

        alert(
            "Não consegui gerar a imagem. Veja o console (F12) pra detalhes."
        );

        return;

    }

    canvas.toBlob(async blob => {

        if(!blob){

            alert("Falha ao gerar a imagem.");

            return;

        }

        const botao =
        document.getElementById("btnExportarWhatsapp");

        const rotuloOriginal =
        botao ? botao.innerHTML : null;

        function baixarBlob(){

            const link = document.createElement("a");

            link.href = URL.createObjectURL(blob);

            link.download =
            `acompanhamento_pedidos_${new Date().toISOString().slice(0,10)}.png`;

            link.click();

            setTimeout(
                () => URL.revokeObjectURL(link.href),
                5000
            );

        }

        if(navigator.clipboard && window.ClipboardItem){

            try{

                await navigator.clipboard.write([

                    new ClipboardItem({
                        "image/png": blob
                    })

                ]);

                if(botao){

                    botao.innerHTML =
                    "✅ Copiado! Cole no WhatsApp (Ctrl+V)";

                    setTimeout(()=>{

                        botao.innerHTML = rotuloOriginal;

                    }, 3500);

                }else{

                    alert(
                        "✅ Imagem copiada! Agora é só abrir a conversa no WhatsApp e colar (Ctrl+V)."
                    );

                }

            }
            catch(erro){

                console.error(erro);

                baixarBlob();

                alert(
                    "Seu navegador não permitiu copiar direto pro clipboard, então baixei a imagem — é só anexar ela no WhatsApp."
                );

            }

        }else{

            baixarBlob();

            alert(
                "Seu navegador não suporta copiar imagens direto. A imagem foi baixada — é só anexar ela no WhatsApp."
            );

        }

    }, "image/png");

}
