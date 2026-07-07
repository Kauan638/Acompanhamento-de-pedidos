// ========================================================
// ========================================================
// SINCRONIZAÇÃO AUTOMÁTICA — File System Access API
//
// Conecta a subpasta "Acompanhamento de Pedidos" (dentro da
// pasta mestre) uma única vez. A partir daí, detecta sozinho
// o arquivo da Consulta da API (extensão .xlsx / .xls) e
// reprocessa automaticamente sempre que ele for salvo/
// atualizado no disco — usando a DATA DE HOJE como o "dia"
// desse arquivo (mesmo padrão do campo de data, que já vem
// pré-preenchido com hoje por padrão).
//
// Reaproveita 100% da lógica já existente no projeto através
// de processarArquivoEData(arquivo, dataIso) — a mesma função
// que o botão manual "Processar e Adicionar Dia" chama por
// baixo dos panos (extraída do script.js original só pra
// não depender de alert()/DOM, sem duplicar nenhuma regra).
//
// IMPORTANTE: como esse projeto acumula um dia por vez, a
// sincronização assume que o arquivo na pasta é sempre a
// consulta DO DIA ATUAL. Se você atualizar o arquivo várias
// vezes no mesmo dia, ele só substitui a coluna daquele dia
// (não duplica).
// ========================================================
// ========================================================

const SYNC_DB_NAME = "acompanhamento-pedidos-sync-db";
const SYNC_STORE_NAME = "handles";
const SYNC_HANDLE_KEY = "pastaAcompanhamento";
const SYNC_INTERVALO_MS = 5000; // checa a cada 5s

let syncDirHandle = null;
let syncArquivoPedidosHandle = null;
let syncLastModifiedPedidos = 0;
let syncIntervalId = null;

// ---------- IndexedDB: persistir o handle da pasta ----------

function syncAbrirDB(){

    return new Promise((resolve, reject)=>{

        const req = indexedDB.open(SYNC_DB_NAME, 1);

        req.onupgradeneeded = ()=>
        req.result.createObjectStore(SYNC_STORE_NAME);

        req.onsuccess = ()=> resolve(req.result);

        req.onerror = ()=> reject(req.error);

    });

}

async function syncSalvarHandle(handle){

    const db = await syncAbrirDB();

    return new Promise((resolve, reject)=>{

        const tx = db.transaction(SYNC_STORE_NAME, "readwrite");

        tx.objectStore(SYNC_STORE_NAME).put(handle, SYNC_HANDLE_KEY);

        tx.oncomplete = resolve;

        tx.onerror = ()=> reject(tx.error);

    });

}

async function syncCarregarHandle(){

    const db = await syncAbrirDB();

    return new Promise((resolve, reject)=>{

        const tx = db.transaction(SYNC_STORE_NAME, "readonly");

        const req = tx.objectStore(SYNC_STORE_NAME).get(SYNC_HANDLE_KEY);

        req.onsuccess = ()=> resolve(req.result || null);

        req.onerror = ()=> reject(req.error);

    });

}

async function syncLimparHandle(){

    const db = await syncAbrirDB();

    const tx = db.transaction(SYNC_STORE_NAME, "readwrite");

    tx.objectStore(SYNC_STORE_NAME).delete(SYNC_HANDLE_KEY);

}

async function syncGarantirPermissao(handle){

    const opcoes = { mode: "read" };

    if((await handle.queryPermission(opcoes)) === "granted") return true;

    if((await handle.requestPermission(opcoes)) === "granted") return true;

    return false;

}

// ---------- UI ----------

function syncSetStatus(tipo, textoExtra){

    const el = document.getElementById("syncStatus");

    if(!el) return;

    const mapa = {

        off: [
            "sync-off",
            '<span class="sync-dot"></span> Sincronização desligada'
        ],

        scan: [
            "sync-scan",
            '<span class="sync-dot"></span> Procurando arquivo na pasta...'
        ],

        on: [
            "sync-on",
            '<span class="sync-dot"></span> Conectado — monitorando' +
            (textoExtra ? ` (${textoExtra})` : "")
        ]

    };

    el.className = mapa[tipo][0];
    el.innerHTML = mapa[tipo][1];

    const btnConectar = document.getElementById("btnConectarPasta");
    const btnDesconectar = document.getElementById("btnDesconectarPasta");

    if(btnConectar) btnConectar.style.display = tipo === "off" ? "inline-block" : "none";
    if(btnDesconectar) btnDesconectar.style.display = tipo === "off" ? "none" : "inline-block";

}

function syncAtualizarUltimaChecagem(){

    const el = document.getElementById("syncUltimaChecagem");

    if(!el) return;

    el.style.display = "inline";

    el.textContent =
    "Última checagem: " +
    new Date().toLocaleTimeString("pt-BR");

}

function syncDataHojeIso(){

    const hoje = new Date();

    return hoje.toISOString().slice(0,10);

}

// ---------- Varredura da subpasta ----------

const SYNC_EXT_PEDIDOS = [".xlsx",".xls"];

function syncTemExtensao(nome, lista){

    const n = nome.toLowerCase();

    return lista.some(ext=> n.endsWith(ext));

}

async function syncVarrerPasta(){

    syncSetStatus("scan");

    syncArquivoPedidosHandle = null;

    for await (const [nome, handle] of syncDirHandle.entries()){

        if(handle.kind !== "file") continue;

        if(
            !syncArquivoPedidosHandle &&
            syncTemExtensao(nome, SYNC_EXT_PEDIDOS)
        ){

            syncArquivoPedidosHandle = handle;

            break;

        }

    }

    if(!syncArquivoPedidosHandle){

        alert(
            "Não encontrei nenhum arquivo da Consulta da API (.xlsx/.xls) " +
            "dentro dessa pasta."
        );

        return false;

    }

    return true;

}

// ---------- Processamento automático (reaproveita processarArquivoEData) ----------

async function syncProcessarArquivos(){

    mostrarLoading();

    try{

        const arquivo =
        await syncArquivoPedidosHandle.getFile();

        const dataIso =
        syncDataHojeIso();

        // mantém o campo de data da UI manual coerente também
        const campoData = document.getElementById("dataArquivo");

        if(campoData) campoData.value = dataIso;

        await processarArquivoEData(arquivo, dataIso);

        ocultarLoading();

        // reflete no campo de nome de arquivo da UI manual também
        document.getElementById("nomePedidos").innerText =
        "🔗 " + arquivo.name + " (auto — dia " + dataIso + ")";

        console.log(
            `Sincronização automática concluída — dia ${dataIso}`
        );

    }catch(erro){

        console.error(erro);

        ocultarLoading();

    }

}

// ---------- Loop de monitoramento ----------

function syncPararMonitoramento(){

    if(syncIntervalId){

        clearInterval(syncIntervalId);

        syncIntervalId = null;

    }

}

function syncIniciarMonitoramento(){

    syncPararMonitoramento();

    syncSetStatus("on", syncArquivoPedidosHandle?.name);

    syncIntervalId = setInterval(
        syncChecarMudancas,
        SYNC_INTERVALO_MS
    );

}

async function syncChecarMudancas(){

    try{

        const file =
        await syncArquivoPedidosHandle.getFile();

        syncAtualizarUltimaChecagem();

        if(file.lastModified !== syncLastModifiedPedidos){

            syncLastModifiedPedidos = file.lastModified;

            await syncProcessarArquivos();

        }

    }catch(erro){

        console.error(
            "Erro ao checar mudanças na pasta:",
            erro
        );

    }

}

// ---------- Ações de UI (botões) ----------

async function conectarPastaAcompanhamento(){

    try{

        syncDirHandle = await window.showDirectoryPicker();

        await syncSalvarHandle(syncDirHandle);

        const encontrou = await syncVarrerPasta();

        if(!encontrou){

            syncSetStatus("off");

            return;

        }

        // primeira carga imediata + marca o lastModified atual
        await syncProcessarArquivos();

        const file = await syncArquivoPedidosHandle.getFile();
        syncLastModifiedPedidos = file.lastModified;

        syncIniciarMonitoramento();

    }catch(erro){

        if(erro.name !== "AbortError"){

            console.error(erro);

            alert("Erro ao conectar a pasta: " + erro.message);

        }

    }

}

async function desconectarPastaAcompanhamento(){

    syncPararMonitoramento();

    syncDirHandle = null;
    syncArquivoPedidosHandle = null;
    syncLastModifiedPedidos = 0;

    await syncLimparHandle();

    syncSetStatus("off");

    const elChecagem = document.getElementById("syncUltimaChecagem");

    if(elChecagem) elChecagem.style.display = "none";

}

// ---------- Reconexão automática ao abrir a página ----------

(async function syncTentarReconectar(){

    const handleSalvo = await syncCarregarHandle();

    if(!handleSalvo) return;

    const temPermissao = await syncGarantirPermissao(handleSalvo);

    if(!temPermissao){

        // não força popup de permissão sem interação do usuário;
        // ele clica em "Conectar Pasta" de novo se precisar
        return;

    }

    syncDirHandle = handleSalvo;

    const encontrou = await syncVarrerPasta();

    if(!encontrou) return;

    const file = await syncArquivoPedidosHandle.getFile();
    syncLastModifiedPedidos = file.lastModified;

    await syncProcessarArquivos();

    syncIniciarMonitoramento();

})();
