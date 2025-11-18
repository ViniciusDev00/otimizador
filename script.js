// ====================================================================================
// PARTE CRÍTICA: DEFINIÇÃO DE REGRAS DE AGRUPAMENTO
// Este mapa associa o NOME LIMPO do Modelo a uma LETRA de GRUPO (A, B, C, etc.).
// ====================================================================================
const MAPA_MODELO_GRUPO = {
    // === BASEADO NOS 10 GRUPOS FORNECIDOS ANTERIORMENTE ===

    // Grupo A (Valp-850 e Variações)
    "VABP-PVT/850": "A", "VABP/850 MA": "A", "VACP-PVT/850": "A", "VACP-PVT/850 MA": "A",
    "VACP/850": "A", "VACP/850 MA": "A", "VAHP-PVT/850 MA": "A", "VAHP/850": "A",
    "VALP- PVT/850": "A", "VALP-PVT/850": "A", "VALP-PVT/850 MA": "A", "VALP/850": "A",

    // Grupo B (VAH/850)
    "VAH/850": "B",

    // Grupo C (VAP/850 e Variações)
    "VAP/850": "C", "VAP/850 MA": "C",

    // Grupo D (VCA2P/1040)
    "VCA2P/1040": "D",

    // Grupo E (VCAG Complexo)
    "VCAG (1,25) + VCA2P (2,50)/1040": "E",

    // Grupo F (VIL-2P/900)
    "VIL-2P/900": "F",

    // Grupo G (VIL-2P/900 CANTO)
    "VIL-2P/900 CANTO": "G",

    // Grupo H (VILP-2P/900 e Variações)
    "VILP-2P/900": "H", "VILP-2P/900 MA": "H",

    // Grupo I (VR-900 e Variações)
    "VR1P/900": "I", "VR2P/900": "I", "VR2P/900 MA": "I", "VR2PA/900": "I",

    // Grupo J (VR-1040 e Variações)
    "VRA1P/1040": "J", "VRA2P/1040": "J", "VRAG (1,25) + VRA2P (2,50)/1040": "J",
    "VRAG/1040": "J", "VRAGR/1040": "J",
};
// ====================================================================================

// Variáveis globais para armazenar os resultados
let lotesGerais = [];     
let lotesDetalhes = [];   
let todasCategorias = new Set(); 

// Inicialização do listener de habilitação do botão
document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('excelFileInput');
    const processButton = document.getElementById('processButton');
    
    fileInput.addEventListener('change', () => {
        processButton.disabled = fileInput.files.length === 0;
        document.getElementById('processButtonText').textContent = 'Processar Planilha';
    });
});

/**
 * Mapeia o nome do modelo para sua versão padrão, se uma regra existir.
 */
function mapearModeloEquivalente(modeloLimpo) {
    // Retorna a letra do GRUPO ou o próprio modelo Limpo (se não mapeado)
    return MAPA_MODELO_GRUPO[modeloLimpo] || modeloLimpo;
}

/**
 * Funções de Limpeza de Dados
 */
function limparPrefixoModelo(modeloOriginal) {
    if (typeof modeloOriginal !== 'string') return String(modeloOriginal).trim().toUpperCase();
    let modelo = modeloOriginal.trim();
    const indexPrimeiroEspaco = modelo.indexOf(' ');
    if (indexPrimeiroEspaco !== -1) {
        modelo = modelo.substring(indexPrimeiroEspaco + 1).trim();
    }
    modelo = modelo.replace(/\s\s+/g, ' ');
    return modelo.toUpperCase();
}

/**
 * Função de limpeza de categoria (Permite VIL aparecer).
 */
function limparCategoria(nomeOriginal) {
    if (typeof nomeOriginal !== 'string') return String(nomeOriginal).trim().toUpperCase();
    let nome = nomeOriginal.trim().toUpperCase();
    
    nome = nome.replace(/ DE /g, ' ').trim(); 
    nome = nome.replace(/\s\s+/g, ' '); 
    return nome; 
}

/**
 * Funções de Interface e Exportação
 */
function gerarBotoesFiltro() {
    const containerXLSX = document.getElementById('filterColXLSX');
    const containerPDF = document.getElementById('filterColPDF');
    
    containerXLSX.innerHTML = '<h4>Download XLSX</h4>';
    containerPDF.innerHTML = '<h4>Download PDF</h4>';

    Array.from(todasCategorias).sort().forEach(categoria => {
        if (!categoria || categoria === 'N/A') return; 
        
        containerXLSX.innerHTML += `<button onclick="exportarRelatorio('detalhe', 'xlsx', '${categoria}')" class="btn">${categoria}</button>`;
        containerPDF.innerHTML += `<button onclick="exportarRelatorio('detalhe', 'pdf', '${categoria}')" class="btn">${categoria}</button>`;
    });
}

function selecionarColunas(data, isFiltered) {
    const colunasPadrao = ['GRUPO', 'LINHA', 'BOJO', 'ALINHAMENTOS', 'DIMENSÃO', 'QUANTIDADE TOTAL'];
    const colunasGerais = ['GRUPO', 'LINHA', 'BOJO', 'ALINHAMENTOS', 'QUANTIDADE TOTAL'];

    let colunasFinais = data[0] && data[0].DIMENSÃO ? colunasPadrao : colunasGerais;

    if (isFiltered) {
        colunasFinais = colunasFinais.filter(col => col !== 'ALINHAMENTOS');
    }

    return data.map(item => {
        const novoItem = {};
        colunasFinais.forEach(col => {
            novoItem[col] = item[col];
        });
        return novoItem;
    });
}

/**
 * Aplica formatação e estilos no XLSX.
 */
function aplicarFormatoBasico(dados, ws) {
    if (dados.length === 0) return;

    const colunasChave = Object.keys(dados[0]);
    const range = XLSX.utils.decode_range(ws['!ref']);
    
    ws['!cols'] = colunasChave.map(colName => {
        let wch = 20; 
        if (colName === 'DIMENSÃO' || colName === 'BOJO' || colName === 'GRUPO') wch = 12; 
        if (colName === 'QUANTIDADE TOTAL') wch = 15;
        if (colName === 'ALINHAMENTOS') wch = 25; 
        return { wch: wch };
    });
    
    ws['!autofilter'] = { ref: XLSX.utils.encode_range(range) };
    
    const headerStyle = {
        fill: { fgColor: { rgb: "007bff" } }, 
        font: { bold: true, color: { rgb: "FFFFFF" } }, 
        alignment: { horizontal: "center", vertical: "center" },
        border: { 
            top: { style: "medium" }, bottom: { style: "medium" }, 
            left: { style: "thin" }, right: { style: "thin" }
        }
    };

    const centerStyle = { alignment: { horizontal: "center", vertical: "center" } };
    
    colunasChave.forEach((colName, index) => {
        const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: range.s.c + index });
        
        if (ws[cellAddress]) {
            ws[cellAddress].s = headerStyle;
        }

        const isNumericOrCode = colName !== 'ALINHAMENTOS'; 

        for (let r = range.s.r + 1; r <= range.e.r; r++) {
            const dataCellAddress = XLSX.utils.encode_cell({ r: r, c: range.s.c + index });
            
            const cellBorder = { 
                border: { 
                    top: { style: "thin" }, bottom: { style: "thin" }, 
                    left: { style: "thin" }, right: { style: "thin" }
                }
            };
            
            if (ws[dataCellAddress]) {
                ws[dataCellAddress].s = { 
                    ...ws[dataCellAddress].s, 
                    ...(isNumericOrCode ? centerStyle : {}),
                    ...cellBorder 
                };
            }
        }
    });
}

function exportarXLSX(dadosParaExportar, nomeArquivo, isFiltered) {
    const dadosFinais = selecionarColunas(dadosParaExportar, isFiltered);
    
    const ws = XLSX.utils.json_to_sheet(dadosFinais);
    
    aplicarFormatoBasico(dadosFinais, ws);
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Lotes Produção");
    
    const dataAtual = new Date().toLocaleDateString('pt-BR').replace(/\//g, '-');
    XLSX.writeFile(wb, `${nomeArquivo}_${dataAtual}.xlsx`, { cellStyles: true }); 
}

function exportarPDF(dadosParaExportar, nomeArquivo, isFiltered) {
    if (typeof window.jspdf === 'undefined' || typeof window.jspdf.jsPDF === 'undefined') {
        alert("Erro: Biblioteca PDF não carregada. Verifique os links CDN no index.html.");
        return;
    }
    
    const dadosFinais = selecionarColunas(dadosParaExportar, isFiltered);

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    
    const headersMap = {
        'GRUPO': 'GRUPO', 
        'LINHA': 'Código do Modelo',
        'BOJO': 'Material (BOJO)',
        'ALINHAMENTOS': 'Categoria',
        'DIMENSÃO': 'DIMENSÃO', 
        'QUANTIDADE TOTAL': 'Quantidade Total'
    };
    
    const colunas = Object.keys(dadosFinais[0]).filter(key => headersMap[key]);
    const head = [colunas.map(key => headersMap[key])];
    const body = dadosFinais.map(item => colunas.map(key => item[key]));

    doc.autoTable({
        head: head,
        body: body,
        startY: 20,
        styles: { fontSize: 8 },
        headStyles: { fillColor: [0, 123, 255] }, 
        didDrawPage: function (data) {
            doc.setFontSize(14);
            doc.text("Lista Otimizada de Lotes - PCP", data.settings.margin.left, 10);
            doc.setFontSize(10);
            doc.text(`Relatório: ${nomeArquivo.replace(/_/g, ' ')} | Data: ${new Date().toLocaleDateString('pt-BR')}`, data.settings.margin.left, 15);
        }
    });

    const dataAtual = new Date().toLocaleDateString('pt-BR').replace(/\//g, '-');
    doc.save(`${nomeArquivo}_${dataAtual}.pdf`);
}

function exportarRelatorio(tipo, formato, filtroCategoria = null) {
    let dadosParaExportar;
    let nomeArquivo;
    let isFiltered = filtroCategoria !== null; 

    if (tipo === 'geral') {
        dadosParaExportar = lotesGerais;
        nomeArquivo = 'Resumo_Geral_Lotes';
    } else { 
        if (isFiltered) {
            dadosParaExportar = lotesDetalhes.filter(lote => lote.ALINHAMENTOS === filtroCategoria);
            nomeArquivo = `Detalhe_Lotes_${filtroCategoria.replace(/\s/g, '_')}`;
        } else {
            dadosParaExportar = lotesDetalhes;
            nomeArquivo = 'Detalhe_Lotes_Completo';
        }
    }

    if (dadosParaExportar.length === 0) {
        alert("Nenhum dado encontrado para o filtro selecionado.");
        return;
    }

    if (formato === 'xlsx') {
        exportarXLSX(dadosParaExportar, nomeArquivo, isFiltered); 
    } else if (formato === 'pdf') {
        exportarPDF(dadosParaExportar, nomeArquivo, isFiltered); 
    }
}


/**
 * Função principal de processamento.
 */
function processarPlanilha() {
    const fileInput = document.getElementById('excelFileInput');
    const statusDiv = document.getElementById('statusMessage');
    
    statusDiv.textContent = 'Processando...';
    document.getElementById('processButton').disabled = true;
    document.getElementById('downloadSection').style.display = 'none'; 

    const file = fileInput.files[0];
    if (!file) {
        statusDiv.textContent = 'Erro: Nenhum arquivo selecionado.';
        document.getElementById('processButton').disabled = false;
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            // Lendo a partir da LINHA 7 (índice 6) para o novo formato
            const range = XLSX.utils.decode_range(worksheet['!ref']);
            range.s.r = 6; // Linha de início dos dados
            const newRange = XLSX.utils.encode_range(range);
            
            const rawDataAOA = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: newRange }); 

            // === NOVOS ÍNDICES DE COLUNAS (Base 0 - F=5, G=6, H=7, I=8) ===
            const INDEX_MODELO = 5;      // Coluna F (ALINHAMENTO)
            const INDEX_DIMENSAO = 6;    // Coluna G (DIMENSÃO)
            const INDEX_BOJO = 7;        // Coluna H (BOJO)
            const INDEX_CATEGORIA = 8;   // Coluna I (LINHA/Setor)
            
            const lotesGeraisMap = {};   
            const lotesDetalhesMap = {}; 
            todasCategorias.clear();
            let linhasProcessadas = 0;

            rawDataAOA.forEach(row => {
                if (row.length < INDEX_CATEGORIA + 1) return;

                const modeloOriginal = String(row[INDEX_MODELO] || '').trim();
                const dimensaoOriginal = String(row[INDEX_DIMENSAO] || '').trim();
                const bojoOriginal = String(row[INDEX_BOJO] || '').trim();
                const categoriaOriginal = String(row[INDEX_CATEGORIA] || '').trim();
                
                // 1. Limpeza e Mapeamento
                const modeloLimpo = limparPrefixoModelo(modeloOriginal);
                
                // Atribuição de GRUPO
                const grupoLetra = mapearModeloEquivalente(modeloLimpo);
                
                const categoriaLimpa = limparCategoria(categoriaOriginal); 
                const bojoNormalizado = bojoOriginal.toUpperCase();
                const dimensaoNormalizada = dimensaoOriginal.toUpperCase();

                if (!modeloLimpo || !bojoNormalizado || !categoriaLimpa || !dimensaoNormalizada) return;

                linhasProcessadas++;
                todasCategorias.add(categoriaLimpa);

                // CHAVES DE AGRUPAMENTO (AGORA SEM O GRUPO NA CHAVE DE CONTAGEM)
                // A chave é apenas a especificação do produto, o que permite a agregação 81->80.
                const chaveGeral = `${modeloLimpo}|${bojoNormalizado}|${categoriaLimpa}`;
                const chaveDetalhe = `${modeloLimpo}|${bojoNormalizado}|${categoriaLimpa}|${dimensaoNormalizada}`; 

                // 1. AGRUPAMENTO GERAL (Resumo)
                if (lotesGeraisMap[chaveGeral]) {
                    lotesGeraisMap[chaveGeral]['QUANTIDADE TOTAL']++;
                } else {
                    lotesGeraisMap[chaveGeral] = {
                        'GRUPO': grupoLetra, 
                        'LINHA': modeloLimpo, 
                        'BOJO': bojoNormalizado,
                        'ALINHAMENTOS': categoriaLimpa,
                        'QUANTIDADE TOTAL': 1
                    };
                }
                
                // 2. AGRUPAMENTO DETALHE (Produção - Com Dimensão)
                if (lotesDetalhesMap[chaveDetalhe]) {
                    lotesDetalhesMap[chaveDetalhe]['QUANTIDADE TOTAL']++;
                } else {
                    lotesDetalhesMap[chaveDetalhe] = {
                        'GRUPO': grupoLetra, 
                        'LINHA': modeloLimpo, 
                        'BOJO': bojoNormalizado,
                        'ALINHAMENTOS': categoriaLimpa,
                        'DIMENSÃO': dimensaoNormalizada, 
                        'QUANTIDADE TOTAL': 1
                    };
                }
            });

            // 3. Finalização e ordenação
            lotesGerais = Object.values(lotesGeraisMap).sort((a, b) => a.LINHA.localeCompare(b.LINHA));
            lotesDetalhes = Object.values(lotesDetalhesMap).sort((a, b) => a.LINHA.localeCompare(b.LINHA));

            if (linhasProcessadas === 0) {
                 statusDiv.textContent = `Nenhuma linha de dados válida processada.`;
                 document.getElementById('processButton').disabled = false;
                 return;
            }

            // A contagem é o número de itens lidos, não o número de linhas do relatório.
            statusDiv.textContent = `Processamento concluído! Total de itens lidos: ${linhasProcessadas}.`;

            // 4. Habilita a interface de download
            document.getElementById('downloadSection').style.display = 'block';
            gerarBotoesFiltro(); 

        } catch (error) {
            console.error("Erro fatal durante o processamento:", error);
            statusDiv.textContent = `Erro fatal! Consulte o console (F12) para o desenvolvedor.`;
        } finally {
            document.getElementById('processButton').disabled = false;
        }
    };

    reader.onerror = function(ex) {
        statusDiv.textContent = 'Erro ao ler o arquivo.';
        console.error(ex);
        document.getElementById('processButton').disabled = false;
    };

    reader.readAsArrayBuffer(file);
}

// Inicialização do listener de habilitação do botão
document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('excelFileInput');
    const processButton = document.getElementById('processButton');
    
    fileInput.addEventListener('change', () => {
        processButton.disabled = fileInput.files.length === 0;
    });
});
