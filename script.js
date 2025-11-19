// ====================================================================================
// PARTE CRÍTICA: DEFINIÇÃO DE REGRAS DE AGRUPAMENTO
// ====================================================================================
const MAPA_MODELO_GRUPO = {
    // === BASEADO NOS 10 GRUPOS FORNECIDOS ANTERIORMENTE ===

    // Grupo A (Valp-850 e Variações)
    "VABP-PVT/850": "A", "VABP/850 MA": "A", "VACP-PVT/850": "A", "VACP-PVT/850 MA": "A",
    "VACP/850": "A", "VACP/750": "A", "VACP/850 MA": "A", "VAHP-PVT/850 MA": "A", "VAHP-PVT/850": "A", "VAHP/850": "A",
    "VALP- PVT/850": "A", "VALP-PVT/850": "A", "VALP-PVT/850 MA": "A", "VALP/850": "A", "VALP/750": "A", "VALP/750 MA": "A", "VABP/850": "A", "VABP/750": "A",
    "VALP/850 MA": "A", 

    // Grupo B (VAH/850)
    "VAH/850": "B", "VAL/850": "B",

    // Grupo C (VAP/850 e Variações)
    "VAP/850": "C", "VAP/850 MA": "C",

    // Grupo D (VCA2P/1040)
    "VCA2P/1040": "D",

    // Grupo E (VCAG Complexo)
    "VCAG (1,25) + VCA2P (2,50)/1040": "E", "VCAG/1040": "E", "VCAGR (1,25) + VCAG1P (1,25)/1040": "E",

    // Grupo F (VIL-2P/900)
    "VIL-2P/900": "F", "VIL-2P CENTRAL/1760": "F", "VIL-2P FRONTAL/900": "F",

    // Grupo G (VIL-2P/900 CANTO": "G",
    "VIL-2P/900 CANTO": "G",

    // Grupo H (VILP-2P/900 e Variações)
    "VILP-2P/900": "H", "VILP-2P/900 MA": "H",

    // Grupo I (VR-900 e Variações)
    "VR1P/900": "I", "VR2P/900": "I", "VR2P/900 MA": "I", "VR2PA/900": "I", "VRQU/900": "I", "VRQR/900": "I", "VRHB/900": "I",


    // Grupo J ()
    "VRA1P/1040": "J", "VRA2P/1040": "J", "VRAG (1,25) + VRA2P (2,50)/1040": "J",
    "VRAG/1040": "J", "VRAGR/1040": "J", "VRAG(3,75) + VRAG1P (1,25)/1040": "J", "VRAG2N/900": "J", "VRAG (2,50) +VRA1P (1,25)/1040": "J",

    // Grupo K ()
    "VIL-2P PONTA CT 180/900": "K",

    // Grupo L ()
    "VIL-3P/900": "L",

    // Grupo M ()
    "ICDT": "M", "ICFT": "M",

    // Grupo N ()
    "VC2P/900": "N", "VCQU/900": "N",
    
    // Grupo O ()
    "IRAS": "O",
    
};

// === MAPA DE CORES PARA AGRUPAMENTO ===
const GRUPO_COLORS = {
    "A": "E0F7FA", // Azul Claro Suave (Light Cyan)
    "B": "F1F8E9", // Verde Menta Suave (Pale Green)
    "C": "FFF3E0", // Âmbar Claro Suave (Very Pale Orange/Amber)
    "D": "FBE4E4", // Rosa Claro Suave (Very Light Pink)
    "E": "E8EAF6", // Índigo Claro Suave (Very Light Indigo)
    "F": "F3E5F5", // Roxo Claro Suave (Very Light Purple/Lavender)
    "G": "E1F5FE", // Azul Bebê (Very Light Blue)
    "H": "FFE0B2", // Laranja Claro (Light Orange)
    "I": "DCF8C6", // Verde Limão (Light Lime Green)
    "J": "FCE4EC", // Rosa Pastel (Pale Pink)
    "K": "FFFDE7", // **Amarelo Gema Muito Claro** (Very Light Yellow/Cream)
    "L": "F0F4C3", // **Verde Oliva Muito Claro** (Very Light Olive/Lime)
    "M": "E1BEE7", // **Malva/Lilás Pálido** (Pale Mauve/Light Purple)
    "N": "B2EBF2", // **Ciano/Turquesa Claro** (Light Cyan/Turquoise)
    "O": "FFCCBC", // **Pêssego/Coral Claro** (Light Peach/Coral)

    // Para grupos não mapeados (se o grupo for o próprio nome do modelo limpo)
    "N/A": "DDDDDD" 
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
 * Adiciona o cabeçalho personalizado (apenas o Título) ao topo da planilha.
 */
function adicionarCabecalhoPersonalizado(ws, nomeRelatorio, dadosIniciais) {
    const colunasChave = Object.keys(dadosIniciais[0]);
    const numColunas = colunasChave.length; 
    const linhaInicial = 0; // Começa na linha 1 do Excel (índice 0)
    
    // Título Principal (Linha 1)
    const titulo = `RELATÓRIO OTIMIZADO DE LOTES - ${nomeRelatorio} | ${new Date().toLocaleDateString('pt-BR')}`;
    XLSX.utils.sheet_add_aoa(ws, [[titulo]], { origin: -1 });
    
    if (!ws['!merges']) ws['!merges'] = [];
    ws['!merges'].push({ s: { r: linhaInicial, c: 0 }, e: { r: linhaInicial, c: numColunas - 1 } }); 

    const tituloCell = XLSX.utils.encode_cell({ r: linhaInicial, c: 0 });
    if (!ws[tituloCell]) ws[tituloCell] = { v: titulo, t: 's' };

    // Layout: Fundo Azul Marinho (003366) e Fonte Maior (16)
    ws[tituloCell].s = {
        font: { name: "Arial", sz: 16, bold: true, color: { rgb: "FFFFFF" } }, 
        alignment: { horizontal: "center", vertical: "center" },
        fill: { fgColor: { rgb: "003366" } },
        border: { bottom: { style: "medium", color: { rgb: "000000" } } } 
    };
    
    // Retornamos 1 para indicar que apenas 1 linha (o título) foi adicionada
    return 1; 
}


/**
 * Aplica formatação e estilos no XLSX. 
 */
function aplicarFormatoBasico(dados, ws, nomeRelatorio, startRow) {
    if (dados.length === 0) return;

    const colunasChave = Object.keys(dados[0]);
    const numColunas = colunasChave.length;

    // Range da tabela de dados
    const range = { 
        s: { r: startRow, c: 0 }, 
        e: { r: startRow + dados.length, c: numColunas - 1 } 
    };
    ws['!ref'] = XLSX.utils.encode_range(range);
    
    // 1. Largura das Colunas
    ws['!cols'] = colunasChave.map(colName => {
        let wch = 25; // Padrão
        if (colName === 'GRUPO' || colName === 'BOJO') wch = 12; 
        if (colName === 'DIMENSÃO') wch = 15;
        if (colName === 'QUANTIDADE TOTAL') wch = 18;
        if (colName === 'LINHA') wch = 22; 
        return { wch: wch };
    });
    
    // 2. Filtro Automático
    ws['!autofilter'] = { ref: XLSX.utils.encode_range(range) };
    
    // 3. Estilo do Cabeçalho da Tabela (FUNDO PRETO)
    const headerStyle = {
        fill: { fgColor: { rgb: "000000" } }, // Fundo Preto
        font: { bold: true, color: { rgb: "FFFFFF" }, name: "Arial", sz: 10 }, 
        alignment: { horizontal: "center", vertical: "center" },
        border: { 
            top: { style: "medium", color: { rgb: "000000" } }, bottom: { style: "medium", color: { rgb: "FFFFFF" } }, 
            left: { style: "thin", color: { rgb: "FFFFFF" } }, right: { style: "thin", color: { rgb: "FFFFFF" } } 
        }
    };

    const centerStyle = { 
        alignment: { horizontal: "center", vertical: "center" }, 
        font: { name: "Arial", sz: 10 }
    };
    
    // Aplica estilo ao Cabeçalho da Tabela e ao Corpo
    colunasChave.forEach((colName, index) => {
        // Estilo do Cabeçalho da Tabela
        const headerAddress = XLSX.utils.encode_cell({ r: range.s.r, c: range.s.c + index });
        const headerCell = ws[headerAddress];
        
        if (headerCell) {
            headerCell.s = headerStyle;
            if (typeof headerCell.v !== 'undefined') {
                headerCell.t = 's';
            }
        }

        // Estilo do Corpo da Tabela
        for (let r = range.s.r + 1; r <= range.e.r; r++) {
            const dataCellAddress = XLSX.utils.encode_cell({ r: r, c: range.s.c + index });
            const dataCell = ws[dataCellAddress];
            
            // === CORES DOS GRUPOS ===
            const grupoAddress = XLSX.utils.encode_cell({ r: r, c: 0 });
            const grupo = (ws[grupoAddress] && ws[grupoAddress].v) || 'N/A';
            const corGrupo = GRUPO_COLORS[grupo] || GRUPO_COLORS['N/A'];
            // ========================

            if (dataCell) {
                const isNumeric = colName === 'QUANTIDADE TOTAL';
                
                const cellBorder = { 
                    border: { 
                        top: { style: "thin", color: { rgb: "DDDDDD" } }, bottom: { style: "thin", color: { rgb: "DDDDDD" } }, 
                        left: { style: "thin", color: { rgb: "DDDDDD" } }, right: { style: "thin", color: { rgb: "DDDDDD" } }
                    }
                };
                
                // Aplica a cor do GRUPO como fundo da linha
                const fillStyle = { fill: { fgColor: { rgb: corGrupo } } };
                
                // Cria o objeto de estilo completo para esta célula
                dataCell.s = { 
                    ...centerStyle,
                    ...cellBorder,
                    ...fillStyle 
                };

                // Garante que a célula tem o tipo correto
                if (isNumeric) {
                    dataCell.t = 'n'; // Number
                } else if (typeof dataCell.v !== 'undefined') {
                    dataCell.t = 's'; // String
                }
            }
        }
    });
}

/**
 * Funçao exportarXLSX (CORRIGIDA PARA COMPACTAÇÃO)
 */
function exportarXLSX(dadosParaExportar, nomeArquivo, isFiltered) {
    const dadosFinais = selecionarColunas(dadosParaExportar, isFiltered);
    
    const ws = {}; 
    
    const nomeRelatorioLimpo = nomeArquivo.replace(/_/g, ' '); 
    
    // 1. Adiciona o TÍTULO na linha 1 (índice 0)
    const totalTitleRows = adicionarCabecalhoPersonalizado(ws, nomeRelatorioLimpo, dadosFinais); 
    
    // 2. Os dados (incluindo o cabeçalho da tabela) começam imediatamente após o título (Linha 2, índice 1)
    // O valor de totalTitleRows agora é 1, garantindo que o START_ROW_DATA seja 1 (Linha 2)
    const START_ROW_DATA = totalTitleRows; // Deve ser 1
    XLSX.utils.sheet_add_json(ws, dadosFinais, { origin: START_ROW_DATA, skipHeader: false });
    
    // 3. Aplica estilos
    aplicarFormatoBasico(dadosFinais, ws, nomeRelatorioLimpo, START_ROW_DATA);
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Lotes Produção");
    
    const dataAtual = new Date().toLocaleDateString('pt-BR').replace(/\//g, '-');
    XLSX.writeFile(wb, `${nomeArquivo}_${dataAtual}.xlsx`, { cellStyles: true }); 
}

/**
 * Funçao exportarPDF (Design mantido)
 */
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
        theme: 'grid', // Tema grid para bordas nítidas
        styles: { 
            fontSize: 9, 
            font: 'helvetica', 
            textColor: [52, 58, 64] // Texto cinza escuro
        },
        headStyles: { 
            fillColor: [0, 123, 255], // Fundo Azul NSF (Primário)
            textColor: 255, // Texto Branco
            fontStyle: 'bold',
            fontSize: 10
        }, 
        alternateRowStyles: {
            fillColor: [248, 249, 250] // Linhas alternadas em cinza muito claro
        },
        bodyStyles: {
            lineColor: [222, 226, 230], // Cor da linha (cinza claro)
            lineWidth: 0.1 
        },
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
            
            // Lendo a partir da LINHA 6 (índice 5)
            const range = XLSX.utils.decode_range(worksheet['!ref']);
            range.s.r = 5; // Linha de início dos dados (Linha 6 da planilha)
            const newRange = XLSX.utils.encode_range(range);
            
            const rawDataAOA = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: newRange }); 

            // === NOVOS ÍNDICES DE COLUNAS (Base 0) ===
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
                
                // Adiciona a letra do GRUPO ou o próprio nome do Modelo
                const grupoLetra = MAPA_MODELO_GRUPO[modeloLimpo] || modeloLimpo;
                
                const categoriaLimpa = limparCategoria(categoriaOriginal); 
                
                // Garante que BOJO e DIMENSÃO tenham um valor padrão se vazios ('N/A')
                const bojoNormalizado = (bojoOriginal || 'N/A').toUpperCase();
                const dimensaoNormalizada = (dimensaoOriginal || 'N/A').toUpperCase();


                // Apenas verifica se o modelo e a categoria estão presentes.
                if (!modeloLimpo || !categoriaLimpa) return; 

                linhasProcessadas++;
                todasCategorias.add(categoriaLimpa);

                // CHAVES DE AGRUPAMENTO
                const chaveGeral = `${modeloLimpo}|${bojoNormalizado}|${categoriaLimpa}|${grupoLetra}`; 
                const chaveDetalhe = `${modeloLimpo}|${bojoNormalizado}|${categoriaLimpa}|${dimensaoNormalizada}|${grupoLetra}`; 

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

            statusDiv.textContent = `Processamento concluído! Total de itens contados: ${linhasProcessadas}.`;

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