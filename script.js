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
    });
});

/**
 * Função para limpar o prefixo numérico da coluna ALINHAMENTO (código do modelo).
 * CORREÇÃO: Remove TUDO que estiver antes do primeiro espaço (que deve ser o prefixo X.X).
 * @param {string} modeloOriginal O valor da coluna ALINHAMENTO.
 * @returns {string} O código do modelo limpo e padronizado.
 */
function limparPrefixoModelo(modeloOriginal) {
    if (typeof modeloOriginal !== 'string') return String(modeloOriginal).trim().toUpperCase();
    
    let modelo = modeloOriginal.trim();
    
    // 1. CORREÇÃO: Encontra o primeiro espaço (que separa o prefixo X.X do código real)
    const indexPrimeiroEspaco = modelo.indexOf(' ');
    
    if (indexPrimeiroEspaco !== -1) {
        // Pega a string a partir do primeiro caractere após o espaço.
        modelo = modelo.substring(indexPrimeiroEspaco + 1).trim();
    }
    
    // 2. PADRONIZAÇÃO ADICIONAL
    modelo = modelo.replace(/\s\s+/g, ' '); // Remove múltiplos espaços
    
    return modelo.toUpperCase();
}


/**
 * Função para padronizar o nome da categoria.
 */
function limparCategoria(nomeOriginal) {
    if (typeof nomeOriginal !== 'string') return String(nomeOriginal).trim().toUpperCase();
    
    let nome = nomeOriginal.trim().toUpperCase();
    
    // 1. CORREÇÃO DE SETOR: Mapeamento de VIL
    if (nome === 'VIL') {
        nome = 'VERTICAL ALTO'; 
    }

    // 2. PADRONIZAÇÃO: Remove " DE "
    nome = nome.replace(/ DE /g, ' ').trim(); 
    nome = nome.replace(/\s\s+/g, ' '); 
    
    return nome; 
}

/**
 * Função para exportar os dados no formato XLSX.
 */
function exportarXLSX(dadosParaExportar, nomeArquivo) {
    const ws = XLSX.utils.json_to_sheet(dadosParaExportar);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Lotes Produção");
    
    const dataAtual = new Date().toLocaleDateString('pt-BR').replace(/\//g, '-');
    XLSX.writeFile(wb, `${nomeArquivo}_${dataAtual}.xlsx`);
}

/**
 * Função para exportar os dados no formato PDF.
 */
function exportarPDF(dadosParaExportar, nomeArquivo) {
    if (typeof window.jspdf === 'undefined' || typeof window.jspdf.jsPDF === 'undefined') {
        alert("Erro: Biblioteca PDF não carregada. Verifique os links CDN no index.html.");
        return;
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    
    // Define cabeçalhos dinamicamente (incluindo DIMENSÃO se presente)
    const headersMap = {
        'LINHA': 'Código do Modelo',
        'BOJO': 'Material (BOJO)',
        'ALINHAMENTOS': 'Categoria',
        'DIMENSÃO': 'DIMENSÃO', 
        'QUANTIDADE TOTAL': 'Quantidade Total'
    };
    
    const colunas = Object.keys(dadosParaExportar[0]).filter(key => headersMap[key]);
    const head = [colunas.map(key => headersMap[key])];
    const body = dadosParaExportar.map(item => colunas.map(key => item[key]));

    doc.autoTable({
        head: head,
        body: body,
        startY: 20,
        styles: { fontSize: 8 },
        headStyles: { fillColor: [40, 167, 69] }, 
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

/**
 * Função que orquestra a exportação (chamada pelos botões).
 */
function exportarRelatorio(tipo, formato, filtroCategoria = null) {
    let dadosParaExportar;
    let nomeArquivo;

    // 1. Seleciona os dados
    if (tipo === 'geral') {
        dadosParaExportar = lotesGerais;
        nomeArquivo = 'Resumo_Geral_Lotes';
    } else { // 'detalhe'
        // Aplica o filtro se houver categoria
        if (filtroCategoria) {
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

    // 2. Exporta no formato desejado
    if (formato === 'xlsx') {
        exportarXLSX(dadosParaExportar, nomeArquivo);
    } else if (formato === 'pdf') {
        exportarPDF(dadosParaExportar, nomeArquivo);
    }
}


/**
 * Processa a planilha, faz os dois agrupamentos e prepara a interface.
 */
function processarPlanilha() {
    const fileInput = document.getElementById('excelFileInput');
    const statusDiv = document.getElementById('statusMessage');
    statusDiv.textContent = 'Processando...';
    document.getElementById('processButton').disabled = true;

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
            
            // Leitura por Índice (AOA): Começa a ler na linha 3 (índice 2)
            const range = XLSX.utils.decode_range(worksheet['!ref']);
            range.s.r = 2; 
            const newRange = XLSX.utils.encode_range(range);
            
            const rawDataAOA = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: newRange }); 

            // === ÍNDICES CORRETOS (Base 0) ===
            const INDEX_MODELO = 3;      // Coluna ALINHAMENTO (código: "1.1 VILP-2P/900")
            const INDEX_DIMENSAO = 4;    // Coluna DIMENSÃO (e.g., "3.75", "1.25")
            const INDEX_BOJO = 5;        // Coluna BOJO (PP/INOX)
            const INDEX_CATEGORIA = 6;   // Coluna LINHA (e.g., VERTICAL ALTO, VIL)
            
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
                
                // Limpezas e padronizações
                const modeloLimpo = limparPrefixoModelo(modeloOriginal);
                const categoriaLimpa = limparCategoria(categoriaOriginal);
                const bojoNormalizado = bojoOriginal.toUpperCase();
                const dimensaoNormalizada = dimensaoOriginal.toUpperCase();

                if (!modeloLimpo || !bojoNormalizado || !categoriaLimpa || !dimensaoNormalizada) return;

                linhasProcessadas++;
                todasCategorias.add(categoriaLimpa);

                // CHAVES DE AGRUPAMENTO
                const chaveGeral = `${modeloLimpo}|${bojoNormalizado}|${categoriaLimpa}`;
                const chaveDetalhe = `${modeloLimpo}|${bojoNormalizado}|${categoriaLimpa}|${dimensaoNormalizada}`; 

                // 1. AGRUPAMENTO GERAL (Resumo)
                if (lotesGeraisMap[chaveGeral]) {
                    lotesGeraisMap[chaveGeral]['QUANTIDADE TOTAL']++;
                } else {
                    lotesGeraisMap[chaveGeral] = {
                        'LINHA': modeloLimpo, 
                        'BOJO': bojoNormalizado,
                        'ALINHAMENTOS': categoriaLimpa,
                        'QUANTIDADE TOTAL': 1
                    };
                }
                
                // 2. AGRUPAMENTO DETALHE (Produção)
                if (lotesDetalhesMap[chaveDetalhe]) {
                    lotesDetalhesMap[chaveDetalhe]['QUANTIDADE TOTAL']++;
                } else {
                    lotesDetalhesMap[chaveDetalhe] = {
                        'LINHA': modeloLimpo, 
                        'BOJO': bojoNormalizado,
                        'ALINHAMENTOS': categoriaLimpa,
                        'DIMENSÃO': dimensaoNormalizada, // Novo Campo!
                        'QUANTIDADE TOTAL': 1
                    };
                }
            });

            // Converte e armazena globalmente
            lotesGerais = Object.values(lotesGeraisMap).sort((a, b) => a.LINHA.localeCompare(b.LINHA));
            lotesDetalhes = Object.values(lotesDetalhesMap).sort((a, b) => a.LINHA.localeCompare(b.LINHA));

            if (linhasProcessadas === 0) {
                 statusDiv.textContent = `Nenhuma linha de dados processada.`;
                 document.getElementById('processButton').disabled = false;
                 return;
            }

            statusDiv.textContent = `Processamento concluído! ${lotesDetalhes.length} lotes detalhados gerados.`;

            // HABILITA E GERA BOTÕES
            document.getElementById('downloadSection').style.display = 'block';
            gerarBotoesFiltro(); 

        } catch (error) {
            console.error("Erro geral de processamento:", error);
            statusDiv.textContent = `Erro! Verifique o console do navegador (F12) para detalhes.`;
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
