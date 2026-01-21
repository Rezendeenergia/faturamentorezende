// Elementos DOM
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const loadingArea = document.getElementById('loadingArea');
const formSection = document.getElementById('formSection');
const nfForm = document.getElementById('nfForm');

// Campos do formulário
const dataEmissao = document.getElementById('dataEmissao');
const numeroNF = document.getElementById('numeroNF');
const tipo = document.getElementById('tipo');
const valorBruto = document.getElementById('valorBruto');
const localidade = document.getElementById('localidade');
const tomador = document.getElementById('tomador');
const inss = document.getElementById('inss');
const iss = document.getElementById('iss');
const retencaoEquatorial = document.getElementById('retencaoEquatorial');
const pisCofins = document.getElementById('pisCofins');
const valorNominalCalc = document.getElementById('valorNominalCalc');

// Checkbox PIS/COFINS (pode não existir no HTML antigo)
const presumirPisCofins = document.getElementById('presumirPisCofins');

// Upload via drag & drop
uploadArea.addEventListener('click', () => fileInput.click());

uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');

    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFile(e.target.files[0]);
    }
});

// Processa arquivo
async function handleFile(file) {
    if (!file.type.includes('pdf')) {
        mostrarErro('Por favor, selecione apenas arquivos PDF.');
        return;
    }

    // Mostra loading
    uploadArea.style.display = 'none';
    loadingArea.style.display = 'block';

    const formData = new FormData();
    formData.append('file', file);

    try {
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (result.success) {
            preencherFormulario(result.dados);
            formSection.style.display = 'block';
            loadingArea.style.display = 'none';
        } else {
            mostrarErro(result.error || 'Erro ao processar arquivo.');
            resetarUpload();
        }
    } catch (error) {
        mostrarErro('Erro ao enviar arquivo: ' + error.message);
        resetarUpload();
    }
}

// Preenche formulário com dados extraídos
function preencherFormulario(dados) {
    // Converter data de dd/mm/yyyy para yyyy-mm-dd
    if (dados.data_emissao) {
        const [dia, mes, ano] = dados.data_emissao.split('/');
        dataEmissao.value = `${ano}-${mes}-${dia}`;
    }

    numeroNF.value = dados.numero_nf || '';
    tipo.value = dados.tipo || '';
    valorBruto.value = dados.valor_bruto || '';
    localidade.value = dados.localidade || '';
    tomador.value = dados.tomador || '';

    // Preenche valores calculados (EDITÁVEIS)
    if (dados.retencoes) {
        inss.value = dados.retencoes.inss.toFixed(2);
        iss.value = dados.retencoes.iss.toFixed(2);
        retencaoEquatorial.value = dados.retencoes.retencao_equatorial.toFixed(2);
        pisCofins.value = dados.retencoes.pis_cofins_csll.toFixed(2);

        // Marca checkbox se existir E se PIS/COFINS foi calculado
        if (presumirPisCofins) {
            presumirPisCofins.checked = dados.retencoes.pis_cofins_csll > 0;
        }
    }

    // Recalcula Valor Nominal
    if (typeof recalcularValorNominal === 'function') {
        recalcularValorNominal();
    } else if (dados.valor_nominal_calculado) {
        valorNominalCalc.value = dados.valor_nominal_calculado.toFixed(2);
    }
}

// Recalcula valores quando tipo ou valor bruto mudam
tipo.addEventListener('change', recalcularValores);
valorBruto.addEventListener('input', recalcularValores);

// Se o checkbox existir, adiciona event listener
if (presumirPisCofins) {
    presumirPisCofins.addEventListener('change', recalcularValores);
}

// Event listeners para recalcular Valor Nominal (se o campo for readonly)
if (valorNominalCalc && valorNominalCalc.hasAttribute('readonly')) {
    inss.addEventListener('input', recalcularValorNominal);
    iss.addEventListener('input', recalcularValorNominal);
    retencaoEquatorial.addEventListener('input', recalcularValorNominal);
    pisCofins.addEventListener('input', recalcularValorNominal);
}

async function recalcularValores() {
    const tipoVal = tipo.value;
    const valorBrutoVal = parseFloat(valorBruto.value);

    if (!tipoVal || !valorBrutoVal || valorBrutoVal <= 0) {
        return;
    }

    try {
        // Verifica se checkbox existe, senão usa false como padrão
        const pisCofinsRetido = presumirPisCofins ? presumirPisCofins.checked : false;

        const response = await fetch('/calcular', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                tipo: tipoVal,
                valor_bruto: valorBrutoVal,
                pis_cofins_retido: pisCofinsRetido
            })
        });

        const result = await response.json();

        if (result.success) {
            inss.value = result.retencoes.inss.toFixed(2);
            iss.value = result.retencoes.iss.toFixed(2);
            retencaoEquatorial.value = result.retencoes.retencao_equatorial.toFixed(2);
            pisCofins.value = result.retencoes.pis_cofins_csll.toFixed(2);

            // Recalcula Valor Nominal (se campo for readonly) ou usa o retornado
            if (valorNominalCalc.hasAttribute('readonly') && typeof recalcularValorNominal === 'function') {
                recalcularValorNominal();
            } else {
                valorNominalCalc.value = result.valor_nominal.toFixed(2);
            }
        }
    } catch (error) {
        console.error('Erro ao recalcular:', error);
    }
}

// Função para recalcular Valor Nominal automaticamente
function recalcularValorNominal() {
    const valorBrutoVal = parseFloat(valorBruto.value) || 0;
    const inssVal = parseFloat(inss.value) || 0;
    const issVal = parseFloat(iss.value) || 0;
    const retencaoEquatorialVal = parseFloat(retencaoEquatorial.value) || 0;
    const pisCofinsVal = parseFloat(pisCofins.value) || 0;

    // Calcula: Valor Bruto - INSS - ISS - Retenção Equatorial - PIS/COFINS
    const valorNominal = valorBrutoVal - inssVal - issVal - retencaoEquatorialVal - pisCofinsVal;

    valorNominalCalc.value = valorNominal.toFixed(2);
}

// Submeter formulário
nfForm.addEventListener('submit', async (e) => {
    e.preventDefault();

    // Converte data de yyyy-mm-dd para dd/mm/yyyy
    const [ano, mes, dia] = dataEmissao.value.split('-');
    const dataFormatada = `${dia}/${mes}/${ano}`;

    const dados = {
        data_emissao: dataFormatada,
        numero_nf: numeroNF.value,
        tipo: tipo.value,
        valor_bruto: parseFloat(valorBruto.value),
        localidade: localidade.value,
        tomador: tomador.value,
        inss: parseFloat(inss.value) || 0,
        iss: parseFloat(iss.value) || 0,
        retencao_equatorial: parseFloat(retencaoEquatorial.value) || 0,
        pis_cofins_csll: parseFloat(pisCofins.value) || 0,
        pis_cofins_retido: presumirPisCofins ? presumirPisCofins.checked : (parseFloat(pisCofins.value) > 0),
        valor_nominal_calculado: parseFloat(valorNominalCalc.value) || 0,
        valor_nominal_conferencia: parseFloat(valorNominalCalc.value) || 0
    };

    try {
        const response = await fetch('/salvar', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(dados)
        });

        const result = await response.json();

        if (result.success) {
            mostrarSucesso(result.message);
            setTimeout(() => {
                limparFormulario();
                resetarUpload();
            }, 2000);
        } else {
            mostrarErro(result.error || 'Erro ao salvar nota fiscal.');
        }
    } catch (error) {
        mostrarErro('Erro ao salvar: ' + error.message);
    }
});

// Funções auxiliares
function limparFormulario() {
    nfForm.reset();
    inss.value = '';
    iss.value = '';
    retencaoEquatorial.value = '';
    pisCofins.value = '';
    valorNominalCalc.value = '';

    // Marca checkbox novamente se existir
    if (presumirPisCofins) {
        presumirPisCofins.checked = true;
    }

    formSection.style.display = 'none';
}

function resetarUpload() {
    uploadArea.style.display = 'block';
    loadingArea.style.display = 'none';
    fileInput.value = '';
}

function mostrarSucesso(mensagem) {
    const modal = document.getElementById('successModal');
    const messageEl = document.getElementById('successMessage');
    messageEl.textContent = mensagem;
    modal.classList.add('show');
}

function mostrarErro(mensagem) {
    const modal = document.getElementById('errorModal');
    const messageEl = document.getElementById('errorMessage');
    messageEl.textContent = mensagem;
    modal.classList.add('show');
}

function fecharModal() {
    document.getElementById('successModal').classList.remove('show');
    document.getElementById('errorModal').classList.remove('show');
}

// Fechar modal ao clicar fora
window.addEventListener('click', (e) => {
    if (e.target.classList.contains('modal')) {
        fecharModal();
    }
});