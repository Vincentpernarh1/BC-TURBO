// assets/script.js

// Global variable to store SAP lookup timeout
let sapLookupTimeout = null;

// TAB SWITCHING LOGIC
function switchTab(tabName) {
    // Remove active class from buttons
    document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
    // Hide all modules
    document.querySelectorAll('.module').forEach(el => el.classList.remove('active'));

    // Get nav-bar element
    const navBar = document.getElementById('nav-bar');

    // Activate specific selection
    if(tabName === 'qme') {
        document.querySelector('#mod-qme').classList.add('active');
        document.querySelectorAll('.nav-item')[0].classList.add('active');
        navBar.classList.add('hide'); // Hide nav-bar
    } else if(tabName === 'freq') {
        document.querySelector('#mod-freq').classList.add('active');
        document.querySelectorAll('.nav-item')[1].classList.add('active');
        navBar.classList.add('hide'); // Hide nav-bar
    } else if(tabName === 'emb') {
        document.querySelector('#mod-emb').classList.add('active');
        document.querySelectorAll('.nav-item')[2].classList.add('active');
        navBar.classList.add('hide'); // Hide nav-bar
    } else if(tabName === 'veh') {
        document.querySelector('#mod-veh').classList.add('active');
        document.querySelectorAll('.nav-item')[3].classList.add('active');
        navBar.classList.remove('hide'); // Show nav-bar on dashboard
    }
    else if(tabName === 'fluxo') {
        document.querySelector('#mod-fluxo').classList.add('active');
        document.querySelectorAll('.nav-item')[4].classList.add('active');
        navBar.classList.remove('hide'); // Show nav-bar on dashboard
    }
    else if(tabName === 'trip') {
        document.querySelector('#mod-trip').classList.add('active');
        document.querySelectorAll('.nav-item')[5].classList.add('active');
        navBar.classList.remove('hide'); // Show nav-bar on dashboard
    }
    else if(tabName === 'waiting') {
        document.querySelector('#mod-waiting').classList.add('active');
        document.querySelectorAll('.nav-item')[6].classList.add('active');
        navBar.classList.remove('hide'); // Show nav-bar on dashboard
    }
    else if(tabName === 'stact') {
        document.querySelector('#mod-stact').classList.add('active');
        document.querySelectorAll('.nav-item')[7].classList.add('active');
        navBar.classList.remove('hide'); // Show nav-bar on dashboard
    }
    else if(tabName === 'dash') {
        document.querySelector('#mod-dash').classList.add('active');
        document.querySelectorAll('.nav-item')[7].classList.add('active');
        navBar.classList.remove('hide'); // Show nav-bar on dashboard
    }
}

// PYTHON INTERACTION (Calls the backend API)
function selectDB() {
    // Show loading overlay
    showDatabaseLoadingOverlay();
    
    // Calls the python function 'select_folder'
    window.pywebview.api.select_folder('db').then(path => {
        const label = document.getElementById('lbl-db');
        label.innerText = path;
        label.style.color = path !== "Not Selected" ? '#006400' : '#333';
        
        // Hide loading overlay
        hideDatabaseLoadingOverlay();
        
        // Enable inputs when database folder is selected
        if (path !== "Not Selected") {
            enableQMEInputs();
            hideDBWarning();
            showToast('✅ Database loaded and ready!', 'success');
        } else {
            disableQMEInputs();
            showDBWarning();
        }
    }).catch(error => {
        hideDatabaseLoadingOverlay();
        showToast('❌ Error loading database: ' + error, 'error');
    });
}

function selectResult() {
    window.pywebview.api.select_folder('result').then(path => {
        const label = document.getElementById('lbl-res');
        label.innerText = path;
        label.style.color = path !== "Not Selected" ? '#006400' : '#333';
    });
}



function lookupSAPData() {
    const codSap = document.getElementById('cod_sap').value;
    const planta = document.getElementById('planta').value;
    const origem = document.getElementById('origem').value;
    const destino = document.getElementById('destino').value;
    
    // Validação do código SAP/IMS
    if (!codSap || codSap.trim() === '') {
        return; // Não faz nada se o campo estiver vazio
    }
    
    // Mostra indicador de carregamento nos campos auto
    showLoadingInFields();
    
    window.pywebview.api.lookup_sap_data(codSap, planta, origem, destino).then(response => {
        hideLoadingInFields();
        
        if (response.status === 'success') {
            const data = response.data;
            // Preenche os campos automaticamente
            if (data['Nome Fornecedor']) document.getElementById('fornecedor').value = data['Nome Fornecedor'];
            if (data.Transportadora) document.getElementById('transportadora').value = data.Transportadora;
            if (data['Estado Fornecedor']) document.getElementById('uf').value = data['Estado Fornecedor'];
            if (data['Veiculo a ser Utilizado']) document.getElementById('veiculo').value = data['Veiculo a ser Utilizado'];
            if (data['Cidade Fornecedor']) document.getElementById('origem').value = data['Cidade Fornecedor'];
            if (data['Destino Materiais']) document.getElementById('destino').value = data['Destino Materiais'];
            if (data['Tipo de Fluxo']) document.getElementById('fluxo').value = data['Tipo de Fluxo'];
            
            // Mostra mensagem de sucesso
            showToast('✅ Dados carregados com sucesso!', 'success');
        } else if (response.status === 'not_found') {
            showToast('⚠️ ' + response.message, 'warning');
        } else if (response.status === 'error') {
            showToast('❌ ' + response.message, 'error');
        }
    }).catch(error => {
        hideLoadingInFields();
        showToast('❌ Erro ao buscar dados: ' + error, 'error');
    });
}

function showLoadingInFields() {
    const fields = ['fornecedor', 'transportadora', 'veiculo', 'fluxo'];
    fields.forEach(fieldId => {
        const field = document.getElementById(fieldId);
        if (field) {
            field.value = '⏳ Carregando...';
            field.classList.add('loading');
        }
    });
}

function hideLoadingInFields() {
    const fields = ['fornecedor', 'transportadora', 'veiculo', 'fluxo'];
    fields.forEach(fieldId => {
        const field = document.getElementById(fieldId);
        if (field) {
            if (field.value === '⏳ Carregando...') {
                field.value = '';
            }
            field.classList.remove('loading');
        }
    });
}

function showToast(message, type) {
    const overlay = document.getElementById('toast-overlay');
    const toast = document.getElementById('toast-message');
    
    if (overlay && toast) {
        toast.innerText = message;
        toast.className = 'toast-message toast-' + type;
        overlay.classList.add('active');
        
        // Auto-hide após 5 segundos
        setTimeout(() => {
            overlay.classList.remove('active');
        }, 3000);
    }
}

// Função de debounce para auto-fetch SAP data
function handleSAPInput(event) {
    // Verifica se o database foi selecionado
    if (!isDatabaseSelected()) {
        showToast('⚠️ Por favor, selecione a pasta Database primeiro!', 'warning');
        event.target.value = '';
        return;
    }
    
    // Limpa o timeout anterior
    if (sapLookupTimeout) {
        clearTimeout(sapLookupTimeout);
    }
    
    const codSap = event.target.value;
    
    // Se o campo estiver vazio, não faz nada
    if (!codSap || codSap.trim() === '') {
        return;
    }
    
    // Aguarda 2 segundos após o usuário parar de digitar
    sapLookupTimeout = setTimeout(() => {
        lookupSAPData();
    }, 2000); // 2 segundos
}

function isDatabaseSelected() {
    const dbLabel = document.getElementById('lbl-db');
    return dbLabel && dbLabel.innerText !== "Not Selected";
}

function enableQMEInputs() {
    const inputs = ['cod_projeto', 'cod_sap', 'planta', 'origem', 'destino'];
    inputs.forEach(inputId => {
        const input = document.getElementById(inputId);
        if (input) {
            input.disabled = false;
            input.classList.remove('disabled-input');
            // Remove o listener de clique de aviso
            input.removeEventListener('click', showDatabaseRequiredMessage);
        }
    });
}

function disableQMEInputs() {
    const inputs = ['cod_projeto', 'cod_sap', 'planta', 'origem', 'destino'];
    inputs.forEach(inputId => {
        const input = document.getElementById(inputId);
        if (input) {
            input.disabled = true;
            input.classList.add('disabled-input');
            // Adiciona listener para mostrar mensagem ao clicar
            input.addEventListener('click', showDatabaseRequiredMessage);
        }
    });
}

function showDatabaseRequiredMessage() {
    showToast('📂 Por favor, selecione a pasta Database Folder primeiro!', 'warning');
}

function showDBWarning() {
    const warning = document.getElementById('db-warning');
    if (warning) {
        warning.style.display = 'block';
    }
}

function hideDBWarning() {
    const warning = document.getElementById('db-warning');
    if (warning) {
        warning.style.display = 'none';
    }
}




function runSimulation() {
    // Gather inputs from the QME form
    const data = {
        cod_projeto: document.getElementById('cod_projeto').value,
        cod_sap: document.getElementById('cod_sap').value,
        fornecedor: document.getElementById('fornecedor').value,
        planta: document.getElementById('planta').value,
        origem: document.getElementById('origem').value,
        destino: document.getElementById('destino').value,
        fluxo: document.getElementById('fluxo').value,
        transportadora: document.getElementById('transportadora').value,
        veiculo: document.getElementById('veiculo').value,
        qme_tobe: document.getElementById('qme_tobe').value
    };

    // Send to Python backend
    window.pywebview.api.calculate_qme(data).then(response => {
        if (response.status === 'error') {
            alert(response.message);
            return;
        }
        
        // Display results in dashboard
        displayResults(response);
        
        // Switch to dashboard tab automatically
        switchTab('dash');
    });
}

function importASIS() {
    console.log('importASIS function called');
    
    if (!window.pywebview || !window.pywebview.api) {
        alert('API ainda não está pronta. Por favor, aguarde um momento.');
        return;
    }
    
    console.log('Calling import_asis_file API...');
    
    window.pywebview.api.import_asis_file().then(result => {
        console.log('API response:', result);
        const statusDiv = document.getElementById('asis-status');
        const fileInfo = document.getElementById('asis-file-info');
        const sampleDiv = document.getElementById('asis-sample');
        const asisDetails = document.getElementById('asis-details');
        const tobeDetails = document.getElementById('tobe-details');
        
        if (result.status === 'success') {
            // Show the status container
            statusDiv.style.display = 'grid';
            fileInfo.style.display = 'block';
            
            // File info
            fileInfo.innerHTML = `✅ <strong>${result.filename}</strong> - ${result.message}`;
            
            if (result.details && result.details.stats) {
                const stats = result.details.stats;
                
                // AS IS Section
                let asisMsg = '';
                if (stats.AS_IS_QME_Total !== undefined) {
                    asisMsg += `• QME Total: <strong>${stats.AS_IS_QME_Total.toLocaleString('pt-BR')}</strong><br>`;
                }
                if (stats.AS_IS_MDR_Distinct && stats.AS_IS_MDR_Distinct.length > 0) {
                    asisMsg += `• MDR Distintos: <strong>${stats.AS_IS_MDR_Distinct.length}</strong><br>`;
                }
                asisDetails.innerHTML = asisMsg;
                
                // TO BE Section
                let tobeMsg = '';
                if (stats.TO_BE_QME_Total !== undefined) {
                    tobeMsg += `• QME Total: <strong>${stats.TO_BE_QME_Total.toLocaleString('pt-BR')}</strong><br>`;
                }
                if (stats.TO_BE_MDR_Distinct && stats.TO_BE_MDR_Distinct.length > 0) {
                    tobeMsg += `• MDR Distintos: <strong>${stats.TO_BE_MDR_Distinct.length}</strong><br>`;
                }
                tobeDetails.innerHTML = tobeMsg;
            }
            
            // Sample PNs
            if (result.details && result.details.sample_pns && result.details.sample_pns.length > 0) {
                sampleDiv.style.display = 'block';
                sampleDiv.innerHTML = `📦 Exemplos de PNs: ${result.details.sample_pns.join(', ')}`;
            }
        } else if (result.status === 'error') {
            statusDiv.style.display = 'none';
            fileInfo.style.display = 'block';
            fileInfo.innerHTML = `❌ Erro: ${result.message}`;
            fileInfo.style.color = 'red';
        } else if (result.status === 'cancel') {
            statusDiv.style.display = 'none';
            fileInfo.style.display = 'block';
            fileInfo.innerHTML = '⚠️ Seleção cancelada';
            fileInfo.style.color = 'orange';
        } else {
            statusDiv.style.display = 'none';
            fileInfo.style.display = 'none';
            sampleDiv.style.display = 'none';
        }
    }).catch(error => {
        console.error('Error calling API:', error);
        const statusDiv = document.getElementById('asis-status');
        const fileInfo = document.getElementById('asis-file-info');
        statusDiv.style.display = 'none';
        fileInfo.style.display = 'block';
        fileInfo.innerHTML = `❌ Erro: ${error}`;
        fileInfo.style.color = 'red';
    });
}

function displayResults(response) {
    // Hide placeholder and show results section in dashboard
    const placeholder = document.getElementById('dashboard-placeholder');
    const resultsSection = document.getElementById('dashboard-results-section');
    
    if (placeholder) placeholder.style.display = 'none';
    if (resultsSection) resultsSection.style.display = 'block';
    
    // Update summary in dashboard
    document.getElementById('dashboard-summary-rows').innerText = response.summary.total_rows;
    document.getElementById('dashboard-summary-savings').innerText = 'R$ ' + response.summary.total_savings.toLocaleString('pt-BR');
    
    // Populate results table in dashboard
    const tbody = document.getElementById('dashboard-results-body');
    tbody.innerHTML = ''; // Clear previous results
    
    response.results.forEach(row => {
        const tr = document.createElement('tr');
        tr.style.borderBottom = '1px solid #eee';
        
        tr.innerHTML = `
            <td style="padding: 8px;">${row.row}</td>
            <td style="padding: 8px;">${row.pn}</td>
            <td style="padding: 8px; text-align: center;">${row.qme_asis}</td>
            <td style="padding: 8px; text-align: center;">${row.mdr_asis || '-'}</td>
            <td style="padding: 8px; text-align: center;">${row.qme_tobe}</td>
            <td style="padding: 8px; text-align: center;">${row.mdr_tobe || '-'}</td>
            <td style="padding: 8px; text-align: center;">${row.vol_asis}</td>
            <td style="padding: 8px; text-align: center;">${row.vol_tobe}</td>
            <td style="padding: 8px; text-align: center; color: green; font-weight: bold;">R$ ${row.savings.toLocaleString('pt-BR')}</td>
            <td style="padding: 8px; text-align: center;">
                <span style="background: ${row.status === 'OK' ? '#d4edda' : '#f8d7da'}; 
                             color: ${row.status === 'OK' ? '#155724' : '#721c24'}; 
                             padding: 3px 8px; border-radius: 3px; font-size: 0.75rem;">
                    ${row.status}
                </span>
            </td>
        `;
        
        tbody.appendChild(tr);
    });
}

function exportResults() {
    window.pywebview.api.export_results().then(response => {
        if (response.status === 'success') {
            alert('✅ ' + response.message);
        } else {
            alert('❌ ' + response.message);
        }
    });
}

// Event Listeners for initial load can go here if needed
window.addEventListener('pywebviewready', function() {
    console.log('PyWebview API is ready');
    
    // Adiciona listener para auto-fetch no campo SAP
    const sapInput = document.getElementById('cod_sap');
    if (sapInput) {
        sapInput.addEventListener('input', handleSAPInput);
        console.log('SAP auto-fetch listener attached');
    }
    
    // Inicializa o estado dos inputs baseado na seleção do database
    initializeInputState();
});

// Fallback: adiciona listener quando DOM estiver pronto
document.addEventListener('DOMContentLoaded', function() {
    const sapInput = document.getElementById('cod_sap');
    if (sapInput && !sapInput.hasAttribute('data-listener-attached')) {
        sapInput.addEventListener('input', handleSAPInput);
        sapInput.setAttribute('data-listener-attached', 'true');
        console.log('SAP auto-fetch listener attached (DOMContentLoaded)');
    }
    
    // Inicializa o estado dos inputs
    initializeInputState();
});

// Database Loading Overlay Functions
function showDatabaseLoadingOverlay() {
    let overlay = document.getElementById('db-loading-overlay');
    if (!overlay) {
        // Create overlay if it doesn't exist
        overlay = document.createElement('div');
        overlay.id = 'db-loading-overlay';
        overlay.className = 'db-loading-overlay';
        overlay.innerHTML = `
            <div class="db-loading-content">
                <div class="spinner"></div>
                <h3>📂 Loading Database Files...</h3>
                <small>This may take a few seconds on first load</small>
            </div>
        `;
        document.body.appendChild(overlay);
    }
    overlay.classList.add('active');
}

function hideDatabaseLoadingOverlay() {
    const overlay = document.getElementById('db-loading-overlay');
    if (overlay) {
        overlay.classList.remove('active');
    }
}

function initializeInputState() {
    // Verifica se o database foi selecionado
    if (isDatabaseSelected()) {
        enableQMEInputs();
        hideDBWarning();
    } else {
        disableQMEInputs();
        showDBWarning();
    }
}