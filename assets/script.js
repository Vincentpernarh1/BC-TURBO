// assets/script.js

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
        navBar.classList.remove('show'); // Hide nav-bar show
    } else if(tabName === 'freq') {
        document.querySelector('#mod-freq').classList.add('active');
        document.querySelectorAll('.nav-item')[1].classList.add('active');
        navBar.classList.remove('show'); // Hide nav-bar
    } else if(tabName === 'emb') {
        document.querySelector('#mod-emb').classList.add('active');
        document.querySelectorAll('.nav-item')[2].classList.add('active');
        navBar.classList.remove('show'); // Hide nav-bar
    } else if(tabName === 'dash') {
        document.querySelector('#mod-dash').classList.add('active');
        document.querySelectorAll('.nav-item')[3].classList.add('active');
        navBar.classList.add('show'); // Show nav-bar on dashboard
    }
}

// PYTHON INTERACTION (Calls the backend API)
function selectDB() {
    // Calls the python function 'select_folder'
    window.pywebview.api.select_folder('db').then(path => {
        const label = document.getElementById('lbl-db');
        label.innerText = path;
        label.style.color = path !== "Not Selected" ? '#006400' : '#333';
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
    
    if (!codSap || !planta || !origem || !destino) {
        return; // Aguarda todos os campos necessários
    }
    
    window.pywebview.api.lookup_sap_data(codSap, planta, origem, destino).then(response => {
        if (response.status === 'success') {
            const data = response.data;
            document.getElementById('fornecedor').value = data.fornecedor || '';
            document.getElementById('transportadora').value = data.transportadora || '';
            document.getElementById('veiculo').value = data.veiculo || '';
            document.getElementById('uf').value = data.uf || '';
        }
    });
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
        uf: document.getElementById('uf').value,
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
        
        if (result.status === 'success') {
            let detailsMsg = `<strong>${result.filename}</strong><br>${result.message}<br>`;
            
            if (result.details && result.details.stats) {
                const stats = result.details.stats;
                
                detailsMsg += '<div style="margin-top: 10px; padding: 10px; background: #f0f8ff; border-radius: 5px;">';
                detailsMsg += '<strong>📊 Resumo dos Dados:</strong><br>';
                
                // AS IS Section
                detailsMsg += '<div style="margin-top: 8px;">';
                detailsMsg += '<strong style="color: #0066cc;">AS IS:</strong><br>';
                if (stats.AS_IS_QME_Total !== undefined) {
                    detailsMsg += `&nbsp;&nbsp;• QME Total: <strong>${stats.AS_IS_QME_Total.toLocaleString('pt-BR')}</strong><br>`;
                }
                if (stats.AS_IS_MDR_Distinct && stats.AS_IS_MDR_Distinct.length > 0) {
                    detailsMsg += `&nbsp;&nbsp;• MDR Distintos: <strong>${stats.AS_IS_MDR_Distinct.length}</strong><br>`;
                }
                detailsMsg += '</div>';
                
                // TO BE Section
                detailsMsg += '<div style="margin-top: 8px;">';
                detailsMsg += '<strong style="color: #cc6600;">TO BE:</strong><br>';
                if (stats.TO_BE_QME_Total !== undefined) {
                    detailsMsg += `&nbsp;&nbsp;• QME Total: <strong>${stats.TO_BE_QME_Total.toLocaleString('pt-BR')}</strong><br>`;
                }
                if (stats.TO_BE_MDR_Distinct && stats.TO_BE_MDR_Distinct.length > 0) {
                    detailsMsg += `&nbsp;&nbsp;• MDR Distintos: <strong>${stats.TO_BE_MDR_Distinct.length}</strong><br>`;
                }
                detailsMsg += '</div>';
                
                detailsMsg += '</div>';
            }
            
            // Sample PNs
            if (result.details && result.details.sample_pns && result.details.sample_pns.length > 0) {
                detailsMsg += `<div style="margin-top: 8px; font-size: 0.85rem; color: #666;">📦 Exemplos: ${result.details.sample_pns.join(', ')}</div>`;
            }
            
            statusDiv.innerHTML = `✅ ${detailsMsg}`;
            statusDiv.style.color = 'green';
            statusDiv.style.fontSize = '0.9rem';
            statusDiv.style.marginTop = '10px';
        } else if (result.status === 'error') {
            statusDiv.innerText = `❌ Erro: ${result.message}`;
            statusDiv.style.color = 'red';
        } else if (result.status === 'cancel') {
            statusDiv.innerText = '⚠️ Seleção cancelada';
            statusDiv.style.color = 'orange';
        } else {
            statusDiv.innerText = '';
        }
    }).catch(error => {
        console.error('Error calling API:', error);
        const statusDiv = document.getElementById('asis-status');
        statusDiv.innerText = `❌ Erro: ${error}`;
        statusDiv.style.color = 'red';
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
});