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
            if (data['Estado Fornecedor']) document.getElementById('uf').value = data['Estado Fornecedor'];
            
            // Preenche origem - pode ser Cidade Fornecedor ou CrossDock
            if (data['Cidade Fornecedor']) {
                document.getElementById('origem').value = data['Cidade Fornecedor'];
            } else if (data['CrossDock']) {
                document.getElementById('origem').value = data['CrossDock'];
            }
            
            // Verifica se TDC precisa de destino IMS
            if (response.is_milk_run_or_line_haul) {
                // --- MODO MILK RUN / LINE HAUL ---
                // Restaura transportadora para select (em caso de a última busca ter virado texto)
                resetToStandardMode();
                // Configura campos especiais
                setMilkRunLineHaulMode(response.normalized_fluxo, response.tdc_options || {});
                showToast(`✅ Fluxo ${response.normalized_fluxo} detectado — campos configurados automaticamente!`, 'success');
            } else if (response.tdc_needs_destino) {
                // Marca o campo destino como requerido (vermelho)
                const destinoField = document.getElementById('destino');
                destinoField.classList.add('required-field');
                destinoField.placeholder = 'Digite o Código IMS Destino';
                
                // Desabilita os dropdowns TDC
                disableTDCDropdowns();
                
                // Mostra mensagem ao usuário
                showToast('⚠️ Digite o Código IMS Destino para carregar dados TDC completos', 'warning');
                
                // Se temos cod_ims_origem, pode mostrar em algum lugar para referência
                if (response.cod_ims_origem) {
                    console.log('IMS Origem:', response.cod_ims_origem);
                }
            } else {
                // Remove marcação de campo requerido
                const destinoField = document.getElementById('destino');
                destinoField.classList.remove('required-field');
                destinoField.placeholder = '';
                
                // Garante modo padrão (select visível) para fluxo não-ML/LH
                resetToStandardMode();
                
                // Popula dropdowns TDC se temos opções
                if (response.tdc_options) {
                    populateTDCDropdowns(response.tdc_options, data);
                    
                    // Habilita dropdowns TDC
                    enableTDCDropdowns();
                }
                
                // Mostra mensagem de sucesso completo
                showToast('✅ Dados carregados com sucesso!', 'success');
            }
            
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

function populateTDCDropdowns(options, data) {
    // Popula Transportadora
    const transportadoraSelect = document.getElementById('transportadora');
    transportadoraSelect.innerHTML = '<option value="">Selecione...</option>';
    if (options.Transportadora && options.Transportadora.length > 0) {
        options.Transportadora.forEach(option => {
            const opt = document.createElement('option');
            opt.value = option;
            opt.textContent = option;
            // Marca como selecionada se for o valor padrão
            if (data.Transportadora === option) {
                opt.selected = true;
            }
            transportadoraSelect.appendChild(opt);
        });
    }
    
    // Popula Veiculo
    const veiculoSelect = document.getElementById('veiculo');
    veiculoSelect.innerHTML = '<option value="">Selecione...</option>';
    if (options.Veiculo && options.Veiculo.length > 0) {
        options.Veiculo.forEach(option => {
            const opt = document.createElement('option');
            opt.value = option;
            opt.textContent = option;
            // Marca como selecionada se for o valor padrão
            if (data.Veiculo === option) {
                opt.selected = true;
            }
            veiculoSelect.appendChild(opt);
        });
    }
    
    // Popula Fluxo
    const fluxoSelect = document.getElementById('fluxo');
    fluxoSelect.innerHTML = '<option value="">Selecione...</option>';
    if (options.Fluxo && options.Fluxo.length > 0) {
        options.Fluxo.forEach(option => {
            const opt = document.createElement('option');
            opt.value = option;
            opt.textContent = option;
            // Marca como selecionada se for o valor padrão
            if (data['Fluxo Viagem'] === option) {
                opt.selected = true;
            }
            fluxoSelect.appendChild(opt);
        });
    }
    
    // Popula Trip
    const tripSelect = document.getElementById('trip');
    tripSelect.innerHTML = '<option value="">Selecione...</option>';
    if (options.Trip && options.Trip.length > 0) {
        options.Trip.forEach(option => {
            const opt = document.createElement('option');
            opt.value = option;
            opt.textContent = option;
            // Marca como selecionada se for o valor padrão
            if (data.Trip === option) {
                opt.selected = true;
            }
            tripSelect.appendChild(opt);
        });
    }
}

function setMilkRunLineHaulMode(normalizedFluxo, options) {
    /**
     * Configura o formulário para modo Milk Run / Line Haul:
     *  - Transportadora vira campo de texto livre (não há opções no TDC)
     *  - Veículo = "Carreta" (constante, desabilitado)
     *  - Fluxo = dropdown com "Milk Run" / "Line Haul", pré-selecionado
     *  - Trip = dropdown com "Round Trip (RT)" / "One Way (OW)"
     */
    // --- Transportadora: esconde select, mostra input de texto ---
    const transportadoraSelect = document.getElementById('transportadora');
    const transportadoraText   = document.getElementById('transportadora-text');
    if (transportadoraSelect && transportadoraText) {
        transportadoraSelect.style.display = 'none';
        transportadoraText.style.display   = '';
        transportadoraText.disabled = false;
        transportadoraText.classList.remove('disabled-input');
        transportadoraText.value = '';
    }

    // --- Veículo: fixo em "Carreta", desabilitado ---
    const veiculoSelect = document.getElementById('veiculo');
    if (veiculoSelect) {
        veiculoSelect.innerHTML = '<option value="Carreta">Carreta</option>';
        veiculoSelect.value = 'Carreta';
        veiculoSelect.disabled = true;
        veiculoSelect.classList.add('disabled-input');
    }

    // --- Fluxo: opções fixas, pré-seleciona o fluxo do PFEP ---
    const fluxoSelect = document.getElementById('fluxo');
    if (fluxoSelect) {
        fluxoSelect.innerHTML = '<option value="">Selecione...</option>';
        (options.Fluxo || ['Milk Run', 'Line Haul']).forEach(opt => {
            const el = document.createElement('option');
            el.value = opt;
            el.textContent = opt;
            if (opt === normalizedFluxo) el.selected = true;
            fluxoSelect.appendChild(el);
        });
        fluxoSelect.disabled = false;
        fluxoSelect.classList.remove('disabled-input');
    }

    // --- Trip: opções fixas ---
    const tripSelect = document.getElementById('trip');
    if (tripSelect) {
        tripSelect.innerHTML = '<option value="">Selecione...</option>';
        (options.Trip || ['Round Trip', 'One Way']).forEach(opt => {
            const el = document.createElement('option');
            el.value = opt;
            el.textContent = opt;
            tripSelect.appendChild(el);
        });
        tripSelect.disabled = false;
        tripSelect.classList.remove('disabled-input');
    }
}

function resetToStandardMode() {
    /**
     * Restaura campos para o modo padrão (TDC-driven) quando o fluxo
     * não é Milk Run nem Line Haul.
     */
    const transportadoraSelect = document.getElementById('transportadora');
    const transportadoraText   = document.getElementById('transportadora-text');
    if (transportadoraSelect && transportadoraText) {
        transportadoraSelect.style.display = '';
        transportadoraText.style.display   = 'none';
        transportadoraText.value = '';
    }
}

async function prepareViajanteData(codSap, cidadeDestino) {
    /**
     * Prepara dados para integração com Viajante (apenas cria arquivo de demanda)
     * Cria arquivo no formato: Mês, COD FORNECEDOR, DESENHO, QTDE
     * O processamento Viajante acontece quando clicar "Executar Simulação"
     */
    if (!codSap || codSap.trim() === '') {
        return { status: 'error', message: 'Código SAP vazio' };
    }
    
    try {
        // Get additional parameters from form
        const veiculo = document.getElementById('veiculo').value || '';
        
        // Prepare demanda file ONLY (no Viajante processing yet)
        console.log('📊 Preparing Viajante demand data...');
        const prepareResponse = await window.pywebview.api.prepare_viajante_data(codSap, cidadeDestino, veiculo);
        
        if (prepareResponse.status === 'success') {
            console.log('✅ Viajante demand file created successfully');
            console.log(`   File: ${prepareResponse.file_path}`);
            console.log(`   Total rows: ${prepareResponse.total_rows}`);
            console.log(`   Unique PNs: ${prepareResponse.unique_pns}`);
            console.log(`   Months: ${prepareResponse.months}`);
        } else {
            console.error('❌ Error preparing Viajante data:', prepareResponse.message);
        }
        
        return prepareResponse;
        
    } catch (error) {
        console.error('❌ Error in prepareViajanteData:', error);
        return { status: 'error', message: error.message };
    }
}

function enableTDCDropdowns() {
    const dropdowns = ['transportadora', 'veiculo', 'fluxo', 'trip'];
    dropdowns.forEach(id => {
        const dropdown = document.getElementById(id);
        if (dropdown) {
            dropdown.disabled = false;
            dropdown.classList.remove('disabled-input');
        }
    });
}

function disableTDCDropdowns() {
    const dropdowns = ['transportadora', 'veiculo', 'fluxo', 'trip'];
    dropdowns.forEach(id => {
        const dropdown = document.getElementById(id);
        if (dropdown) {
            dropdown.disabled = true;
            dropdown.classList.add('disabled-input');
            dropdown.innerHTML = '<option value="">Selecione...</option>';
        }
    });
}

function showLoadingInFields() {
    const fields = ['fornecedor'];
    fields.forEach(fieldId => {
        const field = document.getElementById(fieldId);
        if (field) {
            field.value = '⏳ Carregando...';
            field.classList.add('loading');
        }
    });
    
    // Desabilita dropdowns durante carregamento
    disableTDCDropdowns();
}

function hideLoadingInFields() {
    const fields = ['fornecedor'];
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

// Função de debounce para quando o usuário digita IMS Destino
let destinoLookupTimeout = null;

function handleDestinoInput(event) {
    // Verifica se o database foi selecionado
    if (!isDatabaseSelected()) {
        showToast('⚠️ Por favor, selecione a pasta Database primeiro!', 'warning');
        return;
    }
    
    // Verifica se há um código SAP já digitado
    const codSap = document.getElementById('cod_sap').value;
    if (!codSap || codSap.trim() === '') {
        showToast('⚠️ Por favor, digite o Código SAP/IMS primeiro!', 'warning');
        return;
    }
    
    // Limpa o timeout anterior
    if (destinoLookupTimeout) {
        clearTimeout(destinoLookupTimeout);
    }
    
    const destino = event.target.value;
    
    // Se o campo estiver vazio, não faz nada
    if (!destino || destino.trim() === '') {
        return;
    }
    
    // Aguarda 2 segundos após o usuário parar de digitar
    destinoLookupTimeout = setTimeout(() => {
        // Trigger novo lookup com o destino preenchido
        lookupSAPData();
    }, 2000); // 2 segundos
}

function isDatabaseSelected() {
    const dbLabel = document.getElementById('lbl-db');
    return dbLabel && dbLabel.innerText !== "Not Selected";
}

function enableQMEInputs() {
    const inputs = ['cod_projeto', 'cod_sap', 'planta', 'origem', 'destino', 'trip', 'rt_percent', 'pedagio', 'km_manual'];
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
    const inputs = ['cod_projeto', 'cod_sap', 'planta', 'origem', 'destino', 'trip', 'rt_percent', 'pedagio', 'km_manual'];
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




async function runSimulation() {
    // Gather inputs from the QME form
    const transportadoraText   = document.getElementById('transportadora-text');
    const transportadoraSelect = document.getElementById('transportadora');
    const transportadoraValue  = (transportadoraText && transportadoraText.style.display !== 'none')
        ? transportadoraText.value
        : (transportadoraSelect ? transportadoraSelect.value : '');

    const data = {
        cod_projeto: document.getElementById('cod_projeto').value,
        cod_sap: document.getElementById('cod_sap').value,
        fornecedor: document.getElementById('fornecedor').value,
        planta: document.getElementById('planta').value,
        origem: document.getElementById('origem').value,
        destino: document.getElementById('destino').value,
        fluxo: document.getElementById('fluxo').value,
        transportadora: transportadoraValue,
        veiculo: document.getElementById('veiculo').value,
        trip: document.getElementById('trip').value,
        qme_tobe: document.getElementById('qme_tobe').value,
        rt_percent: parseFloat(document.getElementById('rt_percent').value) || 100,
        pedagio: parseFloat(document.getElementById('pedagio').value) || 0,
        km_manual: parseFloat(document.getElementById('km_manual').value) || 0
    };

    // Validate required fields
    if (!data.cod_sap || data.cod_sap.trim() === '') {
        showToast('⚠️ Por favor, preencha o Código SAP/IMS', 'warning');
        return;
    }

    if (!data.destino || data.destino.trim() === '') {
        showToast('⚠️ Por favor, preencha o Código IMS Destino', 'warning');
        return;
    }

    // Detect Milk Run / Line Haul mode — Viajante is not needed in this mode
    const fluxoValue = data.fluxo ? data.fluxo.toLowerCase() : '';
    const isMilkRunOrLineHaul = fluxoValue.includes('milk run') || fluxoValue.includes('line haul');

    try {
        let viajanteResponse = null;

        if (isMilkRunOrLineHaul) {
            // ── ML/LH MODE: skip Viajante entirely ──
            console.log(`\n🚛 ${data.fluxo} mode detected — skipping Viajante processing`);
            showToast(`⏳ Modo ${data.fluxo}: calculando sem Viajante...`, 'info');
        } else {
            // ── STANDARD MODE: prepare demand + run Viajante ──
            showToast('⏳ Preparando dados de demanda...', 'info');

            // Step 1: Prepare Viajante demand file
            console.log('\n📊 Step 1: Preparing Viajante demand file...');
            const prepareResponse = await prepareViajanteData(data.cod_sap, data.destino);

            if (prepareResponse && prepareResponse.status === 'error') {
                showToast('❌ Erro ao preparar demanda: ' + prepareResponse.message, 'error');
                return;
            }

            // Step 2: Run Viajante processing
            console.log('\n🚚 Step 2: Running Viajante processing...');
            showToast('⏳ Processando Viajante...', 'info');

            viajanteResponse = await window.pywebview.api.run_viajante(data.cod_sap);

            if (viajanteResponse.status !== 'success') {
                console.error('❌ Error running Viajante:', viajanteResponse.message);
                showToast('❌ Erro Viajante: ' + viajanteResponse.message, 'error');
                return;
            }

            console.log('✅ Viajante processing completed successfully');
            console.log(`   Results: ${viajanteResponse.total_rows} rows`);
            showToast('✅ Viajante processado com sucesso!', 'success');
        }

        // Step 3: Calculate QME
        console.log('\n📊 Step 3: Calculating QME...');
        showToast('⏳ Processando simulação QME...', 'info');

        const qmeResponse = await window.pywebview.api.calculate_qme(data);

        if (qmeResponse.status === 'error') {
            showToast('❌ Erro QME: ' + qmeResponse.message, 'error');
            return;
        }

        // Display QME results in dashboard
        displayResults(qmeResponse);
        showToast('✅ Cálculo QME concluído!', 'success');

        // Display Viajante results (only in standard mode)
        if (viajanteResponse) {
            displayViajanteResults(viajanteResponse.results);
        }

        // Update weekly trips in breakdown table
        console.log('\n📊 Step 4: Updating weekly trips in breakdown table...');
        updateWeeklyTrips(qmeResponse, viajanteResponse);

        // Switch to dashboard
        console.log('\n✅ All processing complete! Switching to dashboard...');
        switchTab('dash');
        
    } catch (error) {
        console.error('❌ Error in runSimulation:', error);
        showToast('❌ Erro na simulação: ' + error.message, 'error');
    }
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
    // Update summary cards with total rows from Astobe file
    document.getElementById('dashboard-summary-rows').innerText = response.summary.total_rows || 0;
    document.getElementById('dashboard-matched-pns').innerText = response.summary.matched_rows || 0;
    document.getElementById('dashboard-unmatched-pns').innerText = response.summary.unmatched_rows || 0;
    document.getElementById('dashboard-summary-savings').innerText = 'R$ ' + (response.summary.saving_12_meses || 0).toLocaleString('pt-BR');
    
    const months = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
                   'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
    
    // Month mapping from display names to backend keys
    const monthKeys = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 
                       'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
    
    // Populate Combined Breakdown Table (AS IS vs TO BE side by side)
    const breakdownCombinedBody = document.getElementById('breakdown-combined-body');
    if (breakdownCombinedBody) {
        breakdownCombinedBody.innerHTML = '';
        
        // Volume m³ row - NOW POPULATED WITH ACTUAL DATA
        const volumeRow = document.createElement('tr');
        volumeRow.innerHTML = '<td class="td-label-bold">Volume m³</td>';
        let totalM3Asis = 0;
        let totalM3Tobe = 0;
        monthKeys.forEach(monthKey => {
            const m3Asis = response.summary.monthly_m3_asis?.[monthKey] || 0;
            const m3Tobe = response.summary.monthly_m3_tobe?.[monthKey] || 0;
            totalM3Asis += m3Asis;
            totalM3Tobe += m3Tobe;
            // FIXED: Swapped order - TO BE first, AS IS second to match expected behavior
            volumeRow.innerHTML += `<td class="td-breakdown-as">${m3Asis.toFixed(2)}</td><td class="td-breakdown-to">${m3Tobe.toFixed(2)}</td>`;
        });
        volumeRow.innerHTML += `<td class="td-breakdown-as">${totalM3Asis.toFixed(2)}</td><td class="td-breakdown-to">${totalM3Tobe.toFixed(2)}</td>`;
        breakdownCombinedBody.appendChild(volumeRow);
        
        // Qtde de viagens row - Will be populated by updateWeeklyTrips() after Viajante completes
        const viagensRow = document.createElement('tr');
        const veiculoName = response.veiculo || 'VEÍCULO';
        viagensRow.innerHTML = `<td class="td-label">Qtde de Viagens Semanal</td>`;
        
        monthKeys.forEach(monthKey => {
            // Placeholder values - will be updated after Viajante processing
            viagensRow.innerHTML += '<td class="td-breakdown-as">-</td><td class="td-breakdown-to">-</td>';
        });
        
        viagensRow.innerHTML += '<td class="td-breakdown-as">-</td><td class="td-breakdown-to">-</td>';
        breakdownCombinedBody.appendChild(viagensRow);
        
        // Custo mensal de frete row (trips × tarifa_real)
        const custoSemanalRow = document.createElement('tr');
        custoSemanalRow.innerHTML = `<td class="td-label">Custo Mensal de Frete (${veiculoName})</td>`;
        months.forEach(month => {
            custoSemanalRow.innerHTML += '<td class="td-breakdown-as">R$ -</td><td class="td-breakdown-to">R$ -</td>';
        });
        custoSemanalRow.innerHTML += '<td class="td-breakdown-total">R$ -</td><td class="td-breakdown-total">R$ -</td>';
        breakdownCombinedBody.appendChild(custoSemanalRow);

        // Pedágio row — AS IS and TO BE per month (trips × pedagio per trip)
        const pedagioRow = document.createElement('tr');
        pedagioRow.innerHTML = `<td class="td-label">Pedágio</td>`;
        months.forEach(() => {
            pedagioRow.innerHTML += '<td class="td-breakdown-as">R$ -</td><td class="td-breakdown-to">R$ -</td>';
        });
        pedagioRow.innerHTML += '<td class="td-breakdown-total">R$ -</td><td class="td-breakdown-total">R$ -</td>';
        breakdownCombinedBody.appendChild(pedagioRow);
        
        // Economia mensal de frete row — single column per month (savings = AS IS − TO BE)
        const custoTotalRow = document.createElement('tr');
        custoTotalRow.innerHTML = `<td class="td-label">Economia Mensal de Frete</td>`;
        months.forEach(() => {
            custoTotalRow.innerHTML += '<td class="td-breakdown-as" colspan="2">R$ -</td>';
        });
        custoTotalRow.innerHTML += '<td class="td-breakdown-total" colspan="2">R$ -</td>';
        breakdownCombinedBody.appendChild(custoTotalRow);
        
        // TOTAL SEMANAL row — single column per month (Frete TO BE + Pedágio TO BE)
        const totalSemanalRow = document.createElement('tr');
        totalSemanalRow.innerHTML = '<td class="td-label-bold">TOTAL SEMANAL</td>';
        months.forEach(() => {
            totalSemanalRow.innerHTML += '<td class="td-breakdown-total-bold" colspan="2">R$ -</td>';
        });
        totalSemanalRow.innerHTML += '<td class="td-breakdown-total-bold" colspan="2">R$ -</td>';
        breakdownCombinedBody.appendChild(totalSemanalRow);
        
        // TOTAL MENSAL row — single column per month (TOTAL SEMANAL × 4)
        const totalMensalRow = document.createElement('tr');
        totalMensalRow.innerHTML = '<td class="td-label-bold">TOTAL MENSAL</td>';
        months.forEach(() => {
            totalMensalRow.innerHTML += '<td class="td-breakdown-total-bold" colspan="2">R$ -</td>';
        });
        totalMensalRow.innerHTML += '<td class="td-breakdown-total-bold" colspan="2">R$ -</td>';
        breakdownCombinedBody.appendChild(totalMensalRow);
    }
    
    // Populate monthly aggregation table with QME quantities
    const monthlyBody = document.getElementById('dashboard-monthly-body');
    if (monthlyBody) {
        monthlyBody.innerHTML = '';
        
        // QME AS IS Row (showing quantities, not monetary values)
        const qmeAsisRow = document.createElement('tr');
        qmeAsisRow.innerHTML = '<td class="td-label-bold">QME AS IS (unidades)</td>';
        months.forEach(month => {
            const value = response.summary.monthly_qme_asis?.[month] || 0;
            qmeAsisRow.innerHTML += `<td class="td-center">${value.toFixed(0)}</td>`;
        });
        qmeAsisRow.innerHTML += `<td class="td-total-anual">${(response.summary.total_qme_asis || 0).toFixed(0)}</td>`;
        monthlyBody.appendChild(qmeAsisRow);
        
        // QME TO BE Row (showing quantities, not monetary values)
        const qmeTobeRow = document.createElement('tr');
        qmeTobeRow.innerHTML = '<td class="td-label-bold">QME TO BE (unidades)</td>';
        months.forEach(month => {
            const value = response.summary.monthly_qme_tobe?.[month] || 0;
            qmeTobeRow.innerHTML += `<td class="td-center">${value.toFixed(0)}</td>`;
        });
        qmeTobeRow.innerHTML += `<td class="td-total-anual">${(response.summary.total_qme_tobe || 0).toFixed(0)}</td>`;
        monthlyBody.appendChild(qmeTobeRow);
        
        // Separator row or space before savings
        const separatorRow = document.createElement('tr');
        separatorRow.innerHTML = '<td colspan="14" style="height: 10px; background-color: #f0f0f0;"></td>';
        monthlyBody.appendChild(separatorRow);
        
        // SAVING Row (monetary values)
        const savingRow = document.createElement('tr');
        savingRow.innerHTML = '<td class="td-label-saving">ECONOMIA MENSAL (R$)</td>';
        months.forEach(month => {
            // Calculate monthly savings based on volumes (placeholder - adjust as needed)
            const monthly_saving = (response.summary.saving_12_meses || 0) / 12;
            savingRow.innerHTML += `<td class="td-saving">R$ ${monthly_saving.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>`;
        });
        savingRow.innerHTML += `<td class="td-saving-total">R$ ${(response.summary.saving_12_meses || 0).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>`;
        monthlyBody.appendChild(savingRow);
    }
    
    // Populate detailed results table (initially hidden)
    const tbody = document.getElementById('dashboard-results-body');
    tbody.innerHTML = ''; // Clear previous results
    
    response.results.forEach(row => {
        const tr = document.createElement('tr');
        tr.className = 'table-row';
        
        // Highlight rows with PFEP match (green background)
        if (row.has_pfep_match) {
            tr.style.backgroundColor = '#e8f5e9';  // Light green for matched PNs
        } else {
            tr.style.backgroundColor = '#ffebee';  // Light red for unmatched PNs
        }
        
        // Status badge styling
        let statusClass = 'status-warning';
        if (row.status.includes('Matched - Improvement')) {
            statusClass = 'status-ok';
        } else if (row.status.includes('Matched')) {
            statusClass = 'status-info';
        }
        
        // Get monthly volumes
        const months = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
        const monthlyVolumesHTML = months.map(month => {
            const vol = row.monthly_volumes && row.monthly_volumes[month] ? row.monthly_volumes[month].toLocaleString('pt-BR', {minimumFractionDigits: 0, maximumFractionDigits: 0}) : '-';
            return `<td class="td-center">${vol}</td>`;
        }).join('');
        
        tr.innerHTML = `
            <td class="td-label">${row.row}</td>
            <td class="td-label"><strong>${row.pn}</strong></td>
            ${monthlyVolumesHTML}
            <td class="td-center">${row.qme_asis || '-'}</td>
            <td class="td-center">${row.mdr_asis || '-'}</td>
            <td class="td-center">${row.vol_asis_m3 ? row.vol_asis_m3.toFixed(4) : '-'}</td>
            <td class="td-center">${row.peso_asis_kg ? row.peso_asis_kg.toFixed(2) : '-'}</td>
            <td class="td-center">${row.qme_tobe || '-'}</td>
            <td class="td-center">${row.mdr_tobe || '-'}</td>
            <td class="td-center">${row.vol_tobe_m3 ? row.vol_tobe_m3.toFixed(4) : '-'}</td>
            <td class="td-center">${row.peso_tobe_kg ? row.peso_tobe_kg.toFixed(2) : '-'}</td>
            <!-- <td class="td-value">R$ ${row.savings.toLocaleString('pt-BR')}</td> -->
            <td class="td-center">
                <span class="status-badge ${statusClass}">
                    ${row.status}
                </span>
            </td>
        `;
        
        tbody.appendChild(tr);
    });
    
    // Add TOTALS row at the bottom of the detailed PN table
    const totalsRow = document.createElement('tr');
    totalsRow.className = 'table-row-totals';
    totalsRow.style.backgroundColor = '#f0f0f0';
    totalsRow.style.fontWeight = 'bold';
    totalsRow.style.borderTop = '3px solid #333';
    
    // Calculate totals for each column
    totalsRow.innerHTML = `
        <td class="td-label" colspan="2" style="text-align: right; padding-right: 10px;"><strong>TOTAIS:</strong></td>
    `;
    
    // Add monthly volume totals
    monthKeys.forEach(monthKey => {
        const totalVol = response.summary.monthly_qme_asis?.[monthKey] || 0;
        totalsRow.innerHTML += `<td class="td-center"><strong>${totalVol.toLocaleString('pt-BR', {minimumFractionDigits: 0, maximumFractionDigits: 0})}</strong></td>`;
    });
    
    // Add empty cells for QME, MDR, Vol, Peso columns (or show totals if needed)
    totalsRow.innerHTML += `
        <td class="td-center">-</td>
        <td class="td-center">-</td>
        <td class="td-center"><strong>${(response.summary.monthly_m3_asis ? Object.values(response.summary.monthly_m3_asis).reduce((a,b) => a+b, 0) : 0).toFixed(2)}</strong></td>
        <td class="td-center">-</td>
        <td class="td-center">-</td>
        <td class="td-center">-</td>
        <td class="td-center"><strong>${(response.summary.monthly_m3_tobe ? Object.values(response.summary.monthly_m3_tobe).reduce((a,b) => a+b, 0) : 0).toFixed(2)}</strong></td>
        <td class="td-center">-</td>
       <!-- <td class="td-value"><strong>R$ ${response.summary.total_savings.toLocaleString('pt-BR')}</strong></td> -->
        <td class="td-center">-</td>
    `;
    
    tbody.appendChild(totalsRow);
    
    // Log matching info to console
    if (response.matching) {
        console.log('Matched PNs:', response.matching.matched);
        console.log('Unmatched PNs:', response.matching.unmatched);
    }
}

function toggleDetailedView() {
    const detailsSection = document.getElementById('detailed-results-section');
    const toggleBtn = document.getElementById('toggle-details-btn');
    
    if (detailsSection.classList.contains('hidden-section')) {
        detailsSection.classList.remove('hidden-section');
        toggleBtn.innerHTML = '🔼 Ocultar Detalhes por PN';
        // Scroll to detailed section
        detailsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
    } else {
        detailsSection.classList.add('hidden-section');
        toggleBtn.innerHTML = '🔽 Mostrar Detalhes por PN';
    }
}

function displayViajanteResults(results) {
    /**
     * Displays Viajante processing results (Volume por Rota data)
     * Shows: CARGAS, CAP. ÚTIL (m³), CAP. ÚTIL (%), SATURAÇÃO TOTAL (%), COD DESTINO, Mês
     */
    console.log('📊 Displaying Viajante results:', results);
    
    if (!results || results.length === 0) {
        console.warn('No Viajante results to display');
        return;
    }
    
    // Find or create Viajante results section
    let viajanteSection = document.getElementById('viajante-results-section');
    
    if (!viajanteSection) {
        // Create section if it doesn't exist
        const dashboardSection = document.getElementById('dashboard-section');
        if (dashboardSection) {
            viajanteSection = document.createElement('div');
            viajanteSection.id = 'viajante-results-section';
            viajanteSection.className = 'results-section';
            viajanteSection.innerHTML = `
                <h3>🚚 Resultados Viajante - Volume por Rota</h3>
                <div class="table-container">
                    <table class="results-table">
                        <thead>
                            <tr id="viajante-headers-row">
                                <th>COD DESTINO</th>
                                <th>Destino</th>
                                <th>Mês</th>
                                <th>Fornecedores</th>
                                <th>Veículo</th>
                                <th>CARGAS</th>
                                <th>CAP. ÚTIL (m³)</th>
                                <th>CAP. ÚTIL (%)</th>
                                <th>SATURAÇÃO TOTAL (%)</th>
                                <th>Volume Total (m³)</th>
                                <th>Peso Total (kg)</th>
                                <th>Embalagens Total</th>
                                <th>Sugestão</th>
                            </tr>
                        </thead>
                        <tbody id="viajante-results-tbody">
                        </tbody>
                    </table>
                </div>
            `;
            dashboardSection.appendChild(viajanteSection);
        }
    }
    
    // Populate table
    const tbody = document.getElementById('viajante-results-tbody');
    if (!tbody) {
        console.error('Could not find viajante-results-tbody element');
        return;
    }
    
    tbody.innerHTML = '';
    
    results.forEach(row => {
        const tr = document.createElement('tr');
        
        // Apply highlighting based on saturation
        const saturacao = row['SATURAÇÃO TOTAL (%)'] || 0;
        let rowClass = '';
        if (saturacao > 100) {
            rowClass = 'high-saturation';
        } else if (saturacao < 50) {
            rowClass = 'low-saturation';
        }
        
        tr.className = rowClass;
        
        tr.innerHTML = `
            <td class="td-center">${row['COD DESTINO'] || '-'}</td>
            <td class="td-label">${row['DESTINO'] || '-'}</td>
            <td class="td-center">${row['Mês'] || '-'}</td>
            <td class="td-label">${row['FORNECEDORES NA ROTA'] || '-'}</td>
            <td class="td-center">${row['VEÍCULO'] || '-'}</td>
            <td class="td-center"><strong>${row['CARGAS'] || 0}</strong></td>
            <td class="td-center">${row['CAP. ÚTIL (m³)'] ? row['CAP. ÚTIL (m³)'].toFixed(2) : '-'}</td>
            <td class="td-center">${row['CAP. ÚTIL (%)'] ? row['CAP. ÚTIL (%)'].toFixed(2) : '-'}</td>
            <td class="td-center"><strong>${row['SATURAÇÃO TOTAL (%)'] ? row['SATURAÇÃO TOTAL (%)'].toFixed(2) : '-'}%</strong></td>
            <td class="td-center">${row['VOLUME TOTAL (m³)'] ? row['VOLUME TOTAL (m³)'].toFixed(2) : '-'}</td>
            <td class="td-center">${row['PESO TOTAL (kg)'] ? row['PESO TOTAL (kg)'].toFixed(2) : '-'}</td>
            <td class="td-center">${row['EMBALAGENS TOTAL'] || '-'}</td>
            <td class="td-label">${row['SUGESTÃO'] || '-'}</td>
        `;
        
        tbody.appendChild(tr);
    });
    
    // Scroll to Viajante section
    viajanteSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
    
    console.log(`✅ Displayed ${results.length} Viajante results`);
}

function updateWeeklyTrips(qmeResponse, viajanteResponse) {
    /**
     * Calculate and update weekly trips using QME volumes and Viajante capacity data
     * - TO BE: Volume m³ TO BE / CAP. ÚTIL (m³)
     * - AS IS: Either calculated or from TDC activation count
     */
    console.log('📊 Calculating weekly trips...');
    console.log('QME Response:', qmeResponse);
    console.log('Viajante Response:', viajanteResponse);
    
    if (!qmeResponse) {
        console.warn('Cannot calculate trips: missing QME data');
        return;
    }
    
    // Check if trips were already calculated on backend (includes AS IS trips)
    const weeklyTrips = qmeResponse.weekly_trips;
    
    console.log('🔍 Checking backend trip data...');
    console.log('weeklyTrips object:', weeklyTrips);
    
    if (weeklyTrips) {
        console.log('✅ Using trips calculated by backend');
        
        const monthlyTripsTobe = weeklyTrips.monthly_trips_tobe || {};
        const monthlyTripsAsis = weeklyTrips.monthly_trips_asis || {};
        
        console.log('📊 Monthly Trips TO BE:', monthlyTripsTobe);
        console.log('📊 Monthly Trips AS IS:', monthlyTripsAsis);
        
        const monthKeys = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 
                          'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
        
        // Calculate totals
        let totalTripsTobe = 0;
        let totalTripsAsis = 0;
        
        console.log('\n🔢 Processing monthly trips:');
        monthKeys.forEach(monthKey => {
            const tripsTobe = monthlyTripsTobe[monthKey] || 0;
            const tripsAsis = monthlyTripsAsis[monthKey] || 0;
            
            console.log(`  ${monthKey}: AS IS=${tripsAsis}, TO BE=${tripsTobe}`);
            
            if (typeof tripsTobe === 'number') totalTripsTobe += tripsTobe;
            if (typeof tripsAsis === 'number') totalTripsAsis += tripsAsis;
        });
        
        console.log('\n📈 Totals:');
        console.log('  Total TO BE trips:', totalTripsTobe);
        console.log('  Total AS IS trips:', totalTripsAsis);
        
        // Update the breakdown table with trip values
        const breakdownTable = document.getElementById('breakdown-combined-body');
        if (!breakdownTable) {
            console.error('❌ Breakdown table not found!');
            return;
        }
        
        console.log('\n📋 Updating breakdown table...');
        console.log('Table element:', breakdownTable);
        
        // Find the "Qtde de Viagens Semanal" row (it's the second row)
        const rows = breakdownTable.getElementsByTagName('tr');
        console.log(`Table has ${rows.length} rows`);
        
        if (rows.length < 2) {
            console.error('❌ Trips row not found in breakdown table');
            return;
        }
        
        const viagensRow = rows[1]; // Second row
        const cells = viagensRow.getElementsByTagName('td');
        console.log(`Viagens row has ${cells.length} cells`);
        console.log('Row element:', viagensRow);
        
        // Update cells: skip first cell (label), then pairs of AS IS / TO BE for each month
        let cellIndex = 1; // Start after label
        console.log('\n🔄 Updating cell values:');
        
        monthKeys.forEach((monthKey, idx) => {
            // AS IS cell
            if (cellIndex < cells.length) {
                const tripsAsis = monthlyTripsAsis[monthKey] || 0;
                const displayValue = tripsAsis > 0 ? tripsAsis : '-';
                const oldValue = cells[cellIndex].textContent;
                
                console.log(`  Cell ${cellIndex} (${monthKey} AS IS): ${oldValue} → ${displayValue}`);
                cells[cellIndex].textContent = displayValue;
            } else {
                console.warn(`  Cell ${cellIndex} out of bounds (AS IS ${monthKey})`);
            }
            cellIndex++;
            
            // TO BE cell
            if (cellIndex < cells.length) {
                const tripsTobe = monthlyTripsTobe[monthKey] || 0;
                const displayValue = tripsTobe > 0 ? tripsTobe : '-';
                const oldValue = cells[cellIndex].textContent;
                
                console.log(`  Cell ${cellIndex} (${monthKey} TO BE): ${oldValue} → ${displayValue}`);
                cells[cellIndex].textContent = displayValue;
            } else {
                console.warn(`  Cell ${cellIndex} out of bounds (TO BE ${monthKey})`);
            }
            cellIndex++;
        });
        
        // Update total cells (last two cells)
        console.log('\n📊 Updating total cells:');
        console.log(`  Total cells available: ${cells.length}`);
        
        if (cells.length >= 2) {
            // Second to last: AS IS total
            const totalAsisDisplay = totalTripsAsis > 0 ? totalTripsAsis : '-';
            const asIsIndex = cells.length - 2;
            const oldAsIsValue = cells[asIsIndex].textContent;
            
            console.log(`  Cell ${asIsIndex} (Total AS IS): ${oldAsIsValue} → ${totalAsisDisplay}`);
            cells[asIsIndex].textContent = totalAsisDisplay;
            
            // Last: TO BE total
            const totalTobeDisplay = totalTripsTobe > 0 ? totalTripsTobe : '-';
            const toBeIndex = cells.length - 1;
            const oldToBeValue = cells[toBeIndex].textContent;
            
            console.log(`  Cell ${toBeIndex} (Total TO BE): ${oldToBeValue} → ${totalTobeDisplay}`);
            cells[toBeIndex].textContent = totalTobeDisplay;
        } else {
            console.error('❌ Not enough cells for totals!');
        }
        
        console.log('\n✅ Weekly trips (AS IS & TO BE) updated in breakdown table');
        
        // ── Freight cost rows ──────────────────────────────────────────
        const freight = weeklyTrips.freight;
        if (freight) {
            const fmt = (v) => `R$ ${(v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
            const freightSuccess = freight.status === 'success';
            const monthlyFreightAsis = freightSuccess ? (freight.monthly_freight_asis || {}) : {};
            const monthlyFreightTobe = freightSuccess ? (freight.monthly_freight_tobe || {}) : {};
            // Pedágio is always computed by the backend (trips × pedagio per trip)
            const pedagioTobe = freight.monthly_pedagio_tobe || {};

            // Row 2 — Custo Mensal de Frete (only when tarifa lookup succeeded)
            if (freightSuccess && rows.length >= 3) {
                const cCells = rows[2].getElementsByTagName('td');
                let totalAsis = 0;
                let totalTobe = 0;
                let ci = 1;
                monthKeys.forEach(month => {
                    const asis = monthlyFreightAsis[month] || 0;
                    const tobe = monthlyFreightTobe[month] || 0;
                    totalAsis += asis;
                    totalTobe += tobe;
                    if (ci < cCells.length) cCells[ci].textContent = fmt(asis);
                    if (ci + 1 < cCells.length) cCells[ci + 1].textContent = fmt(tobe);
                    ci += 2;
                });
                if (cCells.length >= 2) {
                    cCells[cCells.length - 2].textContent = fmt(totalAsis);
                    cCells[cCells.length - 1].textContent = fmt(totalTobe);
                }
            }

            // Row 3 — Pedágio: AS IS and TO BE (trips × pedagio per trip)
            const pedagioAsis = freight.monthly_pedagio_asis || {};
            if (rows.length >= 4) {
                const pCells = rows[3].getElementsByTagName('td');
                let totalPedagioAsis = 0;
                let totalPedagioTobe = 0;
                let pi = 1;
                monthKeys.forEach(month => {
                    const asis = pedagioAsis[month] || 0;
                    const tobe = pedagioTobe[month] || 0;
                    totalPedagioAsis += asis;
                    totalPedagioTobe += tobe;
                    if (pi < pCells.length) pCells[pi].textContent = fmt(asis);
                    if (pi + 1 < pCells.length) pCells[pi + 1].textContent = fmt(tobe);
                    pi += 2;
                });
                if (pCells.length >= 2) {
                    pCells[pCells.length - 2].textContent = fmt(totalPedagioAsis);
                    pCells[pCells.length - 1].textContent = fmt(totalPedagioTobe);
                }
            }

            // Row 4 — Economia Mensal de Frete (only when tarifa lookup succeeded)
            if (freightSuccess && rows.length >= 5) {
                const ctCells = rows[4].getElementsByTagName('td');
                let totalSaving = 0;
                monthKeys.forEach((month, idx) => {
                    const saving = (monthlyFreightAsis[month] || 0) - (monthlyFreightTobe[month] || 0);
                    totalSaving += saving;
                    if (idx + 1 < ctCells.length) ctCells[idx + 1].textContent = fmt(saving);
                });
                if (ctCells.length >= 14) ctCells[13].textContent = fmt(totalSaving);
            }

            // Row 5 — TOTAL SEMANAL = Frete TO BE + Pedágio TO BE
            if (rows.length >= 6) {
                const tsCells = rows[5].getElementsByTagName('td');
                let totalSemanal = 0;
                monthKeys.forEach((month, idx) => {
                    const val = (monthlyFreightTobe[month] || 0) + (pedagioTobe[month] || 0);
                    totalSemanal += val;
                    if (idx + 1 < tsCells.length) tsCells[idx + 1].textContent = fmt(val);
                });
                if (tsCells.length >= 14) tsCells[13].textContent = fmt(totalSemanal);
            }

            // Row 6 — TOTAL MENSAL = TOTAL SEMANAL × 4
            if (rows.length >= 7) {
                const tmCells = rows[6].getElementsByTagName('td');
                let totalMensal = 0;
                monthKeys.forEach((month, idx) => {
                    const val = ((monthlyFreightTobe[month] || 0) + (pedagioTobe[month] || 0)) * 4;
                    totalMensal += val;
                    if (idx + 1 < tmCells.length) tmCells[idx + 1].textContent = fmt(val);
                });
                if (tmCells.length >= 14) tmCells[13].textContent = fmt(totalMensal);
            }

            // Savings summary rows (only when freight succeeded)
            if (freightSuccess) {
                const monthlyBodyEl = document.getElementById('dashboard-monthly-body');
                if (monthlyBodyEl) {
                    const savingRowEl = monthlyBodyEl.querySelector('tr:last-child');
                    if (savingRowEl) {
                        const sCells = savingRowEl.getElementsByTagName('td');
                        let totalSaving = 0;
                        let si = 1;
                        monthKeys.forEach(month => {
                            const saving = (monthlyFreightAsis[month] || 0) - (monthlyFreightTobe[month] || 0);
                            totalSaving += saving;
                            if (si < sCells.length) sCells[si].textContent = fmt(saving);
                            si++;
                        });
                        if (sCells.length > 0) sCells[sCells.length - 1].textContent = fmt(totalSaving);
                    }
                }

                const savingsCard = document.getElementById('dashboard-summary-savings');
                if (savingsCard) {
                    const annualSaving = monthKeys.reduce((sum, month) =>
                        sum + (monthlyFreightAsis[month] || 0) - (monthlyFreightTobe[month] || 0), 0);
                    savingsCard.innerText = annualSaving.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
                }
            }

            console.log('✅ Freight costs updated in breakdown table');
        }
        
        return;
    }
    
    // Fallback: if backend didn't calculate, do it here (old logic, TO BE only)
    console.log('⚠️ Backend trips not available, calculating TO BE only...');
    
    const monthlyM3Tobe = qmeResponse.summary?.monthly_m3_tobe;
    if (!monthlyM3Tobe) {
        console.warn('No monthly TO BE data available');
        return;
    }
    
    const viajanteResults = viajanteResponse.results || [];
    if (viajanteResults.length === 0) {
        console.warn('No Viajante results available');
        return;
    }
    
    const monthKeys = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 
                       'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
    
    const englishToPortuguese = {
        'Jan': 'Jan', 'Feb': 'Fev', 'Mar': 'Mar', 'Apr': 'Abr',
        'May': 'Mai', 'Jun': 'Jun', 'Jul': 'Jul', 'Aug': 'Ago',
        'Sep': 'Set', 'Oct': 'Out', 'Nov': 'Nov', 'Dec': 'Dez'
    };
    
    const monthCapacity = {};
    viajanteResults.forEach(row => {
        const mesEnglish = row['Mês'];
        const mesPortuguese = englishToPortuguese[mesEnglish];
        const capUtil = row['CAP. ÚTIL (m³)'];
        
        if (mesPortuguese && capUtil && !monthCapacity[mesPortuguese]) {
            monthCapacity[mesPortuguese] = capUtil;
        }
    });
    
    const monthlyTripsTobe = {};
    let totalTrips = 0;
    
    monthKeys.forEach(monthKey => {
        const volumeTobe = monthlyM3Tobe[monthKey] || 0;
        const capacity = monthCapacity[monthKey] || 0;
        
        let trips = 0;
        if (capacity > 0 && volumeTobe > 0) {
            trips = Math.round(volumeTobe / capacity);
            totalTrips += trips;
        }
        
        monthlyTripsTobe[monthKey] = trips;
    });
    
    const breakdownTable = document.getElementById('breakdown-combined-body');
    if (!breakdownTable) return;
    
    const rows = breakdownTable.getElementsByTagName('tr');
    if (rows.length < 2) return;
    
    const viagensRow = rows[1];
    const cells = viagensRow.getElementsByTagName('td');
    
    let cellIndex = 1;
    monthKeys.forEach(monthKey => {
        cellIndex++; // Skip AS IS
        
        if (cellIndex < cells.length) {
            const trips = monthlyTripsTobe[monthKey] || 0;
            cells[cellIndex].textContent = trips > 0 ? trips : '-';
        }
        cellIndex++;
    });
    
    if (cells.length >= 2) {
        cells[cells.length - 1].textContent = totalTrips > 0 ? totalTrips : '-';
    }
    
    console.log('✅ Weekly trips (TO BE only) updated');
}

function exportResults() {
    console.log('Exporting breakdown results to Excel...');
    
    if (!window.pywebview || !window.pywebview.api) {
        alert('API ainda não está pronta. Por favor, aguarde um momento.');
        return;
    }
    
    window.pywebview.api.export_results().then(response => {
        console.log('Export response:', response);
        
        if (response.status === 'success') {
            showToast('✅ ' + response.message, 'success');
        } else if (response.status === 'error') {
            showToast('❌ ' + response.message, 'error');
        }
    }).catch(error => {
        console.error('Error exporting breakdown:', error);
        showToast('❌ Erro ao exportar: ' + error, 'error');
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
    
    // Adiciona listener para o campo destino (IMS Destino)
    const destinoInput = document.getElementById('destino');
    if (destinoInput) {
        destinoInput.addEventListener('input', handleDestinoInput);
        console.log('Destino auto-fetch listener attached');
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
    
    // Adiciona listener para o campo destino
    const destinoInput = document.getElementById('destino');
    if (destinoInput && !destinoInput.hasAttribute('data-destino-listener-attached')) {
        destinoInput.addEventListener('input', handleDestinoInput);
        destinoInput.setAttribute('data-destino-listener-attached', 'true');
        console.log('Destino auto-fetch listener attached (DOMContentLoaded)');
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
// PN/MDR Filter Functionality
function filterPNTable() {
    const filterInput = document.getElementById('pn-filter');
    const filterValue = filterInput.value.toLowerCase().trim();
    const table = document.getElementById('dashboard-results-table');
    const tbody = table.querySelector('tbody');
    const rows = tbody.querySelectorAll('tr');
    
    let visibleCount = 0;
    
    rows.forEach(row => {
        // Skip totals row if it exists
        if (row.classList.contains('totals-row')) {
            return;
        }
        
        // Get PN (2nd cell) and MDR AS IS (16th cell) and MDR TO BE (20th cell)
        const cells = row.querySelectorAll('td');
        if (cells.length < 2) return;
        
        const pn = cells[1].textContent.toLowerCase();
        const mdrAsIs = cells.length > 15 ? cells[15].textContent.toLowerCase() : '';
        const mdrToBe = cells.length > 19 ? cells[19].textContent.toLowerCase() : '';
        
        // Show row if filter is empty or matches PN or MDR
        if (!filterValue || 
            pn.includes(filterValue) || 
            mdrAsIs.includes(filterValue) || 
            mdrToBe.includes(filterValue)) {
            row.style.display = '';
            visibleCount++;
        } else {
            row.style.display = 'none';
        }
    });
    
    // Update row numbers for visible rows
    let rowNum = 1;
    rows.forEach(row => {
        if (row.style.display !== 'none' && !row.classList.contains('totals-row')) {
            const firstCell = row.querySelector('td:first-child');
            if (firstCell) {
                firstCell.textContent = rowNum++;
            }
        }
    });
}

function clearPNFilter() {
    const filterInput = document.getElementById('pn-filter');
    filterInput.value = '';
    filterPNTable();
}

// Function to export PN table to Excel
function exportPNTable() {
    console.log('Exporting PN table to Excel...');
    
    if (!window.pywebview || !window.pywebview.api) {
        alert('API ainda não está pronta. Por favor, aguarde um momento.');
        return;
    }
    
    window.pywebview.api.export_pn_table().then(result => {
        console.log('Export response:', result);
        
        if (result.status === 'success') {
            showToast('✅ ' + result.message, 'success');
        } else if (result.status === 'error') {
            showToast('❌ ' + result.message, 'error');
        }
    }).catch(error => {
        console.error('Error exporting PN table:', error);
        showToast('❌ Erro ao exportar: ' + error, 'error');
    });
}


// Add event listener for real-time filtering
document.addEventListener('DOMContentLoaded', () => {
    const filterInput = document.getElementById('pn-filter');
    if (filterInput) {
        filterInput.addEventListener('input', filterPNTable);
    }
});

// ── Zoom support (Ctrl+scroll and Ctrl++/-/0) ──────────────────────────
(function initZoom() {
    let currentZoom = 0.9;
    const STEP = 0.1;
    const MIN  = 0.5;
    const MAX  = 2.5;

    function applyZoom(z) {
        currentZoom = Math.min(MAX, Math.max(MIN, z));
        document.documentElement.style.zoom = currentZoom;
    }

    // Apply default zoom immediately
    applyZoom(currentZoom);

    // Ctrl + mouse-wheel
    window.addEventListener('wheel', function(e) {
        if (!e.ctrlKey) return;
        e.preventDefault();
        applyZoom(currentZoom + (e.deltaY < 0 ? STEP : -STEP));
    }, { passive: false });

    // Ctrl + / Ctrl - / Ctrl 0
    window.addEventListener('keydown', function(e) {
        if (!e.ctrlKey) return;
        if (e.key === '+' || e.key === '=' ) { e.preventDefault(); applyZoom(currentZoom + STEP); }
        else if (e.key === '-')             { e.preventDefault(); applyZoom(currentZoom - STEP); }
        else if (e.key === '0')             { e.preventDefault(); applyZoom(0.9); }
    });
})();
