// VARIÁVEIS GLOBAIS
let map;
let colaboradoresData = [];
let bairrosData = null;
let unificacoes = []; // Array para armazenar os grupos de unificação
let routes = {};
let bairroLayers = {};
let generatedRoutesData = {};
let colaboradoresPorRota = {}; // Armazena os colaboradores alocados para cada rota
let baseMaps = {};
let overlayMaps = {};
let tipoRota = 'ENTRADA'; // Padrão: ENTRADA
let editMode = false;
let draggableMarkers = [];
let todosBairros = []; // Armazena a lista completa de bairros do GeoJSON
let selectedBairros = new Set(); // Conjunto para armazenar bairros selecionados
const colorPalette = [
    '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', 
    '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf',
    '#aec7e8', '#ffbb78', '#98df8a', '#ff9896', '#c5b0d5',
    '#c49c94', '#f7b6d2', '#dbdb8d', '#9edae5', '#8c6d31'
];

// Inicialização do mapa
function initMap() {
    map = L.map('map').setView([-3.124488, -59.963292], 12);

    // Configurar camadas base
    baseMaps = {
        "Mapa": L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            attribution: '&copy; OpenStreetMap contributors'
        }),
        "Satélite": L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
            attribution: 'Tiles &copy; Esri &mdash; Source: Esri, i-cubed, USDA, USGS, AEX, GeoEye, Getmapping, Aerogrid, IGN, IGP, UPR-EGP, and the GIS User Community'
        })
    };

    // Adicionar camada padrão
    baseMaps["Mapa"].addTo(map);

    // Inicializar camadas de overlay
    overlayMaps = {
        "Bairros": L.layerGroup(),
        "Rotas": L.layerGroup(),
        "Marcadores": L.layerGroup(),
        "Pontos Referência": L.layerGroup()
    };

    // Adicionar controle de camadas
    const layerControl = L.control.layers(baseMaps, overlayMaps).addTo(map);

    // Adicionar camadas de overlay ao mapa por padrão
    overlayMaps["Bairros"].addTo(map);
    overlayMaps["Rotas"].addTo(map);
    overlayMaps["Marcadores"].addTo(map);
    overlayMaps["Pontos Referência"].addTo(map);

    // Adicionar marcadores de ponto inicial e destino
    adicionarPontosReferencia();

    // Forçar texto preto nos controles de camadas
    setTimeout(() => {
        const labels = document.querySelectorAll('.leaflet-control-layers label');
        const spans = document.querySelectorAll('.leaflet-control-layers span');
        
        labels.forEach(label => {
            label.style.color = '#000000';
        });
        
        spans.forEach(span => {
            span.style.color = '#000000';
        });
    }, 100);
}

// Alternar entre tipos de rota
document.getElementById('btn-entrada').addEventListener('click', function() {
    tipoRota = 'ENTRADA';
    document.getElementById('btn-entrada').classList.add('active');
    document.getElementById('btn-saida').classList.remove('active');
    document.getElementById('ponto-destino').classList.add('active');
    document.getElementById('ponto-inicial').classList.remove('active');
    adicionarPontosReferencia();
});

document.getElementById('btn-saida').addEventListener('click', function() {
    tipoRota = 'SAÍDA';
    document.getElementById('btn-saida').classList.add('active');
    document.getElementById('btn-entrada').classList.remove('active');
    document.getElementById('ponto-inicial').classList.add('active');
    document.getElementById('ponto-destino').classList.remove('active');
    adicionarPontosReferencia();
});

// Alternar modo de edição
document.getElementById('toggle-edit').addEventListener('click', function() {
    editMode = !editMode;
    const editMessage = document.getElementById('edit-message');
    const toggleButton = document.getElementById('toggle-edit');
    
    if (editMode) {
        toggleButton.innerHTML = '<i class="fas fa-check"></i> Finalizar Edição';
        toggleButton.style.background = 'linear-gradient(to right, #2ecc71, #27ae60)';
        editMessage.style.display = 'block';
        
        // Ativar modo de arrastar para todos os marcadores
        draggableMarkers.forEach(marker => {
            marker.dragging.enable();
        });
    } else {
        toggleButton.innerHTML = '<i class="fas fa-edit"></i> Mover Pontos';
        toggleButton.style.background = 'linear-gradient(to right, #1abc9c, #16a085)';
        editMessage.style.display = 'none';
        
        // Desativar modo de arrastar para todos os marcadores
        draggableMarkers.forEach(marker => {
            marker.dragging.disable();
        });
        
        // Recalcular rotas se necessário
        if (draggableMarkers.length > 0) {
            updateStatus('Pontos movidos. Clique em "Processar Rotas" para atualizar.', 'success');
        }
    }
});

// Adicionar pontos de referência (inicial e final)
function adicionarPontosReferencia() {
    // Limpar pontos anteriores
    overlayMaps["Pontos Referência"].clearLayers();

    if (tipoRota === 'ENTRADA') {
        // Para rotas de ENTRADA: Destino Final
        const destinoLat = parseFloat(document.getElementById('destino-lat').value) || -3.124488;
        const destinoLng = parseFloat(document.getElementById('destino-lng').value) || -59.963292;

        const destinoMarker = L.marker([destinoLat, destinoLng], {
            icon: L.icon({
                iconUrl: 'https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-2x-green.png',
                shadowUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/0.7.7/images/marker-shadow.png',
                iconSize: [25, 41],
                iconAnchor: [12, 41],
                popupAnchor: [1, -34],
                shadowSize: [41, 41]
            }),
            draggable: true
        }).addTo(overlayMaps["Pontos Referência"]);

        destinoMarker.bindPopup('<b>Destino Final</b>').openPopup();

        // Atualizar coordenadas quando o marcador for movido
        destinoMarker.on('dragend', function() {
            const latLng = destinoMarker.getLatLng();
            document.getElementById('destino-lat').value = latLng.lat.toFixed(6);
            document.getElementById('destino-lng').value = latLng.lng.toFixed(6);
        });

        draggableMarkers.push(destinoMarker);

    } else {
        // Para rotas de SAÍDA: Ponto Inicial
        const inicialLat = parseFloat(document.getElementById('inicial-lat').value) || -3.124488;
        const inicialLng = parseFloat(document.getElementById('inicial-lng').value) || -59.963292;

        const inicialMarker = L.marker([inicialLat, inicialLng], {
            icon: L.icon({
                iconUrl: 'https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-2x-green.png',
                shadowUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/0.7.7/images/marker-shadow.png',
                iconSize: [25, 41],
                iconAnchor: [12, 41],
                popupAnchor: [1, -34],
                shadowSize: [41, 41]
            }),
            draggable: true
        }).addTo(overlayMaps["Pontos Referência"]);

        inicialMarker.bindPopup('<b>Ponto Inicial</b>').openPopup();

        // Atualizar coordenadas quando o marcador for movido
        inicialMarker.on('dragend', function() {
            const latLng = inicialMarker.getLatLng();
            document.getElementById('inicial-lat').value = latLng.lat.toFixed(6);
            document.getElementById('inicial-lng').value = latLng.lng.toFixed(6);
        });

        draggableMarkers.push(inicialMarker);
    }

    // Desativar arrasto inicialmente
    draggableMarkers.forEach(marker => {
        marker.dragging.disable();
    });
}

// Processar arquivo Excel
document.getElementById('excel-file').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet);

            // Verificar colunas obrigatórias
            const requiredColumns = ['LATITUDE', 'LONGITUDE', 'COLABORADOR', 'MATRICULA'];
            const hasAllColumns = requiredColumns.every(col => 
                jsonData.length > 0 && Object.keys(jsonData[0]).includes(col)
            );

            if (!hasAllColumns) {
                showError('Erro: O arquivo Excel não possui todas as colunas necessárias (LATITUDE, LONGITUDE, COLABORADOR, MATRICULA).');
                return;
            }

            colaboradoresData = jsonData.filter(item => 
                item.LATITUDE && item.LONGITUDE && item.COLABORADOR && item.MATRICULA
            );

            updateStatus(`Dados carregados: ${colaboradoresData.length} colaboradores válidos encontrados.`, 'success');
            updateStats();
        } catch (error) {
            console.error('Erro ao processar arquivo Excel:', error);
            showError('Erro ao processar o arquivo Excel. Verifique o formato.');
        }
    };
    reader.readAsArrayBuffer(file);
});

// Processar arquivo GeoJSON
document.getElementById('geojson-file').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            bairrosData = JSON.parse(e.target.result);

            // Limpar camada de bairros existente e unificações anteriores
            overlayMaps["Bairros"].clearLayers();
            unificacoes = [];
            selectedBairros.clear();
            renderUnificacoesList();
            
            // Adicionar bairros ao mapa e coletar nomes
            todosBairros = [];
            L.geoJSON(bairrosData, {
                style: {
                    color: '#e74c3c',
                    weight: 2,
                    fillColor: '#e74c3c',
                    fillOpacity: 0.1
                },
                onEachFeature: function(feature, layer) {
                    if (feature.properties && feature.properties.Name) {
                        layer.bindPopup(`<b>${feature.properties.Name}</b>`);
                        todosBairros.push(feature.properties.Name);
                    }
                }
            }).addTo(overlayMaps["Bairros"]);

            todosBairros.sort();
            renderBairrosCheckboxes(todosBairros);
            document.querySelector('.unificar-section').style.display = 'block';

            updateStatus(`GeoJSON carregado: ${bairrosData.features.length} bairros encontrados.`, 'success');
        } catch (error) {
            console.error('Erro ao processar arquivo GeoJSON:', error);
            showError('Erro ao processar o arquivo GeoJSON. Verifique o formato.');
        }
    };
    reader.readAsText(file);
});

// Renderizar a lista de bairros como checkboxes
function renderBairrosCheckboxes(bairros) {
    const container = document.getElementById('bairro-select-container');
    container.innerHTML = '';

    bairros.forEach(bairro => {
        const item = document.createElement('div');
        item.className = 'bairro-list-item';
        
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.id = `bairro-${bairro}`;
        checkbox.value = bairro;
        checkbox.checked = selectedBairros.has(bairro);

        const label = document.createElement('label');
        label.htmlFor = `bairro-${bairro}`;
        label.textContent = bairro;

        item.appendChild(checkbox);
        item.appendChild(label);
        
        // Adicionar evento para manter a seleção no conjunto
        checkbox.addEventListener('change', () => {
            if (checkbox.checked) {
                selectedBairros.add(bairro);
            } else {
                selectedBairros.delete(bairro);
            }
        });

        container.appendChild(item);
    });
}

// Filtro de pesquisa de bairros
document.getElementById('bairro-search').addEventListener('input', function(e) {
    const searchTerm = e.target.value.toLowerCase();
    const filteredBairros = todosBairros.filter(bairro => 
        bairro.toLowerCase().includes(searchTerm)
    );
    renderBairrosCheckboxes(filteredBairros);
});

// Adicionar um grupo de unificação
document.getElementById('add-unificacao-btn').addEventListener('click', function() {
    const nomeUnificacao = document.getElementById('unificacao-nome').value.trim();
    const selectedBairrosArray = Array.from(selectedBairros);

    if (selectedBairrosArray.length === 0) {
        showError('Selecione pelo menos um bairro para unificar.');
        return;
    }
    if (nomeUnificacao === '') {
        showError('Dê um nome para a unificação.');
        return;
    }

    // Adicionar a nova unificação ao array
    unificacoes.push({
        nome: nomeUnificacao,
        bairros: selectedBairrosArray
    });

    // Limpar os campos após adicionar
    document.getElementById('unificacao-nome').value = '';
    document.getElementById('bairro-search').value = '';
    selectedBairros.clear();
    renderBairrosCheckboxes(todosBairros);

    renderUnificacoesList();
    updateStatus(`Unificação "${nomeUnificacao}" adicionada com sucesso.`, 'success');
});

// Renderizar a lista de unificações na interface
function renderUnificacoesList() {
    const list = document.getElementById('unificacoes-list');
    list.innerHTML = '';

    unificacoes.forEach((unificacao, index) => {
        const item = document.createElement('div');
        item.className = 'unificacao-item';
        item.innerHTML = `
            <span><strong>${unificacao.nome}</strong>: ${unificacao.bairros.join(', ')}</span>
            <button class="remove-btn" data-index="${index}"><i class="fas fa-trash"></i></button>
        `;
        list.appendChild(item);
    });

    // Adicionar eventos de clique para o botão de remover
    list.querySelectorAll('.remove-btn').forEach(button => {
        button.addEventListener('click', function() {
            const index = this.getAttribute('data-index');
            unificacoes.splice(index, 1);
            renderUnificacoesList();
        });
    });
}

// Processar rotas
document.getElementById('process-btn').addEventListener('click', async function() {
    const apiKey = document.getElementById('api-key').value;
    if (!apiKey) {
        showError('Erro: É necessário informar uma chave API válida.');
        return;
    }

    if (colaboradoresData.length === 0) {
        showError('Erro: Nenhum dado de colaborador carregado.');
        return;
    }

    if (!bairrosData) {
        showError('Erro: Nenhum arquivo GeoJSON de bairros carregado.');
        return;
    }

    // Desativar modo de edição se estiver ativo
    if (editMode) {
        editMode = false;
        document.getElementById('toggle-edit').innerHTML = '<i class="fas fa-edit"></i> Mover Pontos';
        document.getElementById('toggle-edit').style.background = 'linear-gradient(to right, #1abc9c, #16a085)';
        document.getElementById('edit-message').style.display = 'none';

        draggableMarkers.forEach(marker => {
            marker.dragging.disable();
        });
    }

    updateStatus('Processando rotas...', 'loading');

    try {
        // Limpar camadas e marcadores anteriores
        generatedRoutesData = {};
        colaboradoresPorRota = {};
        document.getElementById('download-kml-btn').disabled = true;
        document.getElementById('download-excel-btn').disabled = true;
        overlayMaps["Rotas"].clearLayers();
        overlayMaps["Marcadores"].clearLayers();
        draggableMarkers = []; // Limpa a lista de marcadores para evitar duplicatas

        // Agrupar colaboradores por bairro, considerando as unificações
        colaboradoresPorRota = agruparColaboradoresPorRota(colaboradoresData, bairrosData, unificacoes);

        // Limpar lista de bairros
        const bairrosList = document.getElementById('bairros-list');
        bairrosList.innerHTML = '';

        // Iniciar barra de progresso
        const progressBar = document.getElementById('progress-bar');
        const totalRotas = Object.keys(colaboradoresPorRota).length;
        let processedRotas = 0;

        // Processar cada rota
        let rotaIndex = 0;
        for (const [rotaName, colaboradores] of Object.entries(colaboradoresPorRota)) {
            if (rotaIndex >= colorPalette.length) rotaIndex = 0;

            const color = colorPalette[rotaIndex];
            const coords = colaboradores.map(c => [c.LATITUDE, c.LONGITUDE]);

            // Adicionar marcadores para cada colaborador (arrastáveis)
            colaboradores.forEach(colab => {
                const marker = L.marker([colab.LATITUDE, colab.LONGITUDE], {
                    icon: L.icon({
                        iconUrl: 'https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-2x-blue.png',
                        shadowUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/0.7.7/images/marker-shadow.png',
                        iconSize: [25, 41],
                        iconAnchor: [12, 41],
                        popupAnchor: [1, -34],
                        shadowSize: [41, 41]
                    }),
                    draggable: true
                });

                marker.bindPopup(`
                    <b>${colab.COLABORADOR}</b><br>
                    Matrícula: ${colab.MATRICULA}<br>
                    Bairro: ${colab.BAIRRO_ORIGEM || 'N/A'}<br>
                    Rota: ${rotaName}
                `);

                marker.addTo(overlayMaps["Marcadores"]);

                // Armazenar referência para poder habilitar/desabilitar arrasto
                draggableMarkers.push(marker);

                // Desativar arrasto inicialmente
                marker.dragging.disable();

                // Atualizar dados quando o marcador for movido
                marker.on('dragend', function() {
                    const latLng = marker.getLatLng();
                    // Atualizar os dados do colaborador
                    const colaboradorIndex = colaboradoresData.findIndex(c => 
                        c.MATRICULA === colab.MATRICULA && c.COLABORADOR === colab.COLABORADOR
                    );

                    if (colaboradorIndex !== -1) {
                        colaboradoresData[colaboradorIndex].LATITUDE = latLng.lat;
                        colaboradoresData[colaboradorIndex].LONGITUDE = latLng.lng;
                    }
                });
            });

            // Calcular rota usando a API OpenRouteService
            if (coords.length > 0) {
                try {
                    // Definir pontos de origem e destino baseado no tipo de rota
                    let waypoints;
                    if (tipoRota === 'ENTRADA') {
                        // Para ENTRADA: Colaboradores -> Destino Final
                        const destinoLat = parseFloat(document.getElementById('destino-lat').value) || -3.124488;
                        const destinoLng = parseFloat(document.getElementById('destino-lng').value) || -59.963292;
                        waypoints = [...coords, [destinoLat, destinoLng]];
                    } else {
                        // Para SAÍDA: Ponto Inicial -> Colaboradores
                        const inicialLat = parseFloat(document.getElementById('inicial-lat').value) || -3.124488;
                        const inicialLng = parseFloat(document.getElementById('inicial-lng').value) || -59.963292;
                        waypoints = [[inicialLat, inicialLng], ...coords];
                    }

                    const routeCoords = await calculateRoute(apiKey, waypoints);
                    if (routeCoords && routeCoords.length > 0) {
                        generatedRoutesData[rotaName] = routeCoords;
                        const polyline = L.polyline(routeCoords, {
                            color: tipoRota === 'ENTRADA' ? '#ff7f0e' : '#2ca02c', 
                            weight: 5,
                            opacity: 0.7
                        });

                        // Adicionar popup à rota com o nome do bairro
                        polyline.bindPopup(`<b>Rota ${tipoRota.toLowerCase()}: ${rotaName}</b><br>${colaboradores.length} colaboradores`);
                        polyline.addTo(overlayMaps["Rotas"]);

                        // Armazenar referência à rota
                        if (!bairroLayers[rotaName]) {
                            bairroLayers[rotaName] = [];
                        }
                        bairroLayers[rotaName].push(polyline);
                    }
                } catch (error) {
                    console.error(`Erro ao calcular rota para ${rotaName}:`, error);
                    // Desenhar linha reta em caso de erro na API
                    let waypoints;
                    if (tipoRota === 'ENTRADA') {
                        const destinoLat = parseFloat(document.getElementById('destino-lat').value) || -3.124488;
                        const destinoLng = parseFloat(document.getElementById('destino-lng').value) || -59.963292;
                        waypoints = [...coords, [destinoLat, destinoLng]];
                    } else {
                        const inicialLat = parseFloat(document.getElementById('inicial-lat').value) || -3.124488;
                        const inicialLng = parseFloat(document.getElementById('inicial-lng').value) || -59.963292;
                        waypoints = [[inicialLat, inicialLng], ...coords];
                    }
                    
                    const polyline = L.polyline(waypoints, {
                        color: tipoRota === 'ENTRADA' ? '#ff7f0e' : '#2ca02c',
                        weight: 3,
                        dashArray: '5, 10',
                        opacity: 0.7
                    });
                    polyline.bindPopup(`<b>Rota alternativa ${tipoRota.toLowerCase()}: ${rotaName}</b><br>API indisponível - rota em linha reta`);
                    polyline.addTo(overlayMaps["Rotas"]);
                }
            }
            
            // Adicionar à lista de bairros
            const rotaItem = document.createElement('div');
            rotaItem.className = 'bairro-item';
            rotaItem.innerHTML = `
                <div class="bairro-header">
                    <strong>${rotaName}</strong>
                    <span>${colaboradores.length} colaboradores</span>
                </div>
                <div class="collaborator-list">
                    ${colaboradores.map(c => 
                        `<div class="collaborator-item">
                            <span>${c.MATRICULA}</span>
                            <span>${c.COLABORADOR}</span>
                            <span>Bairro: ${c.BAIRRO_ORIGEM || 'N/A'}</span>
                        </div>`
                    ).join('')}
                </div>
            `;
            
            // Adicionar evento de clique para destacar o bairro
            rotaItem.addEventListener('click', function() {
                highlightBairro(rotaName);
                this.classList.toggle('active');
            });
            
            bairrosList.appendChild(rotaItem);
            
            rotaIndex++;
            processedRotas++;
            
            // Atualizar barra de progresso
            progressBar.style.width = `${(processedRotas / totalRotas) * 100}%`;
        }
        
        // Ajustar a visualização do mapa para mostrar todos os bairros
        if (Object.keys(colaboradoresPorRota).length > 0) {
            const bounds = L.latLngBounds();
            overlayMaps["Marcadores"].eachLayer(layer => {
                bounds.extend(layer.getLatLng());
            });
            overlayMaps["Pontos Referência"].eachLayer(layer => {
                bounds.extend(layer.getLatLng());
            });
            map.fitBounds(bounds, { padding: [50, 50] });
        }
        
        updateStatus(`Processamento concluído: ${totalRotas} rotas processadas.`, 'success');
        updateStats();
        if (Object.keys(generatedRoutesData).length > 0) {
            document.getElementById('download-kml-btn').disabled = false;
        }
        
        if (Object.keys(colaboradoresPorRota).length > 0) {
            document.getElementById('download-excel-btn').disabled = false;
        }
        
    } catch (error) {
        console.error('Erro ao processar rotas:', error);
        showError('Erro ao processar as rotas. Verifique os dados e a conexão.');
    }
});

// Agrupar colaboradores por rota, considerando as unificações
function agruparColaboradoresPorRota(colaboradores, bairrosGeoJSON, unificacoes) {
    const colaboradoresPorRota = {};
    const bairrosUnificados = new Set(unificacoes.flatMap(grupo => grupo.bairros));

    // Mapear bairros unificados para seus nomes de rota
    const mapaUnificacoes = {};
    unificacoes.forEach(grupo => {
        grupo.bairros.forEach(bairro => {
            mapaUnificacoes[bairro] = grupo.nome;
        });
    });

    // Para cada colaborador, encontrar o bairro e atribuí-lo a uma rota
    colaboradores.forEach(colab => {
        const ponto = L.latLng(colab.LATITUDE, colab.LONGITUDE);
        let bairroEncontrado = "Bairro Não Identificado";
        
        // Verifique se o script leaflet-pip foi carregado antes de usá-lo
        if (typeof leafletPip !== 'undefined') {
            bairrosGeoJSON.features.forEach(feature => {
                const polygon = L.geoJSON(feature);
                if (leafletPip.pointInLayer(ponto, polygon).length > 0) {
                    bairroEncontrado = feature.properties.Name || "Bairro Desconhecido";
                }
            });
        }
        
        // Atribuir o colaborador à rota unificada se o bairro dele estiver no grupo
        let rotaName = mapaUnificacoes[bairroEncontrado] || bairroEncontrado;
        
        if (!colaboradoresPorRota[rotaName]) {
            colaboradoresPorRota[rotaName] = [];
        }
        
        // Armazenar o bairro de origem
        colab.BAIRRO_ORIGEM = bairroEncontrado;
        colaboradoresPorRota[rotaName].push(colab);
    });
    
    return colaboradoresPorRota;
}

// Calcular rota usando a API OpenRouteService
async function calculateRoute(apiKey, coordinates) {
    // Formatar coordenadas para a API [long, lat]
    const formattedCoords = coordinates.map(coord => [coord[1], coord[0]]);
    
    // Construir o corpo da requisição
    const body = {
        coordinates: formattedCoords,
        instructions: false,
        preference: 'recommended'
    };
    
    try {
        const response = await fetch('https://api.openrouteservice.org/v2/directions/driving-car/geojson', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': apiKey
            },
            body: JSON.stringify(body)
        });
        
        if (!response.ok) {
            throw new Error(`Erro na API: ${response.status} ${response.statusText}`);
        }
        
        const data = await response.json();
        
        // Extrair coordenadas da rota
        if (data.features && data.features.length > 0) {
            const routeCoords = data.features[0].geometry.coordinates;
            // Converter para formato [lat, lng] para o Leaflet
            return routeCoords.map(coord => [coord[1], coord[0]]);
        }
        
        return null;
    } catch (error) {
        console.error('Erro na requisição à API:', error);
        throw error;
    }
}

// Destacar um bairro no mapa
function highlightBairro(bairroName) {
    // Resetar todos os bairros
    for (const name in bairroLayers) {
        bairroLayers[name].forEach(layer => {
            layer.setStyle({
                weight: 5,
                opacity: 0.7
            });
        });
    }
    
    // Destacar o bairro selecionado
    if (bairroLayers[bairroName]) {
        bairroLayers[bairroName].forEach(layer => {
            layer.setStyle({
                weight: 8,
                opacity: 1
            });
            layer.bringToFront();
        });
        
        // Ajustar a visualização do mapa para mostrar a rota
        if (bairroLayers[bairroName].length > 0) {
            const bounds = L.latLngBounds();
            bairroLayers[bairroName].forEach(layer => {
                bounds.extend(layer.getBounds());
            });
            map.fitBounds(bounds, { padding: [50, 50] });
        }
    }
}

// Limpar o mapa
function clearMap() {
    // Limpar todas as camadas de overlay
    for (const key in overlayMaps) {
        overlayMaps[key].clearLayers();
    }
    
    bairroLayers = {};
    draggableMarkers = [];
    unificacoes = [];
    selectedBairros.clear();
    renderUnificacoesList();
    adicionarPontosReferencia();
    
    // Desativar modo de edição se estiver ativo
    if (editMode) {
        editMode = false;
        document.getElementById('toggle-edit').innerHTML = '<i class="fas fa-edit"></i> Mover Pontos';
        document.getElementById('toggle-edit').style.background = 'linear-gradient(to right, #1abc9c, #16a085)';
        document.getElementById('edit-message').style.display = 'none';
    }
}

// Atualizar status
function updateStatus(message, type) {
    const statusElement = document.getElementById('status');
    let icon = '<i class="fas fa-info-circle"></i>';
    
    if (type === 'error') {
        statusElement.style.backgroundColor = 'rgba(231, 76, 60, 0.2)';
        statusElement.style.borderLeftColor = '#e74c3c';
        icon = '<i class="fas fa-exclamation-circle"></i>';
    } else if (type === 'success') {
        statusElement.style.backgroundColor = 'rgba(46, 204, 113, 0.2)';
        statusElement.style.borderLeftColor = '#2ecc71';
        icon = '<i class="fas fa-check-circle"></i>';
    } else if (type === 'loading') {
        statusElement.style.backgroundColor = 'rgba(241, 196, 15, 0.2)';
        statusElement.style.borderLeftColor = '#f1c40f';
        icon = '<span class="loading"></span>';
    }
    
    statusElement.innerHTML = `${icon} ${message}`;
}

// Mostrar erro
function showError(message) {
    updateStatus(message, 'error');
}

// Atualizar estatísticas
function updateStats() {
    document.getElementById('total-colaboradores').textContent = `${colaboradoresData.length} colaboradores`;
    
    const bairrosCount = Object.keys(bairroLayers).length;
    document.getElementById('total-bairros').textContent = `${bairrosCount} bairros`;
}

// Reiniciar o mapa
document.getElementById('reset-map').addEventListener('click', function() {
    clearMap();
    colaboradoresData = [];
    bairrosData = null;
    generatedRoutesData = {};
    document.getElementById('excel-file').value = '';
    document.getElementById('geojson-file').value = '';
    document.getElementById('bairros-list').innerHTML = '';
    document.querySelector('.unificar-section').style.display = 'none';
    updateStatus('Mapa reiniciado. Aguardando upload de dados...', 'success');
    updateStats();
    
    // Resetar barra de progresso
    document.getElementById('progress-bar').style.width = '0%';
});

// Fechar informações da rota
document.getElementById('close-route-info').addEventListener('click', function() {
    document.getElementById('route-info').style.display = 'none';
});

// Função para criar o conteúdo de um arquivo KML
function createKMLString(routeName, coordinates) {
    // KML formata as coordenadas como longitude,latitude,altitude
    const coordsString = coordinates.map(c => `${c[1]},${c[0]},0`).join(' ');

    return `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>${routeName}</name>
    <Style id="routeStyle">
      <LineStyle>
        <color>7f0e7fff</color> <width>4</width>
      </LineStyle>
    </Style>
    <Placemark>
      <name>${routeName}</name>
      <styleUrl>#routeStyle</styleUrl>
      <LineString>
        <tessellate>1</tessellate>
        <coordinates>${coordsString}</coordinates>
      </LineString>
    </Placemark>
  </Document>
</kml>`;
}

// Evento para o botão de download KML
document.getElementById('download-kml-btn').addEventListener('click', function() {
    if (Object.keys(generatedRoutesData).length === 0) {
        alert('Nenhuma rota foi gerada para baixar.');
        return;
    }

    updateStatus('Gerando arquivo ZIP...', 'loading');
    const zip = new JSZip();

    for (const routeName in generatedRoutesData) {
        const coordinates = generatedRoutesData[routeName];
        const kmlString = createKMLString(routeName, coordinates);
        // Sanitiza o nome do bairro para criar um nome de arquivo válido
        const fileName = routeName.replace(/[^a-z0-9]/gi, '_').toLowerCase();
        zip.file(`${fileName}.kml`, kmlString);
    }

    zip.generateAsync({ type: 'blob' })
        .then(function(content) {
            const link = document.createElement('a');
            link.href = URL.createObjectURL(content);
            link.download = 'rotas_kml.zip';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            updateStatus('Arquivo ZIP gerado com sucesso.', 'success');
        })
        .catch(err => {
            console.error("Erro ao gerar o ZIP:", err);
            showError("Erro ao gerar o arquivo ZIP.");
        });
});

// Evento para o botão de download Excel
document.getElementById('download-excel-btn').addEventListener('click', function() {
    if (Object.keys(colaboradoresPorRota).length === 0) {
        alert('Nenhuma rota foi processada para baixar o arquivo Excel.');
        return;
    }

    const wsData = [["MATRÍCULA", "COLABORADOR", "BAIRRO", "ROTA ALOCADA"]];

    for (const rotaName in colaboradoresPorRota) {
        colaboradoresPorRota[rotaName].forEach(colab => {
            wsData.push([
                colab.MATRICULA,
                colab.COLABORADOR,
                colab.BAIRRO_ORIGEM,
                rotaName
            ]);
        });
    }

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rotas por Colaborador");

    XLSX.writeFile(wb, "rotas_colaboradores.xlsx");
});

// Incluir o script leaflet-pip para a verificação de polígonos
const script = document.createElement('script');
script.src = 'https://unpkg.com/leaflet-pip@1.1.0/leaflet-pip.js';
document.head.appendChild(script);

// Inicializar o mapa quando a página carregar
window.onload = initMap;