// State management
let playersData = [];
let passwordStaff = "staff123";

// DOM Elements
const loginScreen = document.getElementById('login-screen');
const dashboardScreen = document.getElementById('dashboard-screen');
const loginBtn = document.getElementById('login-btn');
const passwordInput = document.getElementById('password-input');
const loginError = document.getElementById('login-error');

const uploadBtn = document.getElementById('upload-excel-btn');
const fileInput = document.getElementById('excel-file-input');
const searchInput = document.getElementById('search-input');
const playersContainer = document.getElementById('players-list');

const statTotalPlayers = document.getElementById('stat-total-players');
const statCash = document.getElementById('stat-cash');
const statPos = document.getElementById('stat-pos');

const chatFab = document.getElementById('chat-fab');
const chatPanel = document.getElementById('chat-panel');
const closeChatBtn = document.getElementById('close-chat-btn');
const chatMessages = document.getElementById('chat-messages');
const chatInput = document.getElementById('chat-input');
const sendChatBtn = document.getElementById('send-chat-btn');

// Settings Elements
const settingsBtn = document.getElementById('settings-btn');
const exportBtn = document.getElementById('export-btn');
const settingsModal = document.getElementById('settings-modal');
const closeSettingsBtn = document.getElementById('close-settings-btn');
const clearDataBtn = document.getElementById('clear-data-btn');
const openaiKeyInput = document.getElementById('openai-key-input');
const firebaseUrlInput = document.getElementById('firebase-url-input');

const DEFAULT_FIREBASE_URL = "https://asd-checkin-default-rtdb.europe-west1.firebasedatabase.app/";
let isSyncing = false;

// Initialization
function init() {
    // Check if logged in
    if(sessionStorage.getItem('isLoggedIn') === 'true') {
        showDashboard();
    }
    
    // Load from local storage if exists
    const savedData = localStorage.getItem('asdTournamentData');
    if(savedData) {
        playersData = JSON.parse(savedData);
        renderPlayers();
    }

    // Start Cloud Sync with Default or Custom URL
    syncFromCloud();
    setInterval(syncFromCloud, 5000); // Polling every 5 seconds
}

async function syncToCloud() {
    const url = localStorage.getItem('firebaseUrl') || DEFAULT_FIREBASE_URL;
    if(!url || isSyncing) return;
    
    try {
        const cleanUrl = url.endsWith('/') ? url.slice(0, -1) : url;
        
        // Sync Players Data
        await fetch(`${cleanUrl}/tournament.json`, {
            method: 'PUT',
            body: JSON.stringify(playersData)
        });

        // Sync API Key Config (so all devices get it)
        const apiKey = localStorage.getItem('openaiApiKey');
        if(apiKey) {
            await fetch(`${cleanUrl}/config.json`, {
                method: 'PUT',
                body: JSON.stringify({ apiKey: apiKey })
            });
        }
    } catch(e) { console.error("Sync To Cloud Error:", e); }
}

async function syncFromCloud() {
    const url = localStorage.getItem('firebaseUrl') || DEFAULT_FIREBASE_URL;
    if(!url || isSyncing) return;
    
    try {
        const cleanUrl = url.endsWith('/') ? url.slice(0, -1) : url;
        
        // 1. Sync Players
        const resp = await fetch(`${cleanUrl}/tournament.json`);
        const data = await resp.json();
        if(data && JSON.stringify(data) !== JSON.stringify(playersData)) {
            playersData = data;
            renderPlayers(searchInput.value);
        }

        // 2. Sync Config (API Key)
        const configResp = await fetch(`${cleanUrl}/config.json`);
        const config = await configResp.json();
        if(config && config.apiKey && !localStorage.getItem('openaiApiKey')) {
            localStorage.setItem('openaiApiKey', config.apiKey);
            console.log("API Key sincronizzata dal cloud!");
        }
    } catch(e) { console.error("Sync From Cloud Error:", e); }
}

// Login
loginBtn.addEventListener('click', () => {
    if(passwordInput.value === passwordStaff) {
        sessionStorage.setItem('isLoggedIn', 'true');
        showDashboard();
    } else {
        loginError.classList.remove('hidden');
    }
});

function showDashboard() {
    loginScreen.classList.add('hidden');
    dashboardScreen.classList.remove('hidden');
    renderPlayers();
}

// Excel Upload & Parsing
uploadBtn.addEventListener('click', () => fileInput.click());

fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if(!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        
        // Use 'File Iscrizioni' if exists, otherwise first sheet
        let sheetName = workbook.SheetNames.find(n => n.toLowerCase().includes('iscrizioni')) || workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Try to find the header row automatically
        let rawData = [];
        for (let i = 0; i < 10; i++) {
            rawData = XLSX.utils.sheet_to_json(worksheet, { range: i });
            if (rawData.length > 0) {
                const firstRowKeys = Object.keys(rawData[0]).join(',').toLowerCase();
                if (firstRowKeys.includes('giocatore') || firstRowKeys.includes('nome') || firstRowKeys.includes('atleta')) {
                    console.log(`Trovata riga intestazione alla riga ${i+1}`);
                    break;
                }
            }
        }
        
        processExcelData(rawData);
    };
    reader.readAsArrayBuffer(file);
});

function processExcelData(rawData) {
    console.log("Raw Data parsed:", rawData);
    
    if (!rawData || rawData.length === 0) {
        alert("Il file caricato sembra vuoto.");
        return;
    }

    const grouped = {};

    rawData.forEach((row, index) => {
        const keys = Object.keys(row);
        const getVal = (possibleKeys) => {
            const key = keys.find(k => possibleKeys.some(pk => k.toLowerCase().includes(pk.toLowerCase())));
            return key ? row[key] : '';
        };

        // Try to get name from 'Giocatore Iscritto' or split 'Nome'/'Cognome'
        let fullName = String(getVal(['giocatore', 'atleta', 'nominativo', 'nome e cognome'])).trim();
        let nome = "";
        let cognome = "";

        if (fullName && fullName !== "undefined") {
            const parts = fullName.split(' ');
            nome = parts[0];
            cognome = parts.slice(1).join(' ');
        } else {
            nome = String(getVal(['nome', 'name'])).trim();
            cognome = String(getVal(['cognome', 'surname', 'last name'])).trim();
            fullName = `${nome} ${cognome}`.trim();
        }

        if (!fullName || fullName === "undefined" || fullName === "") {
            console.warn(`Riga ${index + 1} saltata: Nome non trovato`, row);
            return;
        }

        const categoria = String(getVal(['categoria', 'torneo', 'gara', 'evento']) || 'Iscrizione').trim();
        
        // Cost parsing: handle things like "? 25,00" or "25.00€"
        let costoStr = String(getVal(['costo', 'quota', 'prezzo', 'importo', '€', 'iscrizione']));
        let costo = parseFloat(costoStr.replace(',', '.').replace(/[^\d.-]/g, ''));
        if (isNaN(costo)) costo = 0;

        const fullNameId = fullName.toLowerCase();

        if (!grouped[fullNameId]) {
            grouped[fullNameId] = {
                id: Math.random().toString(36).substr(2, 9),
                nome: nome || fullName,
                cognome: cognome || "",
                totaleDovuto: 0,
                statoPagamento: 'DA_PAGARE',
                metodoPagamento: 'CONTANTI',
                timestampPagamento: null,
                iscrizioni: []
            };
            
            // Check if already paid in Excel (some sheets have 'Contanti' or 'Pos' filled)
            const paidCash = getVal(['contanti']);
            const paidPos = getVal(['pos']);
            if (paidCash || paidPos) {
                grouped[fullNameId].statoPagamento = 'PAGATO';
                grouped[fullNameId].metodoPagamento = paidPos ? 'POS' : 'CONTANTI';
            }
        }

        // Add inscription
        grouped[fullNameId].iscrizioni.push({
            id: Math.random().toString(36).substr(2, 9),
            categoria: categoria,
            costo: costo,
            kitConsegnato: String(getVal(['kit', 'pacco'])).toLowerCase().includes('si'),
            timestampKit: null
        });

        grouped[fullNameId].totaleDovuto += costo;
    });

    const newPlayersData = Object.values(grouped);
    console.log("Processed Players:", newPlayersData);

    if (newPlayersData.length === 0) {
        alert("Non è stato possibile trovare nomi di giocatori nel file. Verifica le intestazioni delle colonne.");
        return;
    }
    
    newPlayersData.forEach(newPlayer => {
        const existingPlayer = playersData.find(p => 
            p.nome.toLowerCase() === newPlayer.nome.toLowerCase() && 
            p.cognome.toLowerCase() === newPlayer.cognome.toLowerCase()
        );
        
        if(existingPlayer) {
            newPlayer.statoPagamento = existingPlayer.statoPagamento;
            newPlayer.metodoPagamento = existingPlayer.metodoPagamento;
            newPlayer.timestampPagamento = existingPlayer.timestampPagamento;
            // Match kits
            newPlayer.iscrizioni.forEach(newIscr => {
                const existingIscr = existingPlayer.iscrizioni.find(ei => ei.categoria === newIscr.categoria);
                if(existingIscr) {
                    newIscr.kitConsegnato = existingIscr.kitConsegnato;
                    newIscr.timestampKit = existingIscr.timestampKit;
                }
            });
        }
    });

    playersData = newPlayersData;
    saveData();
    renderPlayers();
}

function saveData() {
    localStorage.setItem('asdTournamentData', JSON.stringify(playersData));
    updateStats();
    syncToCloud();
}

// Rendering
function renderPlayers(filterText = '') {
    playersContainer.innerHTML = '';
    
    if(playersData.length === 0) {
        playersContainer.innerHTML = `
            <div class="empty-state">
                <i class="fa-solid fa-file-arrow-up"></i>
                <p>Nessun dato presente. Carica il file Excel delle iscrizioni.</p>
            </div>
        `;
        updateStats();
        return;
    }

    const filtered = playersData.filter(p => 
        `${p.nome} ${p.cognome}`.toLowerCase().includes(filterText.toLowerCase())
    );

    // Sort: Unpaid first, Paid at the bottom
    filtered.sort((a, b) => {
        if(a.statoPagamento === 'DA_PAGARE' && b.statoPagamento === 'PAGATO') return -1;
        if(a.statoPagamento === 'PAGATO' && b.statoPagamento === 'DA_PAGARE') return 1;
        return 0;
    });

    filtered.forEach(player => {
        const isPaid = player.statoPagamento === 'PAGATO';
        const allKitsDelivered = player.iscrizioni.every(i => i.kitConsegnato);

        const card = document.createElement('div');
        card.className = `player-card glass ${isPaid ? 'paid' : ''}`;
        
        // Generate Inscriptions HTML
        const iscrizioniHtml = player.iscrizioni.map((iscr, idx) => `
            <div class="inscription-item">
                <span class="category-name">${iscr.categoria} (${iscr.costo}€)</span>
                <label class="checkbox-wrapper">
                    <span class="checkbox-label">Kit</span>
                    <input type="checkbox" onchange="toggleKit('${player.id}', ${idx})" ${iscr.kitConsegnato ? 'checked' : ''}>
                </label>
            </div>
        `).join('');

        let paymentActionsHtml = "";
        if(isPaid) {
            paymentActionsHtml = `
                <div class="action-buttons">
                    <button class="btn btn-pay" disabled>
                        <i class="fa-solid fa-check-circle"></i> Pagato (${player.metodoPagamento})
                    </button>
                    <button class="btn danger-btn" style="flex: 0; padding: 10px;" onclick="cancelPayment('${player.id}')" title="Annulla Pagamento">
                        <i class="fa-solid fa-rotate-left"></i>
                    </button>
                </div>
            `;
        } else {
            paymentActionsHtml = `
                <div class="action-buttons" id="pay-actions-${player.id}">
                    <button class="btn btn-pay" onclick="showPaymentOptions('${player.id}')">
                        <i class="fa-solid fa-hand-holding-dollar"></i> Paga Tutto
                    </button>
                </div>
            `;
        }

        card.innerHTML = `
            <div class="card-header">
                <div class="player-name">
                    <div class="status-dot"></div>
                    ${player.nome} ${player.cognome}
                </div>
                <div class="total-due">${player.totaleDovuto}€</div>
            </div>
            <div class="card-body">
                ${iscrizioniHtml}
            </div>
            <div class="card-actions">
                ${paymentActionsHtml}
                <div class="action-buttons" style="margin-top: 8px;">
                    <button class="btn btn-kit" onclick="deliverAllKits('${player.id}')" ${allKitsDelivered ? 'disabled' : ''}>
                        <i class="fa-solid fa-box-open"></i> ${allKitsDelivered ? 'Kit Consegnati' : 'Consegna Tutti i Kit'}
                    </button>
                </div>
            </div>
        `;
        playersContainer.appendChild(card);
    });

    updateStats();
}

// Actions
window.toggleKit = function(playerId, iscrIndex) {
    const player = playersData.find(p => p.id === playerId);
    if(player) {
        player.iscrizioni[iscrIndex].kitConsegnato = !player.iscrizioni[iscrIndex].kitConsegnato;
        player.iscrizioni[iscrIndex].timestampKit = new Date().toISOString();
        saveData();
        renderPlayers(searchInput.value);
    }
};

window.showPaymentOptions = function(playerId) {
    const container = document.getElementById(`pay-actions-${playerId}`);
    if(container) {
        container.innerHTML = `
            <button class="btn" style="background: #34d399; color: white; flex: 1;" onclick="confirmPayment('${playerId}', 'CONTANTI')">Contanti</button>
            <button class="btn" style="background: #60a5fa; color: white; flex: 1;" onclick="confirmPayment('${playerId}', 'POS')">POS</button>
            <button class="btn" style="background: rgba(255,255,255,0.1); color: white; flex: 0;" onclick="renderPlayers(searchInput.value)"><i class="fa-solid fa-xmark"></i></button>
        `;
    }
};

window.confirmPayment = function(playerId, method) {
    const player = playersData.find(p => p.id === playerId);
    if(player) {
        player.metodoPagamento = method;
        player.statoPagamento = 'PAGATO';
        player.timestampPagamento = new Date().toISOString();
        saveData();
        renderPlayers(searchInput.value);
    }
};

window.cancelPayment = function(playerId) {
    if(confirm("Sei sicuro di voler annullare questo pagamento?")) {
        const player = playersData.find(p => p.id === playerId);
        if(player) {
            player.statoPagamento = 'DA_PAGARE';
            player.timestampPagamento = null;
            saveData();
            renderPlayers(searchInput.value);
        }
    }
};

window.deliverAllKits = function(playerId) {
    const player = playersData.find(p => p.id === playerId);
    if(player) {
        player.iscrizioni.forEach(i => {
            i.kitConsegnato = true;
            i.timestampKit = new Date().toISOString();
        });
        saveData();
        renderPlayers(searchInput.value);
    }
};

// Search
searchInput.addEventListener('input', (e) => {
    renderPlayers(e.target.value);
});

function updateStats() {
    let totPlayers = playersData.length;
    let totCash = 0;
    let totPos = 0;

    playersData.forEach(p => {
        if(p.statoPagamento === 'PAGATO') {
            if(p.metodoPagamento === 'CONTANTI') totCash += p.totaleDovuto;
            if(p.metodoPagamento === 'POS') totPos += p.totaleDovuto;
        }
    });

    statTotalPlayers.innerText = totPlayers;
    statCash.innerText = totCash + '€';
    statPos.innerText = totPos + '€';
}

// Chat AI Logic (Mocked logic for demo purposes as requested)
chatFab.addEventListener('click', () => {
    chatPanel.classList.remove('hidden');
});
closeChatBtn.addEventListener('click', () => {
    chatPanel.classList.add('hidden');
});

function addChatMessage(text, sender) {
    const div = document.createElement('div');
    div.className = `message ${sender}`;
    div.innerText = text;
    chatMessages.appendChild(div);
    chatMessages.scrollTop = chatMessages.scrollHeight;
}

sendChatBtn.addEventListener('click', handleChat);
chatInput.addEventListener('keypress', (e) => {
    if(e.key === 'Enter') handleChat();
});

function handleChat() {
    const query = chatInput.value.trim();
    if(!query) return;

    addChatMessage(query, 'user');
    chatInput.value = '';

    const apiKey = localStorage.getItem('openaiApiKey');
    if(apiKey && apiKey.length > 20) {
        callGemini(query, apiKey);
    } else {
        handleLocalChat(query.toLowerCase());
    }
}

async function callGemini(query, apiKey) {
    addChatMessage("Sto pensando...", 'ai');
    const messages = chatMessages.querySelectorAll('.message');
    const lastMessage = messages[messages.length - 1]; 

    const contextData = playersData.map(p => {
        return `${p.nome} ${p.cognome}: Dovuto ${p.totaleDovuto}€, Pagato: ${p.statoPagamento === 'PAGATO' ? 'Sì ('+p.metodoPagamento+')' : 'No'}, Kit: ${p.iscrizioni.every(i => i.kitConsegnato) ? 'Tutti ritirati' : 'Da ritirare'}`;
    }).join('\n');

    // List of models to try in order based on user console
    const models = ["gemini-3.1-flash", "gemini-3-flash", "gemini-2.5-flash", "gemini-1.5-flash"];
    let success = false;

    for(let model of models) {
        if(success) break;
        try {
            const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    contents: [{
                        role: "user",
                        parts: [{ text: "DATI TORNEO:\n" + contextData + "\n\nDOMANDA: " + query }]
                    }]
                })
            });

            if(response.ok) {
                const data = await response.json();
                if(data.candidates && data.candidates[0].content.parts[0].text) {
                    lastMessage.innerText = data.candidates[0].content.parts[0].text;
                    success = true;
                }
            } else {
                console.warn(`Modello ${model} fallito con stato ${response.status}`);
            }
        } catch (e) {
            console.error(`Errore con modello ${model}:`, e);
        }
    }

    if(!success) {
        lastMessage.innerText = "L'assistente ha riscontrato un problema tecnico con la tua chiave Google. Verifica che il piano sia attivo su AI Studio o riprova più tardi.";
    }
}

function handleLocalChat(query) {
    setTimeout(() => {
        let response = "Mi dispiace, non ho capito la richiesta. Puoi chiedermi degli incassi, di chi ha pagato/manca, o dei kit.";

        if((query.includes('chi') || query.includes('elenco')) && query.includes('pagato') && !query.includes('non')) {
            const pagati = playersData.filter(p => p.statoPagamento === 'PAGATO');
            if(pagati.length === 0) response = "Al momento nessuno ha ancora saldato la quota.";
            else response = `Hanno pagato ${pagati.length} persone: ` + pagati.map(p => `${p.nome} ${p.cognome}`).join(', ') + ".";
        }
        else if(query.includes('chi manca') || query.includes('chi deve pagare') || query.includes('non ha pagato')) {
            const daPagare = playersData.filter(p => p.statoPagamento === 'DA_PAGARE');
            if(daPagare.length === 0) response = "Ottime notizie! Tutti i presenti hanno pagato.";
            else response = `Mancano all'appello ${daPagare.length} giocatori: ` + daPagare.map(p => `${p.nome} ${p.cognome}`).join(', ') + ".";
        }
        else if(query.includes('kit') || query.includes('welcome pack') || query.includes('pacco')) {
            const mancanoKit = playersData.filter(p => p.iscrizioni.some(i => !i.kitConsegnato));
            if(mancanoKit.length === 0) response = "Tutti i kit sono stati consegnati correttamente.";
            else response = `Ci sono ${mancanoKit.length} giocatori che devono ancora ritirare almeno un kit.`;
        }
        else if(query.includes('incass') || query.includes('soldi') || query.includes('totale') || query.includes('contanti') || query.includes('pos')) {
            let cash = 0, pos = 0;
            playersData.forEach(p => {
                if(p.statoPagamento === 'PAGATO') {
                    if(p.metodoPagamento === 'CONTANTI') cash += p.totaleDovuto;
                    else pos += p.totaleDovuto;
                }
            });
            response = `Resoconto incassi:\n- Contanti: ${cash}€\n- POS: ${pos}€\nTotale complessivo: ${cash+pos}€.`;
        }
        else {
            const playerFound = playersData.find(p => query.includes(p.cognome.toLowerCase()) || query.includes(p.nome.toLowerCase()));
            if(playerFound) {
                if(playerFound.statoPagamento === 'PAGATO' && playerFound.timestampPagamento) {
                    const time = new Date(playerFound.timestampPagamento).toLocaleTimeString('it-IT');
                    response = `${playerFound.nome} ${playerFound.cognome} ha pagato alle ore ${time}.`;
                } else {
                    response = `${playerFound.nome} ${playerFound.cognome} non ha ancora pagato.`;
                }
            } else {
                response = "Non trovo il giocatore. Prova a specificare il cognome (es. 'A che ora ha pagato Rossi?').";
            }
        }

        addChatMessage(response, 'ai');
    }, 500);
}

// Settings Logic
const saveApiBtn = document.getElementById('save-api-btn');
const closeSettingsBtnTop = document.getElementById('close-settings-btn-top');

settingsBtn.addEventListener('click', () => {
    openaiKeyInput.value = localStorage.getItem('openaiApiKey') || '';
    firebaseUrlInput.value = localStorage.getItem('firebaseUrl') || '';
    settingsModal.classList.remove('hidden');
});

if(saveApiBtn) {
    saveApiBtn.addEventListener('click', () => {
        const key = openaiKeyInput.value.trim();
        const fireUrl = firebaseUrlInput.value.trim();
        
        localStorage.setItem('openaiApiKey', key);
        localStorage.setItem('firebaseUrl', fireUrl);
        
        alert("Impostazioni salvate correttamente!");
        
        if(fireUrl) {
            syncFromCloud();
            if(playersData.length > 0) syncToCloud(); // Push local data to cloud immediately
            if(!window.syncInterval) {
                window.syncInterval = setInterval(syncFromCloud, 5000);
            }
        }
        
        settingsModal.classList.add('hidden');
    });
}

closeSettingsBtn.addEventListener('click', () => {
    settingsModal.classList.add('hidden');
});
if(closeSettingsBtnTop) {
    closeSettingsBtnTop.addEventListener('click', () => {
        settingsModal.classList.add('hidden');
    });
}

clearDataBtn.addEventListener('click', () => {
    if(confirm("Sei sicuro di voler ELIMINARE TUTTI I DATI correnti? Questa operazione non è annullabile.")) {
        localStorage.removeItem('asdTournamentData');
        playersData = [];
        renderPlayers();
        settingsModal.classList.add('hidden');
    }
});

// Export Logic
exportBtn.addEventListener('click', () => {
    if(playersData.length === 0) {
        alert("Nessun dato da esportare.");
        return;
    }
    
    // Prepare flattened data for export
    const exportData = [];
    playersData.forEach(p => {
        p.iscrizioni.forEach(iscr => {
            exportData.push({
                "Giocatore": `${p.nome} ${p.cognome}`,
                "Categoria": iscr.categoria,
                "Costo Categoria": iscr.costo,
                "Kit Consegnato": iscr.kitConsegnato ? "SI" : "NO",
                "Totale Dovuto dal Giocatore": p.totaleDovuto,
                "Stato Pagamento": p.statoPagamento,
                "Metodo Pagamento": p.metodoPagamento,
                "Data Pagamento": p.timestampPagamento ? new Date(p.timestampPagamento).toLocaleString('it-IT') : ""
            });
        });
    });

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Report Torneo");
    XLSX.writeFile(workbook, `Report_Torneo_${new Date().toISOString().slice(0,10)}.xlsx`);
});

// Run
init();
