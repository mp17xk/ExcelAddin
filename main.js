import campiImportConfig from './campi-import-config.json';
import smistamentoConfig from './smistamento-config.json';
import * as XLSX from 'xlsx';

// Variabile globale per contenere la tabella grezza del CSV (array 2D) per la scrivania
let csvRawData = null;
// Variabile per contenere gli Oggetti mappati per facilitare lo smistamento
let csvObjects = null;

// ── Stato globale per la sezione MIGRAZIONE ──────────────────────────────
let migSourceData = null;       // { sheetNames: [...], sheets: { name: { headers, rows } } }
let migDestColumnsCache = {};   // cache: sheetName -> [headerStrings]
let migDestSheetsGlobal = [];   // cache dei fogli dest presenti nel workbook

// L'API di base di Microsoft: aspetta che Excel abbia finito di caricare tutto il pannellino
Office.onReady((info) => {
    // ── TAB SWITCHER ────────────────────────────────────────────────────
    // Colleghiamo i tab fuori dal check Excel, così l'UI è cliccabile anche in un browser normale
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            document.querySelectorAll('.tab-panel').forEach(p => p.style.display = 'none');
            document.getElementById('tab-' + btn.dataset.tab).style.display = 'block';
        });
    });

    // ── EVENT LISTENER GLOBALI ──────────────────────────────────────────
    // Listener Importazione CSV
    document.getElementById('csvFileInput').addEventListener('change', handleFileSelect);
    document.getElementById('importBtn').addEventListener('click', importCsvToExcel);
    document.getElementById('closeReportBtn').addEventListener('click', () => {
        document.getElementById('reportModal').style.display = 'none';
    });

    // Listener Migrazione Spreadsheet
    document.getElementById('migFileInput').addEventListener('change', handleMigrationFileSelect);
    document.getElementById('migAddColBtn').addEventListener('click', () => {
        if(migSourceData && migSourceData.sheetNames.length > 0) {
            addMappingRow(migSourceData.sheetNames[0], null);
        }
    });
    document.getElementById('migExecuteBtn').addEventListener('click', executeMigration);

    // Verifichiamo di essere davvero dentro Excel per le funzioni core
    if (info.host === Office.HostType.Excel) {
        document.getElementById('statusMessage').textContent = "Add-in Gestionale pronto!";
        document.getElementById('statusMessage').className = "status-message success";
    } else {
        document.getElementById('statusMessage').textContent = "Aperto nel browser (UX Test)";
        document.getElementById('statusMessage').className = "status-message";
    }
});

// Fase 1: Leggere il file che l'utente ha scelto. Niente server esterno, accade in RAM sul browser.
function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) {
        return;
    }

    const reader = new FileReader();

    reader.onload = (e) => {
        const text = e.target.result;

        // Passa il testo bruto alla funzioncina di interpretazione CSV
        csvRawData = parseCSV(text);
        csvObjects = convertToObjects(csvRawData);

        // Aggiorna un po' di grafica
        document.getElementById('fileInfo').style.display = 'block';
        document.getElementById('fileName').textContent = file.name;
        document.getElementById('fileRows').textContent = csvObjects.length;
        document.getElementById('importBtn').disabled = false;

        showStatus("File CSV letto in memoria.", "success");
    };

    reader.onerror = (error) => {
        showStatus("Errore fatale nella lettura del file da PC.", "error");
        console.error(error);
    };

    reader.readAsText(file);
}

// Helper fondamentale: in presenza di date in formato IT (GG/MM/AAAA), Excel Web (che gira in US internally)
// le inverte capendole come (MM/GG/AAAA) sfasando mesi coi giorni (es: 08/01/2026 diventa 1 Agosto).
// Questa funzione intercetta i pattern gg/mm/aaaa e li blinda nel formato universale ISO YYYY-MM-DD
function forceISODates(val) {
    if (typeof val === "string") {
        let v = val.trim();
        let dateRegex = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
        let match = v.match(dateRegex);
        if (match) {
            // YYYY-MM-DD. A quel punto Excel legge la ISO e la auto-formatta come VERA data in locale IT.
            return `${match[3]}-${match[2].padStart(2, '0')}-${match[1].padStart(2, '0')}`;
        }
    }
    return val;
}

// Converte la matrice 2D grezza in un array di JavaScript Objects leggendo le intestazioni (riga 0)
function convertToObjects(matrix) {
    if (!matrix || matrix.length < 2) return [];
    const headers = matrix[0];
    const objects = [];

    for (let i = 1; i < matrix.length; i++) {
        const row = matrix[i];
        let obj = {};
        for (let j = 0; j < headers.length; j++) {
            let key = headers[j] ? headers[j].trim() : "Colonna_" + j;
            obj[key] = row[j] !== undefined ? row[j].trim() : "";
        }
        objects.push(obj);
    }
    return objects;
}

// Analizzatore avanzato di CSV
function parseCSV(text) {
    const rows = [];
    let currentRow = [];
    let inQuotes = false;
    let currentValue = "";

    let delimiter = ',';
    if (text.split(';').length > text.split(',').length) {
        delimiter = ';';
    }

    for (let i = 0; i < text.length; i++) {
        let char = text[i];
        let nextChar = text[i + 1];

        if (inQuotes) {
            if (char === '"' && nextChar === '"') {
                currentValue += '"';
                i++;
            } else if (char === '"') {
                inQuotes = false;
            } else {
                currentValue += char;
            }
        } else {
            if (char === '"') {
                inQuotes = true;
            } else if (char === delimiter) {
                currentRow.push(currentValue);
                currentValue = "";
            } else if (char === '\n' || (char === '\r' && nextChar === '\n')) {
                if (char === '\r') i++;
                currentRow.push(currentValue);
                if (currentRow.length > 1 || (currentRow.length === 1 && currentRow[0] !== "")) {
                    rows.push(currentRow);
                }
                currentRow = [];
                currentValue = "";
            } else {
                currentValue += char;
            }
        }
    }

    if (currentRow.length > 0 || currentValue.length > 0) {
        currentRow.push(currentValue);
        rows.push(currentRow);
    }

    if (rows.length === 0) return rows;

    let maxCols = 0;
    for (let r of rows) {
        if (r.length > maxCols) maxCols = r.length;
    }

    for (let r of rows) {
        while (r.length < maxCols) {
            r.push("");
        }
    }

    return rows;
}

// Fase FINALE: iniettare fisicamente la struttura e applicare le regole
async function importCsvToExcel() {
    if (!csvRawData || csvRawData.length === 0) return;

    document.getElementById('importBtn').disabled = true;
    showStatus("Analisi e smistamento in corso... attendere.", "success");

    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;

            // --- FASE A: AZZERAMENTO E POPOLAMENTO FOGLIO 'importa' (La Scrivania) ---
            let importaSheet = sheets.getItemOrNullObject("importa");
            await context.sync();

            if (importaSheet.isNullObject) {
                importaSheet = sheets.add("importa");
            }

            // Cancella tutto il contenuto esistente
            importaSheet.getRange().clear();

            const totalRows = csvRawData.length;
            const totalCols = csvRawData[0].length;
            const importaRange = importaSheet.getCell(0, 0).getResizedRange(totalRows - 1, totalCols - 1);

            // Formattiamo come Testo ("@") SOLO la colonna del "Codice pratica" 
            // per evitare che Excel trasformi i codici (es: "1-24-2026") in date, 
            // ma lasciamo "General" sulle altre per permettere alle vere date di restare ordinabili!
            let idxCodeImporta = csvRawData[0].indexOf("Codice pratica");
            if (idxCodeImporta === -1) idxCodeImporta = 0;
            
            let formatsImporta = csvRawData.map(row => row.map((val, colIndex) => {
                return (colIndex === idxCodeImporta) ? "@" : "General";
            }));
            
            importaRange.numberFormat = formatsImporta;

            // Applichiamo la benda anti-falsamento di Excel per le date
            let importDataSafe = csvRawData.map(r => r.map(forceISODates));
            importaRange.values = importDataSafe;
            importaRange.format.autofitColumns();

            // --- FASE B: ANALISI E SMISTAMENTO TRAMITE JSON (CASCADE / WATERFALL) ---
            const fogliDestinazione = {};

            for (let rowObj of csvObjects) {
                let hasMatchedRule = false;

                // Le regole nel json vengono lette in ordine sequenziale (Top-Bottom).
                // Al primo scontro positivo, la pratica viene smistata e il ciclo si ferma (Mutuamente esclusivo).
                for (let rule of smistamentoConfig) {
                    let targetFoglio = rule.foglio;
                    let targetCampo = rule.campo;
                    let ruleValore = rule.valore;
                    let match = false;

                    let cellVal = (rowObj[targetCampo] || "").trim();

                    if (rule.regola === "START-WITH") {
                        if (Array.isArray(ruleValore)) {
                            match = ruleValore.some(pref => cellVal.startsWith(String(pref).trim()));
                        } else {
                            match = cellVal.startsWith(String(ruleValore).trim());
                        }
                    } else { // Exact match by default
                        if (Array.isArray(ruleValore)) {
                            match = ruleValore.some(v => cellVal.toLowerCase() === String(v).trim().toLowerCase());
                        } else {
                            match = cellVal.toLowerCase() === String(ruleValore).trim().toLowerCase();
                        }
                    }

                    if (match) {
                        hasMatchedRule = true;
                        if (!fogliDestinazione[targetFoglio]) fogliDestinazione[targetFoglio] = [];

                        let mappedRow = {};
                        for (let colFinale in campiImportConfig) {
                            let colSorgente = campiImportConfig[colFinale];
                            mappedRow[colFinale] = rowObj[colSorgente] || "";
                        }

                        // --- REGOLE CONDIZIONALI SULLE DATE ---
                        const attivitaOriginale = (rowObj["Attività da espletare"] || "").trim();
                        const dataPura = (rowObj["Data attivazione"] || "").trim();

                        if (attivitaOriginale === "Istruttoria Tecnica - Richiesta Integrazioni" ||
                            attivitaOriginale === "Valutazione tecnica" ||
                            attivitaOriginale === "Verifica energetica ambientale") {
                            if (!mappedRow["Data attivazione"]) mappedRow["Data attivazione"] = dataPura;
                        }
                        else if (attivitaOriginale === "Attesa integrazione documentale") {
                            if (!mappedRow["Data invio ordinanza"]) mappedRow["Data invio ordinanza"] = dataPura;
                        }

                        fogliDestinazione[targetFoglio].push(mappedRow);

                        // Ferma tutto. Ha trovato casa (es. se trova VEA non va a farsi scansionare in SCIA)
                        break;
                    }
                }

                // GESTIONE CASI FUORI REGOLA: se la pratica non soddisfa NULLA, va in "ALTRO"
                if (!hasMatchedRule) {
                    if (!fogliDestinazione["ALTRO"]) fogliDestinazione["ALTRO"] = [];
                    let mappedRow = {};
                    for (let colFinale in campiImportConfig) {
                        let colSorgente = campiImportConfig[colFinale];
                        mappedRow[colFinale] = rowObj[colSorgente] || "";
                    }

                    // --- REGOLE CONDIZIONALI SULLE DATE ---
                    const attivitaOriginale = (rowObj["Attività da espletare"] || "").trim();
                    const dataPura = (rowObj["Data attivazione"] || "").trim();

                    if (attivitaOriginale === "Istruttoria Tecnica - Richiesta Integrazioni" ||
                        attivitaOriginale === "Valutazione tecnica" ||
                        attivitaOriginale === "Verifica energetica ambientale") {
                        if (!mappedRow["Data attivazione"]) mappedRow["Data attivazione"] = dataPura;
                    }
                    else if (attivitaOriginale === "Attesa integrazione documentali") {
                        if (!mappedRow["Data invio ordinanza"]) mappedRow["Data invio ordinanza"] = dataPura;
                    }

                    fogliDestinazione["ALTRO"].push(mappedRow);
                }
            }

            // --- FASE C: CONTROLLO DUPLICATI E INSERIMENTO INCROCIATO NEI FOGLI STORICI ---
            let uniqueNewPracticesReport = new Map(); // Mappa Code => Date per il report

            // Poiché una singola pratica (es: 9-011) può avere più attività nel corso del tempo e vogliamo salvarle tutte,
            // la chiave primaria di deduplicazione NON è solo il "Codice pratica", ma è COMPOSTA.
            let pkCode = "Codice pratica";
            let pkTask = "Stato"; // "Stato" è il nome che diamo alla colonna di ricezione per "Attività da espletare"

            for (let sheetName in fogliDestinazione) {
                let destSheet = sheets.getItemOrNullObject(sheetName);
                await context.sync();

                if (destSheet.isNullObject) {
                    destSheet = sheets.add(sheetName);
                    let defaultHeaders = Object.keys(campiImportConfig);
                    let headerRange = destSheet.getCell(0, 0).getResizedRange(0, defaultHeaders.length - 1);
                    headerRange.values = [defaultHeaders];
                    await context.sync();
                }

                let usedRange = destSheet.getUsedRangeOrNullObject();
                usedRange.load("values");
                await context.sync();

                let existingComposites = new Set();
                let lastRowIndex = 0;
                let foglioValues = [];
                let sheetHeaders = [];

                if (!usedRange.isNullObject && usedRange.values && usedRange.values.length > 0) {
                    foglioValues = usedRange.values;
                    lastRowIndex = foglioValues.length;
                    sheetHeaders = foglioValues[0];

                    let idxCode = sheetHeaders.indexOf(pkCode);
                    let idxTask = sheetHeaders.indexOf(pkTask);
                    if (idxCode === -1) idxCode = 0;
                    if (idxTask === -1) idxTask = 0; // Fallback

                    for (let i = 1; i < foglioValues.length; i++) {
                        let valCode = foglioValues[i][idxCode] ? String(foglioValues[i][idxCode]).trim() : "";
                        let valTask = foglioValues[i][idxTask] ? String(foglioValues[i][idxTask]).trim() : "";
                        existingComposites.add(`${valCode}###${valTask}`);
                    }
                } else {
                    sheetHeaders = Object.keys(campiImportConfig);
                    let headerRange = destSheet.getCell(0, 0).getResizedRange(0, sheetHeaders.length - 1);
                    headerRange.values = [sheetHeaders];
                    lastRowIndex = 1;
                }

                let rowsToAppend = [];
                let ramBlockList = new Set();

                for (let mappedObj of fogliDestinazione[sheetName]) {
                    let valCode = String(mappedObj[pkCode] || "").trim();
                    let valTask = String(mappedObj[pkTask] || "").trim();
                    let compositeKey = `${valCode}###${valTask}`;

                    if (!existingComposites.has(compositeKey) && !ramBlockList.has(compositeKey)) {
                        ramBlockList.add(compositeKey);
                        existingComposites.add(compositeKey);

                        let flatRow = sheetHeaders.map(col => forceISODates(mappedObj[col] || ""));
                        rowsToAppend.push(flatRow);

                        // Per il report usiamo comunque il codice pratica per non affollare la lista
                        if (!uniqueNewPracticesReport.has(valCode)) {
                            uniqueNewPracticesReport.set(valCode, mappedObj["Data Praedi"] || "N.D.");
                        }
                    }
                }

                // Spariamo le nuove righe allineate sul foglio!
                if (rowsToAppend.length > 0) {
                    let insertRange = destSheet.getCell(lastRowIndex, 0).getResizedRange(rowsToAppend.length - 1, sheetHeaders.length - 1);

                    // Applichiamo il trucco del Text Format ("", "@", ecc) SOLO al Codice Pratica
                    let idxCodeStorico = sheetHeaders.indexOf(pkCode);
                    let formatsStorico = rowsToAppend.map(row => row.map((val, colIndex) => {
                        return (colIndex === idxCodeStorico) ? "@" : "General";
                    }));
                    
                    insertRange.numberFormat = formatsStorico;

                    insertRange.values = rowsToAppend;
                    insertRange.format.autofitColumns();
                    await context.sync();
                }
            }

            showStatus("Pratiche smistate con successo in " + Object.keys(fogliDestinazione).length + " fogli!", "success");

            // FASE D: MOSTRA IL SUPER REPORT FINALE
            showReportModal(uniqueNewPracticesReport, csvObjects.length);

        });
    } catch (error) {
        let errorMsg = "Errore durante l'importazione";
        if (error instanceof Error) {
            errorMsg += ": " + error.message;
        } else if (typeof error === "string") {
            errorMsg += ": " + error;
        }
        showStatus(errorMsg, "error");
        console.error(error);
        document.getElementById('importBtn').disabled = false;
    }
}

// Funzione Helper UI per il Report
function showReportModal(newPracticesMap, totalScrivania) {
    document.getElementById('statTotalCount').textContent = totalScrivania;
    document.getElementById('statNewCount').textContent = newPracticesMap.size;

    let ul = document.getElementById('reportList');
    ul.innerHTML = ''; // Svuota

    if (newPracticesMap.size === 0) {
        let li = document.createElement('li');
        li.textContent = "Nessuna nuova pratica trovata. Erano tutte già storicizzate.";
        li.style.justifyContent = "center";
        ul.appendChild(li);
    } else {
        newPracticesMap.forEach((dataPraedi, codice) => {
            let li = document.createElement('li');

            let spanCode = document.createElement('strong');
            spanCode.textContent = codice;

            let spanData = document.createElement('span');
            spanData.textContent = dataPraedi;
            spanData.style.color = "var(--text-muted)";

            li.appendChild(spanCode);
            li.appendChild(spanData);
            ul.appendChild(li);
        });
    }

    document.getElementById('reportModal').style.display = 'flex';
}

function showStatus(text, typeClass) {
    const el = document.getElementById('statusMessage');
    el.textContent = text;
    el.className = 'status-message ' + typeClass;
}

function showMigStatus(text, typeClass) {
    const el = document.getElementById('migStatusMessage');
    el.textContent = text;
    el.className = 'status-message ' + typeClass;
}

// ════════════════════════════════════════════════════════════════════════
// SEZIONE MIGRAZIONE
// ════════════════════════════════════════════════════════════════════════

// Legge il file sorgente (CSV o XLSX) e prepara i dati in migSourceData
async function handleMigrationFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;

    const ext = file.name.split('.').pop().toLowerCase();
    showMigStatus('Lettura file in corso...', 'success');
    migDestColumnsCache = {};

    if (ext === 'csv') {
        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                const rawData = parseCSV(e.target.result);
                if (rawData.length < 1) { showMigStatus('File CSV vuoto.', 'error'); return; }
                const headers = rawData[0].map(h => String(h || '').trim());
                const rows = rawData.slice(1);
                migSourceData = { sheetNames: ['Foglio1'], sheets: { 'Foglio1': { headers, rows } } };
                
                document.getElementById('migFileName').textContent = file.name;
                document.getElementById('migFileRows').textContent = rows.length;
                document.getElementById('migFileInfo').style.display = 'block';
                await renderMappingTableMode();
            } catch (err) {
                showMigStatus('Errore CSV: ' + err.message, 'error');
            }
        };
        reader.onerror = () => showMigStatus('Errore lettura file CSV.', 'error');
        reader.readAsText(file, 'UTF-8');

    } else if (ext === 'xlsx' || ext === 'xls' || ext === 'ods') {
        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                migSourceData = { sheetNames: workbook.SheetNames, sheets: {} };
                let totalRows = 0;
                for (const sheetName of workbook.SheetNames) {
                    const ws = workbook.Sheets[sheetName];
                    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
                    const headers = (aoa.length > 0 ? aoa[0] : []).map(h => String(h || '').trim());
                    const rows = aoa.slice(1);
                    migSourceData.sheets[sheetName] = { headers, rows };
                    totalRows += rows.length;
                }
                document.getElementById('migFileName').textContent = file.name;
                document.getElementById('migFileRows').textContent = '~' + totalRows;
                document.getElementById('migFileInfo').style.display = 'block';
                await renderMappingTableMode();
            } catch (err) {
                showMigStatus('Errore XLSX: ' + err.message, 'error');
                console.error(err);
            }
        };
        reader.onerror = () => showMigStatus('Errore lettura file XLSX.', 'error');
        reader.readAsArrayBuffer(file);
    } else {
        showMigStatus('Formato non supportato. Usa .csv, .xlsx o .ods', 'error');
    }
}

// Recupera l'elenco dei fogli del workbook Excel attivo
async function loadDestSheets() {
    try {
        return await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load('items/name');
            await context.sync();
            return sheets.items.map(s => s.name);
        });
    } catch (e) {
        console.warn('loadDestSheets error:', e);
        return [];
    }
}

// Recupera le intestazioni (riga 0) di un foglio di destinazione nel workbook attivo
async function loadDestColumns(sheetName) {
    if (migDestColumnsCache[sheetName]) return migDestColumnsCache[sheetName];
    try {
        const cols = await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            await context.sync();
            if (sheet.isNullObject) return [];
            const used = sheet.getUsedRangeOrNullObject();
            used.load('values');
            await context.sync();
            if (used.isNullObject || !used.values || used.values.length === 0) return [];
            return used.values[0].map(h => String(h || '')).filter(h => h.trim() !== '');
        });
        migDestColumnsCache[sheetName] = cols;
        return cols;
    } catch (e) {
        console.warn('loadDestColumns error for', sheetName, e);
        return [];
    }
}

// Inizializza la UI di mapping
async function renderMappingTableMode() {
    showMigStatus('Caricamento fogli di destinazione...', 'success');
    migDestSheetsGlobal = await loadDestSheets();
    
    document.getElementById('migMappingBody').innerHTML = '';
    
    // Auto popolare le righe per le colonne del primo foglio sorgente per comodità
    const firstSheetName = migSourceData.sheetNames[0];
    const srcCols = migSourceData.sheets[firstSheetName].headers;
    for (const col of srcCols) {
        if (col && col.trim() !== '') {
            addMappingRow(firstSheetName, col);
        }
    }
    
    document.getElementById('migMappingSection').style.display = 'block';
    showMigStatus('File letto. Costruisci le regole di mapping e clicca Esegui.', 'success');
}
// Aggiorna le spunte ✅ sulle colonne sorgente già utilizzate per facilitare l'UX
function refreshOptionSpuntas() {
    const allRows = document.querySelectorAll('#migMappingBody tr');
    const selectedColsBySheet = {};
    
    // 1. Raccoglie tutto quello che è attualmente selezionato
    allRows.forEach(row => {
        const sheetSel = row.querySelector('.src-sheet-select');
        const colSel = row.querySelector('.src-col-select');
        if (sheetSel && colSel && colSel.value) {
            const sn = sheetSel.value;
            if(!selectedColsBySheet[sn]) selectedColsBySheet[sn] = new Set();
            selectedColsBySheet[sn].add(colSel.value);
        }
    });

    // 2. Aggiorna il testo visualizzato
    allRows.forEach(row => {
        const sheetSel = row.querySelector('.src-sheet-select');
        const colSel = row.querySelector('.src-col-select');
        if (sheetSel && colSel) {
            const sn = sheetSel.value;
            const selectedSet = selectedColsBySheet[sn] || new Set();
            
            Array.from(colSel.options).forEach(opt => {
                const val = opt.value;
                if (!val) return;
                // opt.value resta pulito, alteriamo solo l'involcro visivo (textContent)
                opt.textContent = selectedSet.has(val) ? `✅ ${val}` : val;
            });
        }
    });
}

// Aggiunge dinamicamente una riga/regola alla UI
function addMappingRow(defaultSrcSheet, defaultSrcCol) {
    const tbody = document.getElementById('migMappingBody');
    const tr = document.createElement('tr');
    
    // 1. Foglio Sorgente
    const tdSrcSheet = document.createElement('td');
    const srcSheetSel = document.createElement('select');
    srcSheetSel.className = 'src-sheet-select fluent-select';
    for(const sn of migSourceData.sheetNames) {
        const opt = document.createElement('option');
        opt.value = sn; opt.textContent = sn;
        srcSheetSel.appendChild(opt);
    }
    srcSheetSel.value = defaultSrcSheet || migSourceData.sheetNames[0];
    tdSrcSheet.appendChild(srcSheetSel);
    
    // 5. Azioni (Rimozione) - Spostato all'inizio!
    const tdActions = document.createElement('td');
    const removeBtn = document.createElement('button');
    removeBtn.className = 'action-btn';
    removeBtn.textContent = '❌';
    removeBtn.title = "Rimuovi regola";
    removeBtn.addEventListener('click', () => {
        tr.remove();
        refreshOptionSpuntas();
    });
    tdActions.appendChild(removeBtn);
    tr.appendChild(tdActions);

    tr.appendChild(tdSrcSheet);
    
    // 2. Colonna Sorgente
    const tdSrcCol = document.createElement('td');
    const srcColSel = document.createElement('select');
    srcColSel.className = 'src-col-select fluent-select';
    tdSrcCol.appendChild(srcColSel);
    tr.appendChild(tdSrcCol);
    
    const updateSrcCols = (sheetName, selectedCol) => {
        srcColSel.innerHTML = '';
        const headers = migSourceData.sheets[sheetName].headers;
        for(const h of headers) {
            if(!h || h.trim() === '') continue;
            const opt = document.createElement('option');
            opt.value = h; opt.textContent = h;
            srcColSel.appendChild(opt);
        }
        if(selectedCol && headers.includes(selectedCol)) {
            srcColSel.value = selectedCol;
        }
        refreshOptionSpuntas();
    };
    updateSrcCols(srcSheetSel.value, defaultSrcCol);
    
    // Listener: se cambia il foglio sorgente descritto, aggiorna i campi colonna
    srcSheetSel.addEventListener('change', () => {
        updateSrcCols(srcSheetSel.value);
    });

    // 3. Foglio Destinazione
    const tdDestSheet = document.createElement('td');
    const destSheetSel = document.createElement('select');
    destSheetSel.className = 'dest-sheet-select fluent-select';
    const noOpt = document.createElement('option');
    noOpt.value = ""; noOpt.textContent = "— Seleziona —";
    destSheetSel.appendChild(noOpt);
    for(const ds of migDestSheetsGlobal) {
        const opt = document.createElement('option');
        opt.value = ds; opt.textContent = ds;
        destSheetSel.appendChild(opt);
    }
    tdDestSheet.appendChild(destSheetSel);
    tr.appendChild(tdDestSheet);
    
    // 4. Colonna Destinazione
    const tdDestCol = document.createElement('td');
    const destColSel = document.createElement('select');
    destColSel.className = 'dest-col-select fluent-select';
    destColSel.disabled = true;
    destColSel.innerHTML = '<option value="__new__">＋ Nuova colonna</option>';
    tdDestCol.appendChild(destColSel);
    
    const newColInput = document.createElement('input');
    newColInput.type = 'text';
    newColInput.className = 'new-col-input';
    newColInput.placeholder = 'Nome nuova colonna';
    newColInput.style.display = 'none';
    
    newColInput.value = srcColSel.value || "Nuova Colonna";
    srcColSel.addEventListener('change', () => {
        if(destColSel.value === '__new__') newColInput.value = srcColSel.value;
        refreshOptionSpuntas();
    });

    tdDestCol.appendChild(newColInput);
    tr.appendChild(tdDestCol);
    
    // Listener: caricamento colonne della destinazione scelta
    destSheetSel.addEventListener('change', async () => {
        const chosen = destSheetSel.value;
        if (!chosen) {
            destColSel.disabled = true;
            destColSel.innerHTML = '<option value="__new__">＋ Nuova colonna</option>';
            newColInput.style.display = 'none';
            return;
        }
        destColSel.disabled = false;
        destColSel.innerHTML = '<option value="__new__">＋ Nuova colonna</option>';
        const destCols = await loadDestColumns(chosen);
        for (const col of destCols) {
            const opt = document.createElement('option');
            opt.value = col; opt.textContent = col;
            destColSel.appendChild(opt);
        }
        
        // Match guidato se esiste già una colonna con questo nome
        const srcVal = srcColSel.value;
        const match = destCols.find(c => c.toLowerCase() === srcVal.toLowerCase());
        if (match) {
            destColSel.value = match;
            newColInput.style.display = 'none';
        } else {
            destColSel.value = '__new__';
            newColInput.value = srcVal;
            newColInput.style.display = 'block';
        }
    });

    destColSel.addEventListener('change', () => {
        newColInput.style.display = destColSel.value === '__new__' ? 'block' : 'none';
        if (destColSel.value === '__new__') newColInput.value = srcColSel.value;
    });
    
    tbody.appendChild(tr);
}

// Esegue la migrazione
async function executeMigration() {
    if (!migSourceData) return;

    // 1. Raccogliamo e organizziamo tutte le regole di mapping dalla UI
    const mappings = [];
    document.querySelectorAll('#migMappingBody tr').forEach(row => {
        const srcSheetSel = row.querySelector('.src-sheet-select');
        const srcColSel   = row.querySelector('.src-col-select');
        const destSheetSel = row.querySelector('.dest-sheet-select');
        const destColSel   = row.querySelector('.dest-col-select');
        const newInput = row.querySelector('.new-col-input');
        
        const destSheet = destSheetSel.value;
        if (!destSheet) return; // ignorata (non mapped)
        
        const srcSheet = srcSheetSel.value;
        const srcCol = srcColSel.value;
        const isNew = destColSel.value === '__new__';
        const destCol = isNew ? null : destColSel.value;
        const newColName = (newInput && newInput.value.trim()) ? newInput.value.trim() : srcCol;
        
        mappings.push({ srcSheet, srcCol, destSheet, destCol, newColName, isNew });
    });

    if (mappings.length === 0) {
        showMigStatus('Nessuna colonna associata. Mappa almeno una destinazione.', 'error');
        return;
    }

    // 2. Raggruppa per foglio di destinazione
    const bySheet = {};
    for (const m of mappings) {
        if (!bySheet[m.destSheet]) bySheet[m.destSheet] = [];
        bySheet[m.destSheet].push(m);
    }

    // 3. Validiamo: incroci multipli di fogli sorgente verso UN singolo foglio dest NON sono consentiti.
    for (const [destSheet, maps] of Object.entries(bySheet)) {
        const srcSheets = new Set(maps.map(m => m.srcSheet));
        if (srcSheets.size > 1) {
            showMigStatus(`Errore: Il foglio dest '${destSheet}' riceve dati da più fogli sorgenti. Usa un foglio Excel di destinazione univoco per ogni foglio sorgente.`, 'error');
            return;
        }
    }

    document.getElementById('migExecuteBtn').disabled = true;
    showMigStatus('Migrazione in corso...', 'success');

    // Pre-popoliamo la cache per evitare overwrite di colonne pre-esistenti!
    for (const sheetName of Object.keys(bySheet)) {
        await loadDestColumns(sheetName);
    }

    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;

            // Processa ogni foglio di destinazione
            for (const [sheetName, sheetMappings] of Object.entries(bySheet)) {
                let destSheet = sheets.getItemOrNullObject(sheetName);
                await context.sync();
                if (destSheet.isNullObject) {
                    destSheet = sheets.add(sheetName);
                    migDestColumnsCache[sheetName] = []; // brand new
                }

                // Conserviamo l'intestazione esistente E i dati esistenti!
                let existingValues = [];
                const destUsedRange = destSheet.getUsedRangeOrNullObject();
                destUsedRange.load('values');
                await context.sync();

                if (!destUsedRange.isNullObject && destUsedRange.values && destUsedRange.values.length > 0) {
                    existingValues = destUsedRange.values;
                }

                const existingHeaders = (existingValues.length > 0) ? existingValues[0].map(h => String(h || '')) : (migDestColumnsCache[sheetName] || []);

                // Foglio sorgente per questo foglio destinazione
                const srcSheetName = sheetMappings[0].srcSheet;
                const srcSheet = migSourceData.sheets[srcSheetName];
                const srcHeaders = srcSheet.headers;
                const srcRows    = srcSheet.rows;

                // L'ordine finale delle colonne D: prima quelle già esistenti, poi le nuove agganciate in coda
                const destColumns = [...existingHeaders];
                for (const m of sheetMappings) {
                    const colName = m.isNew ? m.newColName : m.destCol;
                    if (!destColumns.includes(colName)) destColumns.push(colName);
                }

                // Costruisce il piano di mapping index: srcColIdx -> destColIdx
                const colPlan = [];
                for (const m of sheetMappings) {
                    const srcIdx  = srcHeaders.indexOf(m.srcCol);
                    const dstName = m.isNew ? m.newColName : m.destCol;
                    const dstIdx  = destColumns.indexOf(dstName);
                    if (srcIdx !== -1 && dstIdx !== -1) colPlan.push({ srcIdx, dstIdx });
                }

                const totalCols = destColumns.length;
                if (totalCols === 0) continue;

                // Mantiene il formattato, svuotando solo prima del replace per evitare sovrapposizioni sporche
                destSheet.getRange().clear("Contents");
                await context.sync();

                const maxRows = Math.max(existingValues.length > 1 ? existingValues.length - 1 : 0, srcRows.length);
                const destValuesMatrix = [];
                
                // Intestazioni
                destValuesMatrix.push(destColumns);

                for (let r = 0; r < maxRows; r++) {
                    const destRow = new Array(totalCols).fill('');
                    
                    // 1. Ripristina i dati esistenti nel foglio per questa riga
                    if (existingValues.length > r + 1) {
                        const oldRow = existingValues[r + 1];
                        for (let c = 0; c < oldRow.length; c++) {
                            destRow[c] = oldRow[c] ?? '';
                        }
                    }

                    // 2. Sovrascrivi dal file di mapping SOLO le colonne interessate
                    if (r < srcRows.length) {
                        const srcRow = srcRows[r];
                        for (const { srcIdx, dstIdx } of colPlan) {
                            const raw = Array.isArray(srcRow) ? (srcRow[srcIdx] ?? '') : '';
                            destRow[dstIdx] = forceISODates(String(raw));
                        }
                    }
                    destValuesMatrix.push(destRow);
                }

                const dataRange = destSheet.getCell(0, 0).getResizedRange(destValuesMatrix.length - 1, totalCols - 1);

                // Formatta tutto come testo ("@") tranne le colonne "Data" o "Date"
                const formatsMatrix = destValuesMatrix.map(() => 
                    destColumns.map(col => {
                        const lowerCol = String(col).toLowerCase();
                        return (lowerCol.includes("data") || lowerCol.includes("date")) ? "General" : "@";
                    })
                );
                dataRange.numberFormat = formatsMatrix;

                dataRange.values = destValuesMatrix;
                
                destSheet.getUsedRange().format.autofitColumns();
                await context.sync();
            }
        });

        const nSheets = Object.keys(bySheet).length;
        showMigStatus(
            `✅ Migrazione completata: ${mappings.length} regole elaborate su ${nSheets} foglio/i.`,
            'success'
        );
        migDestColumnsCache = {}; // invalida cache dopo scrittura per eventuali run successivi
    } catch (error) {
        let msg = 'Errore durante la migrazione';
        if (error instanceof Error) msg += ': ' + error.message;
        showMigStatus(msg, 'error');
        console.error(error);
    } finally {
        document.getElementById('migExecuteBtn').disabled = false;
    }
}
