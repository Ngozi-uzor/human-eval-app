// Application State
let flatData = [];
let groupedData = []; // Array of { questionId, promptType, condition, questionText, rows: [{rowIndex, rowData}] }
let currentGroupIndex = 0;
let evaluations = {}; // Key: rowIndex, Value: { ef, rr, lq, pp, errors:[], comment }
let headers = [];
let annotatorId = "";

const ERROR_TYPES = [
    "Hallucination", "Shallow Description", "Model Breakdown", "Repetition",
    "Western Default", "Nonsensical Response", "Empty Response", "Incomplete Response",
    "Culturally Accurate", "Pan-African Generalisation", "Ethnolinguistic Substitution",
    "Wrong Language", "Contradicts Prompt"
];

// DOM Elements
const screens = {
    onboarding: document.getElementById('onboarding-screen'),
    dashboard: document.getElementById('dashboard-screen'),
    eval: document.getElementById('eval-screen')
};

const ui = {
    inputId: document.getElementById('annotator-id'),
    btnStart: document.getElementById('btn-start-onboarding'),
    loading: document.getElementById('loading-overlay'),
    
    // Dashboard
    dashTotal: document.getElementById('dash-total'),
    dashCompleted: document.getElementById('dash-completed'),
    dashRemaining: document.getElementById('dash-remaining'),
    qList: document.getElementById('question-list-container'),
    btnExportDash: document.getElementById('btn-export-dash'),
    
    // Eval
    btnBackDash: document.getElementById('btn-back-dash'),
    currIdx: document.getElementById('current-item-idx'),
    totalItems: document.getElementById('total-items'),
    promptType: document.getElementById('prompt-type-display'),
    condition: document.getElementById('condition-display'),
    questionText: document.getElementById('question-display'),
    modelsContainer: document.getElementById('models-container'),
    btnPrev: document.getElementById('btn-prev-q'),
    btnSave: document.getElementById('btn-save-q'),
    btnNext: document.getElementById('btn-next-q')
};

// Initialize
function init() {
    showScreen('onboarding');
}

function showScreen(screenName) {
    Object.values(screens).forEach(s => s.classList.add('hidden-screen'));
    screens[screenName].classList.remove('hidden-screen');
}

// Onboarding & Loading
ui.btnStart.addEventListener('click', async () => {
    const val = ui.inputId.value.trim();
    if (!val) {
        alert("Please enter an Annotator ID.");
        return;
    }
    annotatorId = val;
    
    // Load local storage
    const saved = localStorage.getItem(`cabIgboEval_${annotatorId}`);
    if (saved) {
        evaluations = JSON.parse(saved);
    } else {
        evaluations = {};
    }

    ui.loading.classList.remove('hidden-modal');
    try {
        await fetchAndParseDataset();
        renderDashboard();
        showScreen('dashboard');
    } catch (err) {
        alert("Error loading dataset: " + err.message);
    } finally {
        ui.loading.classList.add('hidden-modal');
    }
});

async function fetchAndParseDataset() {
    const response = await fetch('dataset.xlsx');
    if (!response.ok) throw new Error("Could not find dataset.xlsx in the directory.");
    
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(data, {type: 'array'});
    
    // Process each sheet separately to avoid mixing types (EN-EN, EN-IG, IG-IG)
    const allSheetData = []; // [{sheetName, rows, headers}]

    workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const sheetRows = XLSX.utils.sheet_to_json(worksheet, {defval: ""});
        
        if (sheetRows.length > 0) {
            const sheetHeaders = Object.keys(sheetRows[0]);
            // Only include actual annotation sheets (not instruction/decoder sheets)
            const hasQuestionCol = sheetHeaders.some(h => h.toLowerCase().includes('question'));
            const hasModelCol = sheetHeaders.some(h => h.toLowerCase().includes('model'));
            if (hasQuestionCol && hasModelCol) {
                allSheetData.push({ sheetName, rows: sheetRows, headers: sheetHeaders });
                if (!headers.length) headers = sheetHeaders;
            }
        }
    });

    if (allSheetData.length === 0) throw new Error("Could not find any annotation sheets in the dataset.");

    // Helper to find a column value using keyword matching
    const getColFrom = (possibleKeys, row, hdrs) => {
        const key = hdrs.find(h => possibleKeys.some(k => h.toLowerCase().includes(k.toLowerCase())));
        return key ? String(row[key]) : "";
    };

    // Build flat data (all sheets combined, tagged with sheet name)
    flatData = [];
    allSheetData.forEach(sheet => {
        sheet.rows.forEach(row => {
            flatData.push({ ...row, __sheet: sheet.sheetName });
        });
    });

    // Group rows per sheet, per Question Number
    const groupsMap = new Map();

    allSheetData.forEach(sheet => {
        const hdrs = sheet.headers;

        sheet.rows.forEach((row, localIdx) => {
            const globalIdx = flatData.findIndex(r => r === row || (r.__sheet === sheet.sheetName && Object.keys(row).every(k => r[k] === row[k])));

            const qNum  = getColFrom(['question number', 'question no', 'q_num', 'qnum'], row, hdrs);
            const qText = getColFrom(['question text', 'question', 'prompt text', 'input'], row, hdrs);
            const pType = getColFrom(['prompt type'], row, hdrs);
            const domain = getColFrom(['domain'], row, hdrs);

            if (!qText) return;

            // Unique key = sheet + question number (prevents cross-type mixing)
            const groupKey = `${sheet.sheetName}||${qNum || qText}`;

            if (!groupsMap.has(groupKey)) {
                groupsMap.set(groupKey, {
                    questionId: groupsMap.size + 1,
                    promptType: pType,
                    sheetName: sheet.sheetName,
                    domain: domain,
                    questionText: qText,
                    rows: []
                });
            }
            // Push using actual index in flatData
            const flatIdx = flatData.findIndex((r, i) => r.__sheet === sheet.sheetName && r === flatData.filter(x => x.__sheet === sheet.sheetName)[localIdx]);
            groupsMap.get(groupKey).rows.push({ rowIndex: flatIdx !== -1 ? flatIdx : flatData.length - sheet.rows.length + localIdx, rowData: row, hdrs });
        });
    });

    groupedData = Array.from(groupsMap.values());
}

// Dashboard
function renderDashboard() {
    let completedCount = 0;
    ui.qList.innerHTML = '';

    groupedData.forEach((group, i) => {
        // Check if all rows in this group are evaluated
        const isCompleted = group.rows.every(r => {
            const ev = evaluations[r.rowIndex];
            return ev && ev.ef !== '' && ev.rr !== '' && ev.lq !== '' && ev.pp !== '';
        });
        
        if (isCompleted) completedCount++;

        const card = document.createElement('div');
        card.className = `q-card ${isCompleted ? 'completed' : ''}`;
        card.innerHTML = `
            <h4>Question ${i + 1}</h4>
            <p>${group.questionText}</p>
        `;
        card.addEventListener('click', () => {
            currentGroupIndex = i;
            renderEvaluationScreen();
        });
        ui.qList.appendChild(card);
    });

    ui.dashTotal.innerText = groupedData.length;
    ui.dashCompleted.innerText = completedCount;
    ui.dashRemaining.innerText = groupedData.length - completedCount;
}

// Evaluation Screen
function renderEvaluationScreen() {
    showScreen('eval');
    const group = groupedData[currentGroupIndex];
    
    ui.currIdx.innerText = currentGroupIndex + 1;
    ui.totalItems.innerText = groupedData.length;
    ui.promptType.innerText = group.promptType || "-";
    ui.condition.innerText = group.condition || "-";
    ui.questionText.innerText = group.questionText;

    ui.modelsContainer.innerHTML = '';

    const getCol = (possibleKeys, row) => {
        const key = headers.find(h => possibleKeys.some(k => h.toLowerCase().includes(k.toLowerCase())));
        return key ? row[key] : "No data";
    };

    group.rows.forEach((r, idx) => {
        const hdrs = r.hdrs || headers;
        const getColLocal = (keys, row) => {
            const key = hdrs.find(h => keys.some(k => h.toLowerCase().includes(k.toLowerCase())));
            return key ? String(row[key]) : '';
        };
        const modelName = getColLocal(['model code', 'model name', 'model'], r.rowData) || `Model ${String.fromCharCode(65 + idx)}`;
        const responseText = getColLocal(['model answer', 'answer', 'response', 'generation', 'output'], r.rowData);
        const ev = evaluations[r.rowIndex] || { ef:'', rr:'', lq:'', pp:'', errors:[], comment:'' };

        const card = document.createElement('div');
        card.className = 'model-card';
        // Build checkboxes HTML
        const checkboxesHtml = ERROR_TYPES.map(err => {
            const isChecked = ev.errors.includes(err) ? 'checked' : '';
            return `
                <label class="checkbox-item">
                    <input type="checkbox" name="err_${r.rowIndex}" value="${err}" ${isChecked}>
                    ${err}
                </label>
            `;
        }).join('');

        card.innerHTML = `
            <h4>${modelName}</h4>
            <div class="scroll-box">${responseText}</div>
            
            <div class="evaluation-form">
                <div class="eval-grid">
                    <div class="metric-group">
                        <label>Epistemic Fidelity (EF)</label>
                        <p class="metric-desc">0=Wrong, 1=Partial, 2=Fully correct</p>
                        <div class="radio-group">
                            <input type="radio" name="ef_${r.rowIndex}" id="ef0_${r.rowIndex}" value="0" ${ev.ef==='0'?'checked':''}><label for="ef0_${r.rowIndex}">0</label>
                            <input type="radio" name="ef_${r.rowIndex}" id="ef1_${r.rowIndex}" value="1" ${ev.ef==='1'?'checked':''}><label for="ef1_${r.rowIndex}">1</label>
                            <input type="radio" name="ef_${r.rowIndex}" id="ef2_${r.rowIndex}" value="2" ${ev.ef==='2'?'checked':''}><label for="ef2_${r.rowIndex}">2</label>
                        </div>
                    </div>

                    <div class="metric-group">
                        <label>Representational Richness (RR)</label>
                        <p class="metric-desc">0=Empty, 1=Shallow, 2=Deep</p>
                        <div class="radio-group">
                            <input type="radio" name="rr_${r.rowIndex}" id="rr0_${r.rowIndex}" value="0" ${ev.rr==='0'?'checked':''}><label for="rr0_${r.rowIndex}">0</label>
                            <input type="radio" name="rr_${r.rowIndex}" id="rr1_${r.rowIndex}" value="1" ${ev.rr==='1'?'checked':''}><label for="rr1_${r.rowIndex}">1</label>
                            <input type="radio" name="rr_${r.rowIndex}" id="rr2_${r.rowIndex}" value="2" ${ev.rr==='2'?'checked':''}><label for="rr2_${r.rowIndex}">2</label>
                        </div>
                    </div>

                    <div class="metric-group">
                        <label>Linguistic Quality (LQ)</label>
                        <p class="metric-desc">0=Unreadable, 1=Fluent w/ errors, 2=Perfect</p>
                        <div class="radio-group">
                            <input type="radio" name="lq_${r.rowIndex}" id="lq0_${r.rowIndex}" value="0" ${ev.lq==='0'?'checked':''}><label for="lq0_${r.rowIndex}">0</label>
                            <input type="radio" name="lq_${r.rowIndex}" id="lq1_${r.rowIndex}" value="1" ${ev.lq==='1'?'checked':''}><label for="lq1_${r.rowIndex}">1</label>
                            <input type="radio" name="lq_${r.rowIndex}" id="lq2_${r.rowIndex}" value="2" ${ev.lq==='2'?'checked':''}><label for="lq2_${r.rowIndex}">2</label>
                        </div>
                    </div>

                    <div class="metric-group">
                        <label>Pragmatic Proficiency (PP)</label>
                        <p class="metric-desc">0=Off-topic, 1=Close, 2=Perfect fit</p>
                        <div class="radio-group">
                            <input type="radio" name="pp_${r.rowIndex}" id="pp0_${r.rowIndex}" value="0" ${ev.pp==='0'?'checked':''}><label for="pp0_${r.rowIndex}">0</label>
                            <input type="radio" name="pp_${r.rowIndex}" id="pp1_${r.rowIndex}" value="1" ${ev.pp==='1'?'checked':''}><label for="pp1_${r.rowIndex}">1</label>
                            <input type="radio" name="pp_${r.rowIndex}" id="pp2_${r.rowIndex}" value="2" ${ev.pp==='2'?'checked':''}><label for="pp2_${r.rowIndex}">2</label>
                        </div>
                    </div>
                </div>

                <div class="metric-group" style="margin-top: 1.5rem;">
                    <label>Error Taxonomy (Select all that apply)</label>
                    <div class="taxonomy-grid">
                        ${checkboxesHtml}
                    </div>
                </div>

                <div class="metric-group">
                    <label>Comments</label>
                    <textarea id="comment_${r.rowIndex}" rows="2" placeholder="Any notes...">${ev.comment}</textarea>
                </div>
            </div>
        `;
        ui.modelsContainer.appendChild(card);
    });

    ui.btnPrev.disabled = currentGroupIndex === 0;
    ui.btnNext.innerText = currentGroupIndex === groupedData.length - 1 ? "Finish" : "Next Question";
}

function saveCurrentState() {
    const group = groupedData[currentGroupIndex];
    
    group.rows.forEach(r => {
        const getRadio = (name) => {
            const rd = document.querySelector(`input[name="${name}_${r.rowIndex}"]:checked`);
            return rd ? rd.value : '';
        };

        const errorChecks = document.querySelectorAll(`input[name="err_${r.rowIndex}"]:checked`);
        const errors = Array.from(errorChecks).map(cb => cb.value);

        const comment = document.getElementById(`comment_${r.rowIndex}`).value;

        evaluations[r.rowIndex] = {
            ef: getRadio('ef'),
            rr: getRadio('rr'),
            lq: getRadio('lq'),
            pp: getRadio('pp'),
            errors: errors,
            comment: comment
        };
    });

    localStorage.setItem(`cabIgboEval_${annotatorId}`, JSON.stringify(evaluations));
}

ui.btnNext.addEventListener('click', () => {
    saveCurrentState();
    if (currentGroupIndex < groupedData.length - 1) {
        currentGroupIndex++;
        renderEvaluationScreen();
        window.scrollTo(0,0);
    } else {
        renderDashboard();
        showScreen('dashboard');
    }
});

ui.btnPrev.addEventListener('click', () => {
    saveCurrentState();
    if (currentGroupIndex > 0) {
        currentGroupIndex--;
        renderEvaluationScreen();
        window.scrollTo(0,0);
    }
});

ui.btnSave.addEventListener('click', () => {
    saveCurrentState();
    renderDashboard();
    showScreen('dashboard');
});

ui.btnBackDash.addEventListener('click', () => {
    saveCurrentState();
    renderDashboard();
    showScreen('dashboard');
});

// Export
ui.btnExportDash.addEventListener('click', () => {
    const exportData = flatData.map((row, idx) => {
        const ev = evaluations[idx] || { ef:'', rr:'', lq:'', pp:'', errors:[], comment:'' };
        return {
            ...row,
            Annotator_ID: annotatorId,
            EF_score: ev.ef,
            RR_score: ev.rr,
            LQ_score: ev.lq,
            PP_score: ev.pp,
            Error_Taxonomy: ev.errors.join(' | '),
            Comments: ev.comment
        };
    });

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Evaluations");
    
    const safeId = annotatorId.replace(/[^a-z0-9]/gi, '_').toLowerCase() || 'annotator';
    XLSX.writeFile(workbook, `CAB_Igbo_Evaluations_${safeId}.xlsx`);
});

// Modal Logic
const btnGuidelines = document.getElementById('btn-guidelines');
const modalGuidelines = document.getElementById('guidelines-modal');
const closeModal = document.getElementById('close-modal');

if (btnGuidelines) {
    btnGuidelines.addEventListener('click', () => {
        modalGuidelines.classList.remove('hidden-modal');
    });
}
if (closeModal) {
    closeModal.addEventListener('click', () => {
        modalGuidelines.classList.add('hidden-modal');
    });
}
window.addEventListener('click', (e) => {
    if (e.target === modalGuidelines) {
        modalGuidelines.classList.add('hidden-modal');
    }
});
