// ─── State ───────────────────────────────────────────────────────────
let flatData      = [];   // All rows (all sheets combined)
let groupedData   = [];   // All groups regardless of type
let filteredData  = [];   // Groups for the currently selected type
let currentGroupIndex = 0;
let evaluations   = {};   // {rowIndex: {ef, rr, lq, pp, errors, comment}}
let headers       = [];
let annotatorId   = "";
let selectedType  = "";   // e.g. "Type 1 Annotation"

const ERROR_TYPES = [
    "Hallucination", "Shallow Description", "Model Breakdown", "Repetition",
    "Western Default", "Nonsensical Response", "Empty Response", "Incomplete Response",
    "Culturally Accurate", "Pan-African Generalisation", "Ethnolinguistic Substitution",
    "Wrong Language", "Contradicts Prompt"
];

// ─── Screens ──────────────────────────────────────────────────────────
const screens = {
    login:     document.getElementById('screen-login'),
    type:      document.getElementById('screen-type'),
    dashboard: document.getElementById('screen-dashboard'),
    eval:      document.getElementById('screen-eval')
};

function showScreen(name) {
    Object.values(screens).forEach(s => s.classList.add('hidden-screen'));
    screens[name].classList.remove('hidden-screen');
}

// ─── UI refs ─────────────────────────────────────────────────────────
const ui = {
    inputId:        document.getElementById('annotator-id'),
    btnLogin:       document.getElementById('btn-login'),
    loading:        document.getElementById('loading-overlay'),
    typeGrid:       document.getElementById('type-grid'),
    btnLogout:      document.getElementById('btn-logout'),
    dashTitle:      document.getElementById('dash-type-title'),
    dashTotal:      document.getElementById('dash-total'),
    dashCompleted:  document.getElementById('dash-completed'),
    dashRemaining:  document.getElementById('dash-remaining'),
    qList:          document.getElementById('question-list-container'),
    btnExportDash:  document.getElementById('btn-export-dash'),
    btnBackType:    document.getElementById('btn-back-type'),
    btnBackDash:    document.getElementById('btn-back-dash'),
    currIdx:        document.getElementById('current-item-idx'),
    totalItems:     document.getElementById('total-items'),
    metaType:       document.getElementById('meta-type'),
    metaDomain:     document.getElementById('meta-domain'),
    metaQnum:       document.getElementById('meta-qnum'),
    questionText:   document.getElementById('question-display'),
    modelsContainer:document.getElementById('models-container'),
    btnPrev:        document.getElementById('btn-prev-q'),
    btnSave:        document.getElementById('btn-save-q'),
    btnNext:        document.getElementById('btn-next-q')
};

// ─── Allowed Annotator IDs (add more names here when needed) ─────────
const ALLOWED_IDS = ['ngozi'];

// ─── Login ────────────────────────────────────────────────────────────
ui.btnLogin.addEventListener('click', async () => {
    const val = ui.inputId.value.trim().toLowerCase();
    if (!val) { alert("Please enter your Annotator ID."); return; }

    if (!ALLOWED_IDS.includes(val)) {
        alert("❌ Unrecognised Annotator ID.\nPlease check your ID and try again.");
        return;
    }

    annotatorId = val;

    const saved = localStorage.getItem(`cabIgboEval_${annotatorId}`);
    evaluations = saved ? JSON.parse(saved) : {};

    ui.loading.classList.remove('hidden-modal');
    try {
        await fetchAndParseDataset();
        renderTypeSelection();
        showScreen('type');
    } catch (err) {
        alert("Error loading dataset: " + err.message + "\n\nMake sure dataset.xlsx is in the same folder as index.html.");
    } finally {
        ui.loading.classList.add('hidden-modal');
    }
});

ui.btnLogout.addEventListener('click', () => {
    annotatorId = "";
    selectedType = "";
    ui.inputId.value = "";
    showScreen('login');
});

// ─── Dataset Loading ──────────────────────────────────────────────────
async function fetchAndParseDataset() {
    const response = await fetch('dataset.xlsx');
    if (!response.ok) throw new Error("dataset.xlsx not found.");

    const ab = await response.arrayBuffer();
    const workbook = XLSX.read(new Uint8Array(ab), { type: 'array' });

    const allSheetData = [];

    workbook.SheetNames.forEach(sheetName => {
        const ws = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (rows.length > 0) {
            const hdrs = Object.keys(rows[0]);
            const hasQ = hdrs.some(h => h.toLowerCase().includes('question'));
            const hasM = hdrs.some(h => h.toLowerCase().includes('model'));
            if (hasQ && hasM) {
                allSheetData.push({ sheetName, rows, hdrs });
                if (!headers.length) headers = hdrs;
            }
        }
    });

    if (!allSheetData.length) throw new Error("No annotation sheets found.");

    // Build flatData with sheet tag
    flatData = [];
    allSheetData.forEach(sheet => {
        sheet.rows.forEach(row => flatData.push({ ...row, __sheet: sheet.sheetName }));
    });

    // Helper
    const getCol = (keys, row, hdrs) => {
        const k = hdrs.find(h => keys.some(kk => h.toLowerCase().includes(kk.toLowerCase())));
        return k ? String(row[k]).trim() : "";
    };

    // Build groups: one group per (sheet × question number)
    const groupsMap = new Map();
    let flatIdx = 0;

    allSheetData.forEach(sheet => {
        sheet.rows.forEach((row, li) => {
            const qNum  = getCol(['question number', 'question no', 'qnum', 'q_num', 'q num'], row, sheet.hdrs);
            const qText = getCol(['question text', 'question', 'input', 'prompt text'], row, sheet.hdrs);
            const pType = getCol(['prompt type'], row, sheet.hdrs);
            const domain= getCol(['domain'], row, sheet.hdrs);

            if (!qText) { flatIdx++; return; }

            const groupKey = `${sheet.sheetName}||${qNum || qText}`;

            if (!groupsMap.has(groupKey)) {
                groupsMap.set(groupKey, {
                    questionId: groupsMap.size + 1,
                    sheetName:  sheet.sheetName,
                    promptType: pType,
                    domain,
                    qNum,
                    questionText: qText,
                    rows: []
                });
            }
            groupsMap.get(groupKey).rows.push({ rowIndex: flatIdx, rowData: row, hdrs: sheet.hdrs });
            flatIdx++;
        });
    });

    groupedData = Array.from(groupsMap.values());

    // ── Pre-load any scores already filled in the Excel ──────────────────
    // Do NOT overwrite scores the annotator has already entered this session
    flatData.forEach((row, idx) => {
        // If this row already has an evaluation saved from local storage, skip it
        if (evaluations[idx] && evaluations[idx].ef !== '') return;

        const hdrs = Object.keys(row).filter(k => k !== '__sheet');

        const getC = (keys) => {
            const k = hdrs.find(h => keys.some(kk => h.toLowerCase().includes(kk.toLowerCase())));
            return k ? String(row[k]).trim() : '';
        };

        const ef = getC(['ef (0', 'ef(0', ' ef ']);
        const rr = getC(['rr (0', 'rr(0', ' rr ']);
        const lq = getC(['lq (0', 'lq(0', ' lq ']);
        const pp = getC(['pp (0', 'pp(0', ' pp ']);
        const comment = getC(['comments', 'comment']);

        // Only pre-load if at least one score exists in the Excel
        const validVals = ['0', '1', '2'];
        if (validVals.includes(ef) || validVals.includes(rr) || validVals.includes(lq) || validVals.includes(pp)) {
            evaluations[idx] = {
                ef:      validVals.includes(ef) ? ef : '',
                rr:      validVals.includes(rr) ? rr : '',
                lq:      validVals.includes(lq) ? lq : '',
                pp:      validVals.includes(pp) ? pp : '',
                errors:  [],
                comment: comment
            };
        }
    });

    // Persist the pre-loaded evaluations to local storage
    localStorage.setItem(`cabIgboEval_${annotatorId}`, JSON.stringify(evaluations));
}

// ─── Type Selection ───────────────────────────────────────────────────
function renderTypeSelection() {
    // Get unique sheet names (= types)
    const types = [...new Set(groupedData.map(g => g.sheetName))];

    // Friendly labels
    const labels = {
        'Type 1': 'Type 1 — English Question, English Answer (EN→EN)',
        'Type 2': 'Type 2 — English Question, Igbo Answer (EN→IG)',
        'Type 3': 'Type 3 — Igbo Question, Igbo Answer (IG→IG)',
    };

    ui.typeGrid.innerHTML = '';
    types.forEach(sheetName => {
        // Match "Type 1", "Type 2", "Type 3" anywhere in the sheet name
        const typeKey = Object.keys(labels).find(k => sheetName.includes(k)) || sheetName;
        const label = labels[typeKey] || sheetName;

        const qCount = groupedData.filter(g => g.sheetName === sheetName).length;
        const doneCount = groupedData.filter(g => g.sheetName === sheetName && isGroupComplete(g)).length;

        const card = document.createElement('div');
        card.className = 'type-card';
        card.innerHTML = `
            <h3>${label}</h3>
            <p>${doneCount} / ${qCount} questions completed</p>
            <div class="progress-bar"><div class="progress-fill" style="width:${qCount>0?(doneCount/qCount*100):0}%"></div></div>
        `;
        card.addEventListener('click', () => {
            selectedType = sheetName;
            filteredData = groupedData.filter(g => g.sheetName === sheetName);
            renderDashboard();
            showScreen('dashboard');
        });
        ui.typeGrid.appendChild(card);
    });
}

// ─── Dashboard ────────────────────────────────────────────────────────
function isGroupComplete(group) {
    return group.rows.every(r => {
        const ev = evaluations[r.rowIndex];
        return ev && ev.ef !== '' && ev.rr !== '' && ev.lq !== '' && ev.pp !== '';
    });
}

function renderDashboard() {
    // Derive friendly type label
    const typeKey = ['Type 1','Type 2','Type 3'].find(k => selectedType.includes(k)) || selectedType;
    ui.dashTitle.innerText = `${typeKey} — Questions`;

    let completedCount = 0;
    ui.qList.innerHTML = '';

    filteredData.forEach((group, i) => {
        const done = isGroupComplete(group);
        if (done) completedCount++;

        const card = document.createElement('div');
        card.className = `q-card ${done ? 'completed' : ''}`;
        card.innerHTML = `
            <h4>Q${group.qNum || (i + 1)} ${done ? '✓' : ''}</h4>
            <p title="${group.questionText}">${group.questionText}</p>
            ${group.domain ? `<span class="domain-tag">${group.domain}</span>` : ''}
        `;
        card.addEventListener('click', () => {
            currentGroupIndex = i;
            renderEvalScreen();
            showScreen('eval');
            window.scrollTo(0, 0);
        });
        ui.qList.appendChild(card);
    });

    ui.dashTotal.innerText     = filteredData.length;
    ui.dashCompleted.innerText = completedCount;
    ui.dashRemaining.innerText = filteredData.length - completedCount;
}

// ─── Back buttons ─────────────────────────────────────────────────────
ui.btnBackType.addEventListener('click', () => showScreen('type'));
ui.btnBackDash.addEventListener('click', () => {
    saveCurrentState();
    renderDashboard();
    showScreen('dashboard');
});

// ─── Evaluation Screen ────────────────────────────────────────────────
function renderEvalScreen() {
    const group = filteredData[currentGroupIndex];

    ui.currIdx.innerText   = currentGroupIndex + 1;
    ui.totalItems.innerText = filteredData.length;
    ui.metaType.innerText  = group.promptType || selectedType;
    ui.metaDomain.innerText = group.domain || '-';
    ui.metaQnum.innerText  = group.qNum || (currentGroupIndex + 1);
    ui.questionText.innerText = group.questionText;

    ui.modelsContainer.innerHTML = '';

    group.rows.forEach((r, idx) => {
        const hdrs = r.hdrs || headers;
        const getC = (keys) => {
            const k = hdrs.find(h => keys.some(kk => h.toLowerCase().includes(kk.toLowerCase())));
            return k ? String(r.rowData[k]).trim() : '';
        };

        const modelName    = getC(['model code', 'model name', 'model']) || `Model ${String.fromCharCode(65 + idx)}`;
        const responseText = getC(['model answer', 'answer', 'response', 'output', 'generation']) || 'No response found';
        const ev = evaluations[r.rowIndex] || { ef:'', rr:'', lq:'', pp:'', errors:[], comment:'' };

        const checkboxesHtml = ERROR_TYPES.map(err => {
            const checked = ev.errors.includes(err) ? 'checked' : '';
            return `<label class="checkbox-item"><input type="checkbox" name="err_${r.rowIndex}" value="${err}" ${checked}> ${err}</label>`;
        }).join('');

        const radioHtml = (metric, label, desc) => `
            <div class="metric-group">
                <label>${label}</label>
                <p class="metric-desc">${desc}</p>
                <div class="radio-group">
                    ${[0,1,2].map(v => `
                        <input type="radio" name="${metric}_${r.rowIndex}" id="${metric}${v}_${r.rowIndex}" value="${v}" ${ev[metric]===String(v)?'checked':''}>
                        <label for="${metric}${v}_${r.rowIndex}">${v}</label>
                    `).join('')}
                </div>
            </div>`;

        const card = document.createElement('div');
        card.className = 'model-card';
        card.innerHTML = `
            <h4>${modelName}</h4>
            <div class="scroll-box">${responseText}</div>
            <div class="evaluation-form">
                <div class="eval-grid">
                    ${radioHtml('ef', 'Epistemic Fidelity (EF)', '0 = Wrong or made-up &nbsp;|&nbsp; 1 = Partly correct &nbsp;|&nbsp; 2 = Fully correct and grounded')}
                    ${radioHtml('rr', 'Representational Richness (RR)', '0 = Wrong or empty &nbsp;|&nbsp; 1 = Knows a little &nbsp;|&nbsp; 2 = Deep and accurate')}
                    ${radioHtml('lq', 'Linguistic Quality (LQ)', '0 = Unreadable &nbsp;|&nbsp; 1 = Fluent with minor errors &nbsp;|&nbsp; 2 = Fluent, correct, well-structured')}
                    ${radioHtml('pp', 'Pragmatic Proficiency (PP)', '0 = Off-topic &nbsp;|&nbsp; 1 = Close but off &nbsp;|&nbsp; 2 = Perfect fit')}
                </div>
                <div class="metric-group" style="margin-top:1.5rem;">
                    <label>Error Taxonomy (select all that apply)</label>
                    <div class="taxonomy-grid">${checkboxesHtml}</div>
                </div>
                <div class="metric-group">
                    <label>Comments</label>
                    <textarea id="comment_${r.rowIndex}" rows="2" placeholder="Any notes...">${ev.comment}</textarea>
                </div>
            </div>`;
        ui.modelsContainer.appendChild(card);
    });

    ui.btnPrev.disabled = currentGroupIndex === 0;
    ui.btnNext.innerText = currentGroupIndex === filteredData.length - 1 ? 'Finish ✓' : 'Next →';
}

// ─── Save ─────────────────────────────────────────────────────────────
function saveCurrentState() {
    const group = filteredData[currentGroupIndex];
    group.rows.forEach(r => {
        const getRadio = metric => {
            const sel = document.querySelector(`input[name="${metric}_${r.rowIndex}"]:checked`);
            return sel ? sel.value : '';
        };
        const errors = [...document.querySelectorAll(`input[name="err_${r.rowIndex}"]:checked`)].map(cb => cb.value);
        const comment = (document.getElementById(`comment_${r.rowIndex}`) || {}).value || '';
        evaluations[r.rowIndex] = { ef: getRadio('ef'), rr: getRadio('rr'), lq: getRadio('lq'), pp: getRadio('pp'), errors, comment };
    });
    localStorage.setItem(`cabIgboEval_${annotatorId}`, JSON.stringify(evaluations));
}

ui.btnNext.addEventListener('click', () => {
    saveCurrentState();
    if (currentGroupIndex < filteredData.length - 1) {
        currentGroupIndex++;
        renderEvalScreen();
        window.scrollTo(0, 0);
    } else {
        renderDashboard();
        showScreen('dashboard');
    }
});

ui.btnPrev.addEventListener('click', () => {
    saveCurrentState();
    if (currentGroupIndex > 0) {
        currentGroupIndex--;
        renderEvalScreen();
        window.scrollTo(0, 0);
    }
});

ui.btnSave.addEventListener('click', () => {
    saveCurrentState();
    renderDashboard();
    showScreen('dashboard');
});

// ─── Export ───────────────────────────────────────────────────────────
ui.btnExportDash.addEventListener('click', () => {
    const exportData = flatData.map((row, idx) => {
        const ev = evaluations[idx] || { ef:'', rr:'', lq:'', pp:'', errors:[], comment:'' };
        return { ...row, Annotator_ID: annotatorId, EF_score: ev.ef, RR_score: ev.rr, LQ_score: ev.lq, PP_score: ev.pp, Error_Taxonomy: ev.errors.join(' | '), Comments: ev.comment };
    });
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Evaluations");
    XLSX.writeFile(wb, `CAB_Igbo_${annotatorId}_${selectedType.replace(/\s+/g,'_')}.xlsx`);
});

// ─── Guidelines Modal ─────────────────────────────────────────────────
const btnG = document.getElementById('btn-guidelines');
const modalG = document.getElementById('guidelines-modal');
const closeG = document.getElementById('close-modal');
btnG.addEventListener('click', () => modalG.classList.remove('hidden-modal'));
closeG.addEventListener('click', () => modalG.classList.add('hidden-modal'));
window.addEventListener('click', e => { if (e.target === modalG) modalG.classList.add('hidden-modal'); });
