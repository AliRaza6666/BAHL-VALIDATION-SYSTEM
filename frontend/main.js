// State Management
let selectedFile = null;
let currentFileId = null;
let validationAbortController = null;
const rawApiUrl = import.meta.env.VITE_API_URL;
const normalizedApiUrl = typeof rawApiUrl === 'string' ? rawApiUrl.trim() : '';
const API_URL = !normalizedApiUrl || normalizedApiUrl.includes('your-backend.vercel.app')
    ? '/api'
    : normalizedApiUrl.replace(/\/+$/, '');
const buildApiUrl = (path) => `${API_URL}${path.startsWith('/') ? path : `/${path}`}`;

// DOM Elements
const el = {
    profileSelect: document.getElementById('profile-select'),
    dropZone: document.getElementById('drop-zone'),
    fileInput: document.getElementById('file-input'),
    browseBtn: document.getElementById('browse-btn'),
    fileDetails: document.getElementById('file-details'),
    statusMessage: document.getElementById('status-message'),
    displayFilename: document.getElementById('display-filename'),
    displayRows: document.getElementById('display-rows'),
    validateActionGroup: document.getElementById('validate-action-group'),
    validateBtn: document.getElementById('validate-btn'),
    healthRetryBtn: document.getElementById('retry-health-btn'),
    
    progressContainer: document.getElementById('progress-container'),
    progressStatusText: document.getElementById('progress-status-text'),
    progressBar: document.getElementById('progress-bar'),
    progressPercent: document.getElementById('progress-percent'),
    
    resultDashboard: document.getElementById('result-dashboard'),
    resTotal: document.getElementById('res-total'),
    resPassed: document.getElementById('res-passed'),
    resFailed: document.getElementById('res-failed'),
    resTime: document.getElementById('res-time'),
    errorBreakdownList: document.getElementById('error-breakdown-list'),
    
    previewThead: document.getElementById('preview-thead'),
    previewTbody: document.getElementById('preview-tbody'),
    
    downloadPassedBtn: document.getElementById('download-passed-btn'),
    downloadRejectedBtn: document.getElementById('download-rejected-btn'),
    resetBtn: document.getElementById('reset-btn')
};

function setStatus(message, isError = false) {
    if (!el.statusMessage) return;
    if (!message) {
        el.statusMessage.textContent = '';
        el.statusMessage.className = 'mt-4 text-sm hidden';
        return;
    }

    el.statusMessage.textContent = message;
    el.statusMessage.className = `mt-4 text-sm rounded-lg border px-3 py-2 ${isError ? 'border-red-200 bg-red-50 text-red-700' : 'border-emerald-200 bg-emerald-50 text-emerald-700'}`;
}

function setAppEnabled(enabled) {
    if (enabled) {
        el.dropZone.classList.remove('opacity-50', 'pointer-events-none');
        el.browseBtn.disabled = false;
        el.fileInput.disabled = false;
        el.healthRetryBtn.classList.add('hidden');
    } else {
        el.dropZone.classList.add('opacity-50', 'pointer-events-none');
        el.browseBtn.disabled = true;
        el.fileInput.disabled = true;
    }
}

function getServiceErrorMessage(error) {
    if (!error || typeof error.message !== 'string') {
        return 'Service temporarily unavailable. Retry upload.';
    }

    const message = error.message;
    if (message.includes('NetworkError') || message.includes('failed to fetch') || message.includes('Service temporarily unavailable') || message.includes('Validation failed')) {
        return 'Service temporarily unavailable. Retry upload.';
    }

    return message;
}

async function checkHealth() {
    setStatus('Checking service status...', false);
    setAppEnabled(false);

    const healthUrl = buildApiUrl('/health');
    const maxAttempts = 3;

    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 5000);

        try {
            console.log("Health URL:", healthUrl);

            const response = await fetch(healthUrl, {
                method: 'GET',
                cache: 'no-store',
                signal: controller.signal
            });

            console.log("Response status:", response.status);
            console.log("Response ok:", response.ok);

            const data = await response.json();
            console.log("Response data:", data);

            if (!response.ok) {
                throw new Error('Service temporarily unavailable. Retry upload.');
            }

            if (data.status !== 'ok') {
                throw new Error('Service temporarily unavailable. Retry upload.');
            }

            clearTimeout(timeoutId);
            setStatus('Service available. Upload is enabled.', false);
            setAppEnabled(true);
            return true;

        } catch (error) {
            clearTimeout(timeoutId);

            console.error("[HEALTH] Error object:", error);

            if (attempt < maxAttempts) {
                await new Promise((resolve) => setTimeout(resolve, 500));
                continue;
            }

            setStatus(getServiceErrorMessage(error), true);
            setAppEnabled(false);
            el.healthRetryBtn.classList.remove('hidden');
            console.error('[HEALTH] Service check failed', error);
            return false;
        }
    }
}

window.addEventListener('DOMContentLoaded', () => {
    checkHealth().then((ok) => {
        if (!ok) {
            setStatus('Service temporarily unavailable. Retry upload.', true);
        }
    });
    setDownloadButtonsState(false);
    el.healthRetryBtn.classList.add('hidden');
    el.healthRetryBtn.addEventListener('click', async () => {
        await checkHealth();
    });
});

// Reset all validation state when new file is selected
function resetValidationState() {
    console.log('[STATE] Resetting validation state');
    
    // Cancel previous request if it's still pending
    if (validationAbortController) {
        validationAbortController.abort();
        validationAbortController = null;
    }
    
    // Clear UI elements
    el.progressContainer.classList.add('hidden');
    el.resultDashboard.classList.add('hidden');
    el.validateActionGroup.classList.add('hidden');
    el.previewThead.innerHTML = '';
    el.previewTbody.innerHTML = '';
    el.errorBreakdownList.innerHTML = '';
    
    // Reset state variables
    currentFileId = null;
    setDownloadButtonsState(false);
    
    // Reset profile selection
    el.profileSelect.value = 'auto';
}

// --- Drag & Drop / File Selection ---
el.dropZone.addEventListener('click', () => el.fileInput.click());
el.browseBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    el.fileInput.click();
});

el.fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) handleFileSelection(e.target.files[0]);
});

el.dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    el.dropZone.classList.add('drag-over');
});

el.dropZone.addEventListener('dragleave', () => {
    el.dropZone.classList.remove('drag-over');
});

el.dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    el.dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files.length) handleFileSelection(e.dataTransfer.files[0]);
});

// Analyze File Metadata in Background
async function handleFileSelection(file) {
    console.log('[FILE] New file selected:', file.name);
    
    // Reset validation state from previous upload/validation
    resetValidationState();
    
    const allowedTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel',
        'application/octet-stream'
    ];
    const isExcelName = /\.(xlsx|xls)$/i.test(file.name || '');
    const isExcelType = allowedTypes.includes(file.type) || isExcelName;

    if (!isExcelName || !isExcelType) {
        alert('Please upload a valid Excel file (.xlsx or .xls). Zip files and other archives are not supported.');
        return;
    }

    console.log('[UPLOAD] Uploading file...', file.name, file.type || 'unknown');
    console.log('Uploading to:', buildApiUrl('/analyze'));
    setStatus('Uploading...', false);
    
    selectedFile = file;
    el.displayFilename.textContent = file.name;
    el.displayRows.textContent = "Analyzing...";
    el.fileDetails.classList.remove('hidden');
    el.validateActionGroup.classList.add('hidden');
    
    // Call background analysis
    const formData = new FormData();
    formData.append('file', file);
    
    try {
        const response = await fetch(buildApiUrl('/analyze'), {
            method: 'POST',
            body: formData
        });
        console.log('Status:', response.status);
        
        if (!response.ok) {
            let err = { message: 'Service temporarily unavailable. Retry upload.' };
            try {
                err = await response.json();
            } catch {
                // ignore invalid JSON and keep the fallback message
            }
            throw new Error(err.message || err.error || 'Service temporarily unavailable. Retry upload.');
        }
        
        const data = await response.json();
        el.displayRows.textContent = data.row_count.toLocaleString();
        
        // Auto-select profile if detected
        if (data.profile && data.profile !== 'unknown') {
            el.profileSelect.value = data.profile;
        } else {
            el.profileSelect.value = 'auto';
        }
        
        setStatus(`File received. ${data.row_count.toLocaleString()} rows detected.`, false);
        console.log('[UPLOAD] Backend connected');
        console.log('[UPLOAD] Ready for validation');
        el.validateActionGroup.classList.remove('hidden');
    } catch (e) {
        const message = getServiceErrorMessage(e);
        setStatus(message, true);
        console.error('[UPLOAD] Upload failed', e);
        el.fileDetails.classList.add('hidden');
    }
}

// --- Validation Trigger ---
el.validateBtn.addEventListener('click', async () => {
    if (!selectedFile) return;
    
    console.log('[VALIDATION] Starting validation for:', selectedFile.name);
    
    // Cancel any previous validation request
    if (validationAbortController) {
        validationAbortController.abort();
    }
    validationAbortController = new AbortController();
    
    // Reset UI states
    el.progressContainer.classList.remove('hidden');
    el.resultDashboard.classList.add('hidden');
    el.validateBtn.disabled = true;
    
    const selectedProfile = el.profileSelect.value;
    
    const formData = new FormData();
    formData.append('file', selectedFile);
    formData.append('profile', selectedProfile);
    setStatus('Validation started...', false);
    console.log('[VALIDATION] Validation started');
    console.log('Uploading to:', buildApiUrl('/validate'));
    
    // Simulate steps in UI
    const steps = ['read', 'headers', 'rows', 'duplicates', 'export'];
    let currentStepIdx = 0;
    
    function setStep(stepKey, percent, text) {
        document.querySelectorAll('.step').forEach(s => {
            s.classList.remove('bg-emerald-100', 'text-emerald-700', 'font-bold');
            s.classList.add('text-gray-400');
        });
        const activeStep = document.querySelector(`.step[data-step="${stepKey}"]`);
        if (activeStep) {
            activeStep.classList.remove('text-gray-400');
            activeStep.classList.add('bg-emerald-100', 'text-emerald-700', 'font-bold');
        }
        el.progressStatusText.textContent = text;
        el.progressBar.style.width = `${percent}%`;
        el.progressPercent.textContent = `${percent}%`;
    }
    
    // Simulated steps interval
    const stepInterval = setInterval(() => {
        if (currentStepIdx < steps.length - 1) {
            currentStepIdx++;
            const step = steps[currentStepIdx];
            const percent = (currentStepIdx + 1) * 20;
            const texts = {
                headers: 'Checking Headers...',
                rows: 'Validating Rows...',
                duplicates: 'Checking Duplicates...',
                export: 'Generating Excel...'
            };
            setStep(step, percent, texts[step]);
        }
    }, 400);
    
    // Start step
    setStep('read', 20, 'Reading File...');
    
    try {
        const response = await fetch(buildApiUrl('/validate'), {
            method: 'POST',
            body: formData,
            signal: validationAbortController.signal
        });
        console.log('Status:', response.status);
        
        clearInterval(stepInterval);
        
        if (!response.ok) {
            let err = { message: 'Service temporarily unavailable. Retry upload.' };
            try {
                err = await response.json();
            } catch {
                // ignore invalid JSON and keep the fallback message
            }
            throw new Error(err.message || err.error || 'Service temporarily unavailable. Retry upload.');
        }
        
        const data = await response.json();
        
        // Fast forward progress bar to 100%
        setStep('export', 100, 'Validation Complete!');
        setStatus('Validation complete. Download is ready.', false);
        console.log('[VALIDATION] Validation complete');
        setTimeout(() => {
            el.progressContainer.classList.add('hidden');
            renderDashboard(data);
            el.validateBtn.disabled = false;
            
            // Reset file input so same file can be selected again
            el.fileInput.value = '';
            console.log('[VALIDATION] Ready for next upload');
        }, 300);
        
    } catch (e) {
        clearInterval(stepInterval);
        
        // Don't show error if request was aborted (new file selected)
        if (e.name === 'AbortError') {
            console.log('[VALIDATION] Validation cancelled (new file selected)');
            el.validateBtn.disabled = false;
            return;
        }
        
        const message = getServiceErrorMessage(e);
        setStatus(message, true);
        console.error('[VALIDATION] Validation failed', e);
        el.progressContainer.classList.add('hidden');
        el.validateBtn.disabled = false;
        
        // Still reset file input so user can try again
        el.fileInput.value = '';
    }
});

// --- Render Dashboard ---
function renderDashboard(data) {
    currentFileId = data.file_id;
    
    el.resTotal.textContent = data.summary.total.toLocaleString();
    el.resPassed.textContent = data.summary.passed.toLocaleString();
    el.resFailed.textContent = data.summary.rejected.toLocaleString();
    el.resTime.textContent = data.duration;
    
    // Error Breakdown
    el.errorBreakdownList.innerHTML = '';
    let hasErrors = false;
    Object.entries(data.summary.errors_breakdown).forEach(([name, count]) => {
        if (count > 0) {
            hasErrors = true;
            const li = document.createElement('li');
            li.className = "bg-red-50 text-red-700 px-3 py-1.5 rounded-lg border border-red-200 flex items-center gap-1.5";
            li.innerHTML = `❌ <span>${name}</span> <span class="bg-red-200 text-red-800 px-2 py-0.5 rounded-full text-[10px] font-bold">${count}</span>`;
            el.errorBreakdownList.appendChild(li);
        }
    });
    
    if (!hasErrors) {
        const li = document.createElement('li');
        li.className = "bg-emerald-50 text-emerald-700 px-3 py-1.5 rounded-lg border border-emerald-200";
        li.textContent = "🎉 All records passed verification successfully!";
        el.errorBreakdownList.appendChild(li);
    }
    
    // Enable/disable the report download buttons
    setDownloadButtonsState(!!currentFileId);

    renderPreviewTable(data);
    
    el.resultDashboard.classList.remove('hidden');
}

// --- Render Preview Table ---
function renderPreviewTable(data) {
    el.previewThead.innerHTML = '';
    el.previewTbody.innerHTML = '';
    
    // Headers
    const trH = document.createElement('tr');
    [...data.headers, 'Status'].forEach(h => {
        const th = document.createElement('th');
        th.className = "px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wider border-b border-gray-200";
        th.textContent = h;
        trH.appendChild(th);
    });
    el.previewThead.appendChild(trH);
    
    // Rows
    data.rows.forEach((row, i) => {
        const res = data.results[i];
        const tr = document.createElement('tr');
        tr.className = res.status === 'PASS' ? 'pass border-b border-gray-100 hover:bg-emerald-50/20' : 'fail border-b border-gray-100 hover:bg-red-50/20';
        
        data.headers.forEach((h, colIdx) => {
            const td = document.createElement('td');
            td.className = "px-4 py-3 text-xs text-gray-700 whitespace-nowrap";
            
            const cellVal = row[h];
            td.textContent = cellVal !== undefined ? cellVal : '';
            
            // Check if this specific cell has a validation error
            const err = res.errors.find(e => e.col === colIdx);
            if (err) {
                td.className = "px-4 py-3 text-xs whitespace-nowrap invalid-cell";
                td.setAttribute('data-error', `Rule: ${err.field}\nExpected: ${err.expected}\nActual: ${err.actual}`);
            }
            tr.appendChild(td);
        });
        
        // Status Cell
        const tdStatus = document.createElement('td');
        tdStatus.className = `px-4 py-3 text-xs font-bold ${res.status === 'PASS' ? 'text-emerald-600' : 'text-red-600'}`;
        tdStatus.textContent = res.status;
        tr.appendChild(tdStatus);
        
        el.previewTbody.appendChild(tr);
    });
}

function setDownloadButtonsState(enabled) {
    [el.downloadPassedBtn, el.downloadRejectedBtn].forEach((button) => {
        if (enabled) {
            button.disabled = false;
            button.classList.remove('opacity-50', 'cursor-not-allowed', 'bg-gray-200', 'text-gray-500', 'text-slate-700', 'border-gray-300');
            button.classList.add('text-white', 'shadow-md', 'rounded-lg', 'font-bold');
            if (button.id === 'download-passed-btn') {
                button.classList.add('bg-[#006B3F]', 'hover:bg-[#005233]');
                button.classList.remove('bg-[#16A34A]', 'hover:bg-[#15803D]');
            } else {
                button.classList.add('bg-[#B91C1C]', 'hover:bg-[#991B1B]');
                button.classList.remove('bg-[#C62828]', 'hover:bg-[#A61B1B]');
            }
        } else {
            button.disabled = true;
            button.classList.remove('bg-[#006B3F]', 'hover:bg-[#005233]', 'bg-[#B91C1C]', 'hover:bg-[#991B1B]', 'text-white');
            button.classList.add('opacity-70', 'cursor-not-allowed', 'bg-gray-100', 'text-slate-700', 'border', 'border-gray-300');
        }
    });
}

// --- Download File Handlers ---
el.downloadPassedBtn.addEventListener('click', () => {
    if (currentFileId) {
        window.location.href = buildApiUrl(`/api/download/passed/${currentFileId}`);
    }
});

el.downloadRejectedBtn.addEventListener('click', () => {
    if (currentFileId) {
        window.location.href = buildApiUrl(`/api/download/rejected/${currentFileId}`);
    }
});

// --- Reset App ---
el.resetBtn.addEventListener('click', () => {
    console.log('[STATE] User clicked reset button');
    selectedFile = null;
    currentFileId = null;
    
    // Cancel any pending validation
    if (validationAbortController) {
        validationAbortController.abort();
        validationAbortController = null;
    }
    
    el.fileInput.value = '';
    el.fileDetails.classList.add('hidden');
    el.validateActionGroup.classList.add('hidden');
    el.progressContainer.classList.add('hidden');
    el.resultDashboard.classList.add('hidden');
    el.previewThead.innerHTML = '';
    el.previewTbody.innerHTML = '';
    el.errorBreakdownList.innerHTML = '';
    setStatus('', false);
    el.profileSelect.value = 'auto';
    el.validateBtn.disabled = false;
    console.log('[STATE] App reset to idle state');
});
