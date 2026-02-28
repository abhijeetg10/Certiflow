document.getElementById('themeToggle').addEventListener('click', () => {
    const html = document.documentElement;
    const currentTheme = html.getAttribute('data-bs-theme');
    html.setAttribute('data-bs-theme', currentTheme === 'dark' ? 'light' : 'dark');
});

// Google OAuth2 Logic
let oauthWindow;

document.getElementById('googleLoginBtn').addEventListener('click', () => {
    const width = 500;
    const height = 650;
    const left = window.screenX + (window.outerWidth - width) / 2;
    const top = window.screenY + (window.outerHeight - height) / 2;
    oauthWindow = window.open('/auth/google', 'GoogleAuth', `width=${width},height=${height},left=${left},top=${top}`);
});

window.addEventListener('message', (event) => {
    if (event.origin !== window.location.origin) return;

    if (event.data.type === 'oauth-success') {
        const { tokens, email } = event.data;

        document.getElementById('hiddenSenderEmail').value = email;
        document.getElementById('hiddenRefreshToken').value = tokens.refresh_token;

        document.getElementById('googleLoginBtn').classList.add('d-none');
        document.getElementById('googleProfile').classList.remove('d-none');
        document.getElementById('connectedEmail').innerText = email;
    }
});

document.getElementById('googleLogoutBtn').addEventListener('click', () => {
    document.getElementById('hiddenSenderEmail').value = '';
    document.getElementById('hiddenRefreshToken').value = '';

    document.getElementById('googleLoginBtn').classList.remove('d-none');
    document.getElementById('googleProfile').classList.add('d-none');
    document.getElementById('connectedEmail').innerText = '';
});

// Visitor count functionality removed per user request

// Handle Skip Email UI Toggle
const skipEmailToggle = document.getElementById('skipEmailToggle');
const emailSubjectInput = document.querySelector('input[name="subject"]');
const emailBodyInput = document.querySelector('textarea[name="body"]');

if (skipEmailToggle) {
    skipEmailToggle.addEventListener('change', (e) => {
        const isSkipped = e.target.checked;
        emailSubjectInput.disabled = isSkipped;
        emailBodyInput.disabled = isSkipped;

        if (isSkipped) {
            emailSubjectInput.closest('.col-12').style.opacity = '0.5';
            emailBodyInput.closest('.col-12').style.opacity = '0.5';
        } else {
            emailSubjectInput.closest('.col-12').style.opacity = '1';
            emailBodyInput.closest('.col-12').style.opacity = '1';
        }
    });
}

document.getElementById('certForm').addEventListener('submit', async (e) => {
    e.preventDefault();

    const senderEmail = document.getElementById('hiddenSenderEmail').value;
    const isSkipEmail = e.submitter && e.submitter.id === 'downloadZipBtn';

    if (!senderEmail && !isSkipEmail) {
        alert("Please Sign in with Google first before generating certificates.");
        return;
    }

    const form = e.target;
    const submitBtn = document.getElementById('submitBtn');
    const downloadZipBtn = document.getElementById('downloadZipBtn');
    const progressSection = document.getElementById('progressSection');
    const downloadSection = document.getElementById('downloadSection');

    // Reset UI
    if (submitBtn) submitBtn.disabled = true;
    if (downloadZipBtn) downloadZipBtn.disabled = true;

    if (isSkipEmail) {
        if (downloadZipBtn) downloadZipBtn.innerHTML = '⏳ Zipping... Please wait';
    } else {
        if (submitBtn) submitBtn.innerHTML = '⏳ Processing... Please wait';
    }

    progressSection.classList.remove('d-none');
    downloadSection.classList.add('d-none');
    document.getElementById('progressBar').style.width = '0%';
    document.getElementById('progressBar').innerText = '0%';
    document.getElementById('totalStudents').innerText = '0';
    document.getElementById('sentEmails').innerText = '0';
    document.getElementById('failedEmails').innerText = '0';

    const formData = new FormData(form);
    const payloadFields = fields.map(f => {
        let scaledFontSize = 30;
        let rawFontSize = parseFloat(document.getElementById(`fontSize_${f.id}`).value) || 30;
        if (currentScale && currentScale > 0) {
            scaledFontSize = rawFontSize / currentScale;
        }

        return {
            name: document.getElementById(`fieldName_${f.id}`).value || '[Empty Field]',
            fontSize: scaledFontSize.toFixed(1),
            fontFamily: document.getElementById(`fontFamily_${f.id}`).value,
            textColor: document.getElementById(`textColor_${f.id}`).value,
            xPos: document.getElementById(`xPos_${f.id}`).value,
            yPos: document.getElementById(`yPos_${f.id}`).value
        };
    });

    console.log("SENDING SCALED PAYLOAD FIELDS:", payloadFields);
    formData.append('fieldsPayload', JSON.stringify(payloadFields));
    formData.append('skipEmailToggle', isSkipEmail ? 'true' : 'false');

    try {
        const res = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });
        const data = await res.json();

        if (res.status !== 200 || data.error) {
            throw new Error(data.error || 'Upload failed');
        }

        document.getElementById('totalStudents').innerText = data.total;
        pollStatus(data.jobId);
    } catch (error) {
        alert('Error: ' + error.message);
        if (submitBtn) {
            submitBtn.disabled = false;
            submitBtn.innerHTML = '<i class="fa-solid fa-paper-plane me-2"></i> Generate & Send';
        }
        if (downloadZipBtn) {
            downloadZipBtn.disabled = false;
            downloadZipBtn.innerHTML = '<i class="fa-solid fa-file-zipper me-2"></i> Download ZIP Only';
        }
    }
});

function pollStatus(jobId) {
    const interval = setInterval(async () => {
        try {
            const res = await fetch(`/api/status/${jobId}`);

            if (!res.ok) {
                clearInterval(interval);
                alert('Oops! The server process restarted or interrupted. This can happen on free hosting if memory limits or timeouts are reached. Please refresh the page. If it happens again, try formatting your image template to a lighter file size!');
                const submitBtn = document.getElementById('submitBtn');
                const downloadZipBtn = document.getElementById('downloadZipBtn');
                if (submitBtn) { submitBtn.disabled = false; submitBtn.innerHTML = '<i class="fa-solid fa-paper-plane me-2"></i> Generate & Send'; }
                if (downloadZipBtn) { downloadZipBtn.disabled = false; downloadZipBtn.innerHTML = '<i class="fa-solid fa-file-zipper me-2"></i> Download ZIP Only'; }
                return;
            }

            const data = await res.json();

            if (data.error) {
                clearInterval(interval);
                alert('Status check failed: ' + data.error);
                const submitBtn = document.getElementById('submitBtn');
                const downloadZipBtn = document.getElementById('downloadZipBtn');
                if (submitBtn) { submitBtn.disabled = false; submitBtn.innerHTML = '<i class="fa-solid fa-paper-plane me-2"></i> Generate & Send'; }
                if (downloadZipBtn) { downloadZipBtn.disabled = false; downloadZipBtn.innerHTML = '<i class="fa-solid fa-file-zipper me-2"></i> Download ZIP Only'; }
                return;
            }

            document.getElementById('totalStudents').innerText = data.total;
            document.getElementById('sentEmails').innerText = data.success;
            document.getElementById('failedEmails').innerText = data.failed;

            const pct = data.total > 0 ? Math.round((data.processed / data.total) * 100) : 0;
            const progressEl = document.getElementById('progressBar');
            progressEl.style.width = pct + '%';
            progressEl.innerText = pct + '%';

            if (data.status === 'completed' || data.status === 'error') {
                clearInterval(interval);
                const submitBtn = document.getElementById('submitBtn');
                const downloadZipBtn = document.getElementById('downloadZipBtn');
                if (submitBtn) { submitBtn.disabled = false; submitBtn.innerHTML = '<i class="fa-solid fa-paper-plane me-2"></i> Generate & Send'; }
                if (downloadZipBtn) { downloadZipBtn.disabled = false; downloadZipBtn.innerHTML = '<i class="fa-solid fa-file-zipper me-2"></i> Download ZIP Only'; }

                if (data.status === 'completed') {
                    document.getElementById('downloadSection').classList.remove('d-none');
                    document.getElementById('downloadBtn').href = data.zipPath;
                } else {
                    alert('Zip creation failed or process encountered an error.');
                }
            }
        } catch (err) {
            console.error("Polling error", err);
        }
    }, 1500);
}

// Feedback Form Logic
document.getElementById('feedbackForm').addEventListener('submit', async (e) => {
    e.preventDefault();

    const name = document.getElementById('feedbackName').value.trim() || 'Anonymous';
    const message = document.getElementById('feedbackMessage').value.trim();
    const btn = document.getElementById('feedbackBtn');
    const successMsg = document.getElementById('feedbackSuccess');
    const errorMsg = document.getElementById('feedbackError');

    if (!message) return;

    btn.disabled = true;
    btn.innerHTML = '⏳ Submitting...';
    successMsg.classList.add('d-none');
    errorMsg.classList.add('d-none');

    try {
        const res = await fetch('/api/feedback', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ name, message })
        });

        if (!res.ok) throw new Error('Failed to submit feedback');

        successMsg.classList.remove('d-none');
        document.getElementById('feedbackForm').reset();
    } catch (err) {
        errorMsg.innerText = err.message || 'Something went wrong. Please try again.';
        errorMsg.classList.remove('d-none');
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<i class="fa-solid fa-paper-plane me-2"></i> Submit Feedback';
    }
});

// --- Visual Preview Logic & Multiple Fields ---
const templateFileInput = document.getElementById('templateFileInput');
const previewContainer = document.getElementById('previewContainer');
const templateCanvas = document.getElementById('templateCanvas');
const pdfWidthInput = document.getElementById('pdfWidthInput');
const pdfHeightInput = document.getElementById('pdfHeightInput');
const fieldsContainer = document.getElementById('fieldsContainer');
const addFieldBtn = document.getElementById('addFieldBtn');

const enableQrCodeCheckbox = document.getElementById('enableQrCode');
const qrSettingsDiv = document.getElementById('qrSettingsDiv');
const qrOverlay = document.getElementById('qrOverlay');
const qrXPosInput = document.getElementById('qrXPosInput');
const qrYPosInput = document.getElementById('qrYPosInput');

let currentScale = 1;
let actualPdfWidth = 0;
let actualPdfHeight = 0;
let fields = [];
let fieldCounter = 0;

function createFieldRow() {
    const id = fieldCounter++;
    const row = document.createElement('div');
    row.className = 'row g-3 mb-3 bg-white p-3 rounded border align-items-end shadow-sm field-row';
    row.id = `fieldRow_${id}`;

    row.innerHTML = `
        <div class="col-md-2">
            <label class="form-label fw-semibold fs-6 mb-1">Field Name (e.g. [Date])</label>
            <input type="text" class="form-control rounded-3 py-1" id="fieldName_${id}" value="[Name]">
        </div>
        <div class="col-md-2">
            <label class="form-label fw-semibold fs-6 mb-1">Color</label>
            <input type="color" class="form-control form-control-color rounded-3 p-1 w-100" id="textColor_${id}" value="#000000">
        </div>
        <div class="col-md-2">
            <label class="form-label fw-semibold fs-6 mb-1">Size</label>
            <input type="number" class="form-control rounded-3 py-1" id="fontSize_${id}" value="30">
        </div>
        <div class="col-md-3">
            <label class="form-label fw-semibold fs-6 mb-1">Font Family</label>
            <select class="form-select rounded-3 py-1" id="fontFamily_${id}">
                <option value="Helvetica">Helvetica (Default)</option>
                <option value="PinyonScript">Pinyon Script</option>
                <option value="GreatVibes">Great Vibes</option>
                <option value="AlexBrush">Alex Brush</option>
                <option value="DancingScript">Dancing Script</option>
                <option value="Parisienne">Parisienne</option>
                <option value="Playball">Playball</option>
                <option value="Rochester">Rochester</option>
                <option value="Satisfy">Satisfy</option>
                <option value="Tangerine">Tangerine</option>
                <option value="Allura">Allura</option>
                <option value="MrDeHaviland">Mr De Haviland</option>
            </select>
        </div>
        <div class="col-md-2 d-flex gap-1">
            <div>
                <label class="form-label fw-semibold fs-6 mb-1">X</label>
                <input type="number" step="0.1" class="form-control rounded-3 py-1" id="xPos_${id}" readonly>
            </div>
            <div>
                <label class="form-label fw-semibold fs-6 mb-1">Y</label>
                <input type="number" step="0.1" class="form-control rounded-3 py-1" id="yPos_${id}" readonly>
            </div>
        </div>
        <div class="col-md-1 text-center">
            <button type="button" class="btn btn-outline-danger btn-sm rounded-circle px-2 py-1 remove-field-btn" data-id="${id}"><i class="fa-solid fa-trash"></i></button>
        </div>
    `;

    // Create corresponding overlay
    const overlay = document.createElement('div');
    overlay.className = 'draggable-item position-absolute user-select-none font-pinyon';
    overlay.style.cursor = 'grab';
    overlay.style.display = previewContainer.style.display === 'none' ? 'none' : 'block';
    overlay.id = `overlay_${id}`;

    // Style settings based on row
    overlay.innerText = '[Student Name]';
    overlay.style.left = '50%';
    overlay.style.top = `${40 + (id * 10)}%`; // V-stack them slightly
    previewContainer.appendChild(overlay);

    fields.push({ id, row, overlay });
    fieldsContainer.appendChild(row);

    // Event listeners
    document.getElementById(`fieldName_${id}`).addEventListener('input', (e) => { overlay.innerText = e.target.value; updateTextStyle(id); });
    document.getElementById(`fontSize_${id}`).addEventListener('input', () => updateTextStyle(id));
    document.getElementById(`fontFamily_${id}`).addEventListener('change', () => updateTextStyle(id));
    document.getElementById(`textColor_${id}`).addEventListener('input', () => updateTextStyle(id));

    const removeBtn = row.querySelector('.remove-field-btn');
    removeBtn.addEventListener('click', () => {
        if (fields.length <= 1) return alert("You must have at least one field.");
        fields = fields.filter(f => f.id !== id);
        row.remove();
        overlay.remove();
    });

    initDraggable(overlay, document.getElementById(`xPos_${id}`), document.getElementById(`yPos_${id}`));
    setTimeout(() => updateTextStyle(id), 10);
    setTimeout(() => triggerInitialPositions(), 50);
}

function updateTextStyle(id) {
    const field = fields.find(f => f.id === id);
    if (!field) return;

    const size = parseFloat(document.getElementById(`fontSize_${id}`).value) || 30;
    field.overlay.style.fontSize = (size * currentScale) + 'px';
    field.overlay.style.color = document.getElementById(`textColor_${id}`).value;

    const family = document.getElementById(`fontFamily_${id}`).value;
    if (family === 'PinyonScript') field.overlay.style.fontFamily = "'Pinyon Script', cursive";
    else if (family === 'GreatVibes') field.overlay.style.fontFamily = "'Great Vibes', cursive";
    else if (family === 'AlexBrush') field.overlay.style.fontFamily = "'Alex Brush', cursive";
    else if (family === 'DancingScript') field.overlay.style.fontFamily = "'Dancing Script', cursive";
    else if (family === 'Parisienne') field.overlay.style.fontFamily = "'Parisienne', cursive";
    else if (family === 'Playball') field.overlay.style.fontFamily = "'Playball', cursive";
    else if (family === 'Rochester') field.overlay.style.fontFamily = "'Rochester', cursive";
    else if (family === 'Satisfy') field.overlay.style.fontFamily = "'Satisfy', cursive";
    else if (family === 'Tangerine') field.overlay.style.fontFamily = "'Tangerine', cursive";
    else if (family === 'Allura') field.overlay.style.fontFamily = "'Allura', cursive";
    else if (family === 'MrDeHaviland') field.overlay.style.fontFamily = "'Mr De Haviland', cursive";
    else field.overlay.style.fontFamily = "var(--font-body)";
}

function initDraggable(element, xInput, yInput) {
    let isDragging = false;

    element.addEventListener('mousedown', (e) => {
        isDragging = true;
        element.style.cursor = 'grabbing';
    });

    document.addEventListener('mousemove', (e) => {
        if (!isDragging) return;
        const rect = previewContainer.getBoundingClientRect();
        let x = e.clientX - rect.left;
        let y = e.clientY - rect.top;

        if (x < 0) x = 0;
        if (y < 0) y = 0;
        if (x > rect.width) x = rect.width;
        if (y > rect.height) y = rect.height;

        element.style.left = x + 'px';
        element.style.top = y + 'px';

        if (actualPdfWidth > 0 && xInput && yInput) {
            const scaleX = actualPdfWidth / rect.width;
            const scaleY = actualPdfHeight / rect.height;

            const actualX = x * scaleX;
            const actualY = y * scaleY;
            const pdfLibY = actualPdfHeight - actualY;

            xInput.value = actualX.toFixed(1);
            yInput.value = pdfLibY.toFixed(1);
        }
    });

    document.addEventListener('mouseup', () => {
        if (isDragging) {
            isDragging = false;
            element.style.cursor = 'grab';
        }
    });
}

function triggerInitialPositions() {
    fields.forEach(f => {
        const xInput = document.getElementById(`xPos_${f.id}`);
        const yInput = document.getElementById(`yPos_${f.id}`);
        if (actualPdfWidth > 0) {
            const rect = previewContainer.getBoundingClientRect();
            const scaleX = actualPdfWidth / rect.width;
            const scaleY = actualPdfHeight / rect.height;

            let px = parseFloat(f.overlay.style.left);
            let py = parseFloat(f.overlay.style.top);

            let realX = typeof f.overlay.style.left === 'string' && f.overlay.style.left.includes('%') ? rect.width * (parseFloat(f.overlay.style.left) / 100) : (px || rect.width / 2);
            let realY = typeof f.overlay.style.top === 'string' && f.overlay.style.top.includes('%') ? rect.height * (parseFloat(f.overlay.style.top) / 100) : (py || rect.height / 2);

            xInput.value = (realX * scaleX).toFixed(1);
            yInput.value = (actualPdfHeight - (realY * scaleY)).toFixed(1);
        }
    });
}

// Ensure at least one field row exists
addFieldBtn.addEventListener('click', createFieldRow);
createFieldRow(); // Initialize default Name row

// Form QR Toggle
enableQrCodeCheckbox.addEventListener('change', (e) => {
    if (e.target.checked) {
        qrSettingsDiv.classList.remove('d-none');
        if (previewContainer.style.display !== 'none') qrOverlay.classList.remove('d-none');
    } else {
        qrSettingsDiv.classList.add('d-none');
        qrOverlay.classList.add('d-none');
        qrXPosInput.value = '';
        qrYPosInput.value = '';
    }
});
initDraggable(qrOverlay, qrXPosInput, qrYPosInput);

if (templateFileInput) {
    templateFileInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) {
            previewContainer.style.display = 'none';
            return;
        }

        previewContainer.style.display = 'inline-block';
        fields.forEach(f => f.overlay.style.display = 'block');
        if (enableQrCodeCheckbox.checked) qrOverlay.classList.remove('d-none');

        const fileReader = new FileReader();

        const calculateAndRender = (w, h, renderFn) => {
            actualPdfWidth = w;
            actualPdfHeight = h;
            if (pdfWidthInput) pdfWidthInput.value = actualPdfWidth;
            if (pdfHeightInput) pdfHeightInput.value = actualPdfHeight;

            const maxW = 800;
            currentScale = w > maxW ? maxW / w : 1;
            templateCanvas.width = w * currentScale;
            templateCanvas.height = h * currentScale;

            renderFn();
            triggerInitialPositions();
            fields.forEach(f => updateTextStyle(f.id));

            // Trigger QR Initial
            qrOverlay.style.left = '80%';
            qrOverlay.style.top = '80%';
            if (actualPdfWidth > 0) {
                const rect = previewContainer.getBoundingClientRect();
                const scaleX = actualPdfWidth / rect.width;
                const scaleY = actualPdfHeight / rect.height;
                const rx = rect.width * 0.8;
                const ry = rect.height * 0.8;
                qrXPosInput.value = (rx * scaleX).toFixed(1);
                qrYPosInput.value = (actualPdfHeight - (ry * scaleY)).toFixed(1);
            }
        };

        if (file.type === 'application/pdf') {
            fileReader.onload = async function () {
                const typedarray = new Uint8Array(this.result);
                pdfjsLib.getDocument(typedarray).promise.then(pdf => {
                    pdf.getPage(1).then(page => {
                        const vp = page.getViewport({ scale: 1 });
                        calculateAndRender(vp.width, vp.height, () => {
                            page.render({
                                canvasContext: templateCanvas.getContext('2d'),
                                viewport: page.getViewport({ scale: currentScale })
                            });
                        });
                    });
                });
            };
            fileReader.readAsArrayBuffer(file);
        } else if (file.type.match('image.*')) {
            fileReader.onload = function (event) {
                const img = new Image();
                img.onload = () => calculateAndRender(img.width, img.height, () => {
                    templateCanvas.getContext('2d').drawImage(img, 0, 0, templateCanvas.width, templateCanvas.height);
                });
                img.src = event.target.result;
            }
            fileReader.readAsDataURL(file);
        }
    });
}
