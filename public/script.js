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

// Fetch and set real-time visitor count
async function initVisitorCount() {
    try {
        const res = await fetch('/api/visitor-count');
        const data = await res.json();
        const countStr = data.count.toString().padStart(6, '0');
        const digitSpans = document.querySelectorAll('.counter-digit');

        if (digitSpans.length === 6) {
            for (let i = 0; i < 6; i++) {
                digitSpans[i].innerText = countStr[i];
            }
        }
    } catch (err) {
        console.error("Failed to fetch visitor count:", err);
    }
}
initVisitorCount();

document.getElementById('certForm').addEventListener('submit', async (e) => {
    e.preventDefault();

    const senderEmail = document.getElementById('hiddenSenderEmail').value;
    if (!senderEmail) {
        alert("Please Sign in with Google first before generating certificates.");
        return;
    }

    const form = e.target;
    const submitBtn = document.getElementById('submitBtn');
    const progressSection = document.getElementById('progressSection');
    const downloadSection = document.getElementById('downloadSection');

    // Reset UI
    submitBtn.disabled = true;
    submitBtn.innerHTML = '⏳ Processing... Please wait';
    progressSection.classList.remove('d-none');
    downloadSection.classList.add('d-none');
    document.getElementById('progressBar').style.width = '0%';
    document.getElementById('progressBar').innerText = '0%';
    document.getElementById('totalStudents').innerText = '0';
    document.getElementById('sentEmails').innerText = '0';
    document.getElementById('failedEmails').innerText = '0';

    const formData = new FormData(form);

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
        submitBtn.disabled = false;
        submitBtn.innerHTML = '🚀 Generate & Send Certificates';
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
                submitBtn.disabled = false;
                submitBtn.innerHTML = '🚀 Generate & Send Certificates';
                return;
            }

            const data = await res.json();

            if (data.error) {
                clearInterval(interval);
                alert('Status check failed: ' + data.error);
                const submitBtn = document.getElementById('submitBtn');
                submitBtn.disabled = false;
                submitBtn.innerHTML = '🚀 Generate & Send Certificates';
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
                submitBtn.disabled = false;
                submitBtn.innerHTML = '🚀 Generate & Send Certificates';

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
