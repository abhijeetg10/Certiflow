const express = require('express');
const multer = require('multer');
const csv = require('csv-parser');
const xlsx = require('xlsx');
const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');
const fontkit = require('@pdf-lib/fontkit');
const nodemailer = require('nodemailer');
const archiver = require('archiver');
const fs = require('fs');
const path = require('path');
const cors = require('cors');
require('dotenv').config();

const { google } = require('googleapis');
const qrcode = require('qrcode');

const app = express();
const port = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

const UPLOADS_DIR = path.join(__dirname, 'uploads');
const GENERATED_DIR = path.join(__dirname, 'generated');

if (!fs.existsSync(UPLOADS_DIR)) fs.mkdirSync(UPLOADS_DIR, { recursive: true });
if (!fs.existsSync(GENERATED_DIR)) fs.mkdirSync(GENERATED_DIR, { recursive: true });

// Visitor Counter logic removed per user request

// Multer Configuration
const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, UPLOADS_DIR),
    filename: (req, file, cb) => cb(null, Date.now() + '-' + file.originalname)
});
const upload = multer({ storage });

const jobs = {};

// OAuth2 endpoints
app.get('/auth/google', (req, res) => {
    if (!process.env.CLIENT_ID || !process.env.CLIENT_SECRET) {
        return res.status(500).send("Developer error: Google Client ID/Secret missing in .env");
    }
    const oauth2Client = new google.auth.OAuth2(
        process.env.CLIENT_ID,
        process.env.CLIENT_SECRET,
        process.env.REDIRECT_URI || 'http://localhost:3000/auth/google/callback'
    );
    const url = oauth2Client.generateAuthUrl({
        access_type: 'offline',
        prompt: 'consent', // Force consent so we always get a refresh token
        scope: ['https://mail.google.com/', 'https://www.googleapis.com/auth/userinfo.email', 'https://www.googleapis.com/auth/userinfo.profile']
    });
    res.redirect(url);
});

app.get('/auth/google/callback', async (req, res) => {
    const code = req.query.code;
    try {
        const oauth2Client = new google.auth.OAuth2(
            process.env.CLIENT_ID,
            process.env.CLIENT_SECRET,
            process.env.REDIRECT_URI || 'http://localhost:3000/auth/google/callback'
        );
        const { tokens } = await oauth2Client.getToken(code);
        oauth2Client.setCredentials(tokens);

        // Get user email
        const oauth2 = google.oauth2({ version: 'v2', auth: oauth2Client });
        const userInfo = await oauth2.userinfo.get();
        const email = userInfo.data.email;

        // Pass tokens back to opener window securely
        res.send(`
            <script>
                window.opener.postMessage({
                    type: 'oauth-success',
                    tokens: ${JSON.stringify(tokens)},
                    email: '${email}'
                }, window.location.origin);
                window.close();
            </script>
        `);
    } catch (error) {
        console.error('Error in callback:', error);
        res.status(500).send('Authentication failed');
    }
});

app.post('/api/upload', upload.fields([{ name: 'csvFile' }, { name: 'dataFile' }, { name: 'templateFile' }]), async (req, res) => {
    try {
        const uploadedDataFile = req.files['dataFile'] ? req.files['dataFile'][0] : (req.files['csvFile'] ? req.files['csvFile'][0] : null);
        if (!uploadedDataFile || !req.files['templateFile']) {
            return res.status(400).json({ error: 'Missing Data or Template files' });
        }

        const dataFile = uploadedDataFile;
        const templateFile = req.files['templateFile'][0];

        const jobId = Date.now().toString();
        const jobDir = path.join(GENERATED_DIR, jobId);
        fs.mkdirSync(jobDir, { recursive: true });

        jobs[jobId] = {
            total: 0,
            processed: 0,
            success: 0,
            failed: 0,
            status: 'parsing',
            zipPath: null
        };

        const students = [];
        const startProcessing = async () => {
            jobs[jobId].total = students.length;
            jobs[jobId].status = 'processing';
            res.json({ jobId, total: students.length, message: 'Processing started' });

            try {
                const config = {
                    senderEmail: req.body.senderEmail,
                    refreshToken: req.body.refreshToken,
                    subject: req.body.subject,
                    body: req.body.body,
                    fieldsPayload: req.body.fieldsPayload,
                    enableQrCode: req.body.enableQrCode === 'on' || req.body.enableQrCode === 'true',
                    skipEmail: req.body.skipEmailToggle === 'on' || req.body.skipEmailToggle === 'true',
                    qrXPos: req.body.qrXPos,
                    qrYPos: req.body.qrYPos
                };
                await processBatch(jobId, students, templateFile, config, jobDir);
            } catch (err) {
                console.error('Batch error:', err);
                jobs[jobId].status = 'error';
            }

            if (fs.existsSync(dataFile.path)) fs.unlinkSync(dataFile.path);
            if (fs.existsSync(templateFile.path)) fs.unlinkSync(templateFile.path);
        };

        const ext = path.extname(dataFile.originalname).toLowerCase();
        if (ext === '.xlsx' || ext === '.xls') {
            try {
                const workbook = xlsx.readFile(dataFile.path);
                const sheetName = workbook.SheetNames[0];
                const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });
                students.push(...sheetData);
                startProcessing();
            } catch (err) {
                console.error(err);
                if (fs.existsSync(dataFile.path)) fs.unlinkSync(dataFile.path);
                res.status(500).json({ error: 'Failed to read Excel file' });
            }
        } else if (ext === '.pdf') {
            try {
                const pdfParse = require('pdf-parse');
                const dataBuffer = fs.readFileSync(dataFile.path);
                const pdfData = await pdfParse(dataBuffer);
                const lines = pdfData.text.split('\n').filter(l => l.trim().length > 0);

                for (let i = 0; i < lines.length; i++) {
                    const line = lines[i].trim();
                    const emailMatch = line.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/i);
                    if (emailMatch) {
                        const email = emailMatch[1];
                        let name = line.replace(email, '').replace(/,/g, '').trim();
                        // If no name found on the same line, check the previous line
                        if (!name && i > 0) {
                            if (!lines[i - 1].match(/@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+/)) {
                                name = lines[i - 1].replace(/,/g, '').trim();
                            }
                        }
                        if (!name) name = 'Student';

                        // To be compatible with our smart field detection, add standard object keys
                        students.push({ Name: name, Email: email });
                    }
                }
                startProcessing();
            } catch (err) {
                console.error(err);
                if (fs.existsSync(dataFile.path)) fs.unlinkSync(dataFile.path);
                res.status(500).json({ error: 'Failed to read PDF file' });
            }
        } else {
            fs.createReadStream(dataFile.path)
                .pipe(csv())
                .on('data', (data) => students.push(data))
                .on('end', startProcessing)
                .on('error', (err) => {
                    console.error(err);
                    res.status(500).json({ error: 'Failed to read CSV' });
                });
        }
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'Server error starting the job' });
    }
});

async function processBatch(jobId, students, templateFile, config, jobDir) {
    const { senderEmail, refreshToken, subject, body, fieldsPayload, enableQrCode, skipEmail, qrXPos, qrYPos } = config;
    const templateBytes = fs.readFileSync(templateFile.path);
    const isPdf = templateFile.originalname.toLowerCase().endsWith('.pdf');

    let transporter, gmail;
    if (!skipEmail) {
        const oauth2Client = new google.auth.OAuth2(
            process.env.CLIENT_ID,
            process.env.CLIENT_SECRET,
            process.env.REDIRECT_URI || 'http://localhost:3000/auth/google/callback'
        );

        oauth2Client.setCredentials({ refresh_token: refreshToken });

        try {
            const res = await oauth2Client.getAccessToken();
            // token gets automatically managed by google-auth-library
        } catch (err) {
            console.error("Failed to acquire access token:", err);
            jobs[jobId].status = 'error';
            return;
        }

        transporter = nodemailer.createTransport({
            streamTransport: true,
            newline: 'windows'
        });

        gmail = google.gmail({ version: 'v1', auth: oauth2Client });
    }

    // Cache the base PDF template OUTSIDE the loop to prevent OOM memory crashes on Render
    let basePdfDoc;
    if (isPdf) {
        basePdfDoc = await PDFDocument.load(templateBytes);
    } else {
        basePdfDoc = await PDFDocument.create();
        let image;
        if (templateFile.originalname.toLowerCase().endsWith('.png')) {
            image = await basePdfDoc.embedPng(templateBytes);
        } else {
            image = await basePdfDoc.embedJpg(templateBytes);
        }
        const page = basePdfDoc.addPage([image.width, image.height]);
        page.drawImage(image, { x: 0, y: 0, width: image.width, height: image.height });
    }

    basePdfDoc.registerFontkit(fontkit);

    let customFontNames = [
        'PinyonScript', 'GreatVibes', 'AlexBrush', 'DancingScript',
        'Parisienne', 'Playball', 'Rochester', 'Satisfy', 'Tangerine',
        'Allura', 'MrDeHaviland'
    ];
    let loadedCustomFonts = {};
    for (let fName of customFontNames) {
        const fontPath = path.join(__dirname, 'fonts', `${fName}-Regular.ttf`);
        if (fs.existsSync(fontPath)) {
            loadedCustomFonts[fName] = fs.readFileSync(fontPath);
        }
    }

    const hexToRgbFn = (hex) => {
        if (!hex) return rgb(0, 0, 0);
        hex = hex.replace(/^#/, '');
        if (hex.length === 3) hex = hex.split('').map(c => c + c).join('');
        const int = parseInt(hex, 16);
        return rgb(((int >> 16) & 255) / 255, ((int >> 8) & 255) / 255, (int & 255) / 255);
    };

    let fields = [];
    try {
        if (fieldsPayload) fields = JSON.parse(fieldsPayload);
    } catch (e) {
        console.error("Failed to parse fieldsPayload:", e);
    }

    const BATCH_SIZE = 3;
    for (let i = 0; i < students.length; i += BATCH_SIZE) {
        const batch = students.slice(i, i + BATCH_SIZE);

        await Promise.all(batch.map(async (bStudent) => {
            let studentName = null;
            let studentEmail = null;
            for (let key in bStudent) {
                const normalizedKey = key.toString().toLowerCase().trim();
                if (normalizedKey.includes('name')) studentName = bStudent[key];
                if (normalizedKey.includes('mail')) studentEmail = bStudent[key];
            }
            if (!studentName) studentName = Object.values(bStudent)[0];
            if (!studentEmail) studentEmail = Object.values(bStudent)[1];

            console.log(`[Cert Rendering] Processing student: ${studentName}, Email: ${studentEmail}`);

            if (!studentName || !studentEmail) {
                if (jobs[jobId]) jobs[jobId].failed++;
                if (jobs[jobId]) jobs[jobId].processed++;
                return;
            }

            try {
                // Instatiate a blank document and securely copy the cached template
                const pdfDoc = await PDFDocument.create();
                pdfDoc.registerFontkit(fontkit);
                const copiedPages = await pdfDoc.copyPages(basePdfDoc, [0]);
                const firstPage = copiedPages[0];
                pdfDoc.addPage(firstPage);

                const fontCache = {};

                // Render dynamic text fields
                for (let field of fields) {
                    let fieldText = field.name || "";
                    console.log(`[Before Replace] Original form text field: "${fieldText}"`);

                    // Direct string lookup for accuracy over regex
                    if (fieldText.toLowerCase().includes('[name]')) {
                        fieldText = fieldText.replace(/\[name\]/gi, studentName || '');
                    } else if (fieldText.toLowerCase().includes('[student name]')) {
                        fieldText = fieldText.replace(/\[student name\]/gi, studentName || '');
                    }

                    console.log(`[After Replace] Generated text field: "${fieldText}"`);

                    // Support email variable replacement if requested
                    if (fieldText.toLowerCase().includes('[email]')) {
                        fieldText = fieldText.replace(/\[email\]/gi, studentEmail || '');
                    }

                    // Skip empty fields
                    if (!fieldText.trim() || fieldText === '[Empty Field]') continue;

                    let currentFont;
                    if (field.fontFamily && loadedCustomFonts[field.fontFamily]) {
                        if (!fontCache[field.fontFamily]) {
                            fontCache[field.fontFamily] = await pdfDoc.embedFont(loadedCustomFonts[field.fontFamily]);
                        }
                        currentFont = fontCache[field.fontFamily];
                    } else {
                        if (!fontCache['Helvetica']) {
                            fontCache['Helvetica'] = await pdfDoc.embedFont(StandardFonts.Helvetica);
                        }
                        currentFont = fontCache['Helvetica'];
                    }

                    const size = parseFloat(field.fontSize) || 30;
                    const textWidth = currentFont.widthOfTextAtSize(fieldText, size);
                    const textHeight = currentFont.heightAtSize(size);

                    let x = parseFloat(field.xPos);
                    let y = parseFloat(field.yPos);

                    if (!isNaN(x)) x = x - (textWidth / 2);
                    if (!isNaN(y)) {
                        // Use CSS optical centering equivalence for line-height: 1
                        y = y - (size * 0.28);
                    }

                    if (isNaN(x)) x = (firstPage.getWidth() / 2) - (textWidth / 2);
                    if (isNaN(y)) y = (firstPage.getHeight() / 2) - (textHeight / 2);

                    firstPage.drawText(fieldText, {
                        x,
                        y,
                        size,
                        font: currentFont,
                        color: hexToRgbFn(field.textColor)
                    });
                }

                // Embed QR code if enabled
                if (enableQrCode) {
                    const qrData = `Verified Certificate for ${studentName} (${studentEmail})`;
                    const qrDataUrl = await qrcode.toDataURL(qrData, { margin: 1, color: { dark: '#000000', light: '#ffffff' } });
                    const qrBuffer = Buffer.from(qrDataUrl.split(',')[1], 'base64');
                    const qrImage = await pdfDoc.embedPng(qrBuffer);

                    const qrDim = 100; // Resize QR to sensible 100x100
                    let qX = parseFloat(qrXPos);
                    let qY = parseFloat(qrYPos);

                    if (!isNaN(qX)) qX = qX - (qrDim / 2);
                    if (!isNaN(qY)) qY = qY - (qrDim / 2);

                    if (isNaN(qX)) qX = firstPage.getWidth() - qrDim - 20;
                    if (isNaN(qY)) qY = 20;

                    firstPage.drawImage(qrImage, {
                        x: qX,
                        y: qY,
                        width: qrDim,
                        height: qrDim
                    });
                }

                const pdfBytes = await pdfDoc.save();
                const fileName = `${studentName.replace(/[^a-zA-Z0-9]/g, '_')}_certificate.pdf`;
                const filePath = path.join(jobDir, fileName);
                fs.writeFileSync(filePath, pdfBytes);

                if (!skipEmail) {
                    const mailOptions = {
                        from: senderEmail,
                        to: studentEmail,
                        subject: subject || "Your Participation Certificate",
                        text: body ? body.replace(/\[Name\]/gi, studentName) : `Dear ${studentName}, Please find attached your certificate.`,
                        attachments: [
                            {
                                filename: fileName,
                                path: filePath
                            }
                        ]
                    };

                    const info = await transporter.sendMail(mailOptions);
                    let raw = '';
                    for await (const chunk of info.message) {
                        raw += chunk.toString();
                    }
                    const encodedMessage = Buffer.from(raw).toString('base64')
                        .replace(/\+/g, '-')
                        .replace(/\//g, '_')
                        .replace(/=+$/, '');

                    await gmail.users.messages.send({
                        userId: 'me',
                        requestBody: {
                            raw: encodedMessage
                        }
                    });
                }

                if (jobs[jobId]) jobs[jobId].success++;
            } catch (error) {
                console.error(`Error processing ${studentEmail}:`, error.message || error);
                if (jobs[jobId]) jobs[jobId].failed++;
            }
            if (jobs[jobId]) jobs[jobId].processed++;
        }));

        // Add a 1.5 second delay between batches to respect Gmail rate limits
        if (i + BATCH_SIZE < students.length) {
            await new Promise(resolve => setTimeout(resolve, 1500));
        }
    }

    // transorter.close() is not needed for streamTransport

    try {
        const zipPath = path.join(GENERATED_DIR, `${jobId}.zip`);
        const output = fs.createWriteStream(zipPath);
        const archive = archiver('zip', { zlib: { level: 9 } });

        output.on('close', async () => {
            if (jobs[jobId]) {
                jobs[jobId].status = 'completed';
                jobs[jobId].zipPath = `/api/download/${jobId}`;

                await logToGoogleSheet(senderEmail || 'ZIP Download [Skipped Email]', jobs[jobId].success, jobs[jobId].failed);
            }
            try {
                if (fs.existsSync(jobDir)) {
                    fs.rmSync(jobDir, { recursive: true, force: true });
                }
            } catch (cleanupErr) {
                console.error('Directory cleanup failed:', cleanupErr);
            }
        });

        archive.on('error', (err) => {
            console.error('Archiver error:', err);
            if (jobs[jobId]) jobs[jobId].status = 'error';
        });

        archive.pipe(output);
        archive.directory(jobDir, false);
        await archive.finalize();
    } catch (err) {
        console.error('Zip generation failed:', err);
        if (jobs[jobId]) jobs[jobId].status = 'error';
    }
}

async function logToGoogleSheet(userEmail, successCount, failedCount) {
    if (!process.env.SPREADSHEET_ID || !process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || !process.env.GOOGLE_PRIVATE_KEY) return;
    if (!userEmail) userEmail = 'ZIP Download [Skipped Email]';

    try {
        const auth = new google.auth.GoogleAuth({
            credentials: {
                client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
                private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n')
            },
            scopes: ['https://www.googleapis.com/auth/spreadsheets']
        });

        const sheets = google.sheets({ version: 'v4', auth });

        const timestamp = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });
        const request = {
            spreadsheetId: process.env.SPREADSHEET_ID,
            range: 'A:D',
            valueInputOption: 'USER_ENTERED',
            insertDataOption: 'INSERT_ROWS',
            resource: {
                values: [
                    [timestamp, userEmail, successCount, failedCount]
                ]
            }
        };

        await sheets.spreadsheets.values.append(request);
        console.log(`Logged stats to Google Sheets for ${userEmail}`);
    } catch (error) {
        console.error("Failed to log to Google Sheets:", error);
    }
}

app.get('/api/status/:jobId', (req, res) => {
    const job = jobs[req.params.jobId];
    if (!job) return res.status(404).json({ error: 'Job not found' });
    res.json(job);
});

// Visitor count endpoint removed

app.get('/api/download/:jobId', (req, res) => {
    const zipPath = path.join(GENERATED_DIR, `${req.params.jobId}.zip`);
    if (!fs.existsSync(zipPath)) return res.status(404).json({ error: 'Zip not found' });
    res.download(zipPath);
});

app.post('/api/feedback', async (req, res) => {
    const { name, message } = req.body;

    if (!message) {
        return res.status(400).json({ error: 'Feedback message is required' });
    }

    if (!process.env.SPREADSHEET_ID || !process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || !process.env.GOOGLE_PRIVATE_KEY) {
        console.error("Missing Google Sheets credentials for feedback");
        return res.status(500).json({ error: 'Server configuration error' });
    }

    try {
        const auth = new google.auth.GoogleAuth({
            credentials: {
                client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
                private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n')
            },
            scopes: ['https://www.googleapis.com/auth/spreadsheets']
        });

        const sheets = google.sheets({ version: 'v4', auth });
        const timestamp = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });

        // Note: Appending to columns F, G, H so it doesn't conflict with certificate logs (A, B, C, D)
        const request = {
            spreadsheetId: process.env.SPREADSHEET_ID,
            range: 'F:H',
            valueInputOption: 'USER_ENTERED',
            insertDataOption: 'INSERT_ROWS',
            resource: {
                values: [
                    [timestamp, name, message]
                ]
            }
        };

        await sheets.spreadsheets.values.append(request);
        res.status(200).json({ success: true });
    } catch (error) {
        console.error('Failed to log feedback to Google Sheets:', error);
        res.status(500).json({ error: 'Failed to submit feedback' });
    }
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
