const express = require('express');
const multer = require('multer');
const csv = require('csv-parser');
const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');
const nodemailer = require('nodemailer');
const archiver = require('archiver');
const fs = require('fs');
const path = require('path');
const cors = require('cors');
require('dotenv').config();

const { google } = require('googleapis');

const app = express();
const port = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

const UPLOADS_DIR = path.join(__dirname, 'uploads');
const GENERATED_DIR = path.join(__dirname, 'generated');

if (!fs.existsSync(UPLOADS_DIR)) fs.mkdirSync(UPLOADS_DIR, { recursive: true });
if (!fs.existsSync(GENERATED_DIR)) fs.mkdirSync(GENERATED_DIR, { recursive: true });

// Persistent Visitor Counter
const COUNTER_FILE = path.join(__dirname, 'visitor_count.json');
let visitorCount = 0;
if (fs.existsSync(COUNTER_FILE)) {
    try {
        const data = fs.readFileSync(COUNTER_FILE, 'utf8');
        visitorCount = JSON.parse(data).count || 0;
    } catch (e) { console.error("Could not read counter file"); }
}

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

app.post('/api/upload', upload.fields([{ name: 'csvFile' }, { name: 'templateFile' }]), async (req, res) => {
    try {
        if (!req.files || !req.files['csvFile'] || !req.files['templateFile']) {
            return res.status(400).json({ error: 'Missing CSV or Template files' });
        }

        const csvFile = req.files['csvFile'][0];
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
        fs.createReadStream(csvFile.path)
            .pipe(csv())
            .on('data', (data) => students.push(data))
            .on('end', async () => {
                jobs[jobId].total = students.length;
                jobs[jobId].status = 'processing';
                res.json({ jobId, total: students.length, message: 'Processing started' });

                try {
                    await processBatch(jobId, students, templateFile, req.body, jobDir);
                } catch (err) {
                    console.error('Batch error:', err);
                    jobs[jobId].status = 'error';
                }

                if (fs.existsSync(csvFile.path)) fs.unlinkSync(csvFile.path);
                if (fs.existsSync(templateFile.path)) fs.unlinkSync(templateFile.path);
            })
            .on('error', (err) => {
                console.error(err);
                res.status(500).json({ error: 'Failed to read CSV' });
            });
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'Server error starting the job' });
    }
});

async function processBatch(jobId, students, templateFile, config, jobDir) {
    const { senderEmail, refreshToken, subject, body, fontSize, xPos, yPos } = config;
    const templateBytes = fs.readFileSync(templateFile.path);
    const isPdf = templateFile.originalname.toLowerCase().endsWith('.pdf');

    const oauth2Client = new google.auth.OAuth2(
        process.env.CLIENT_ID,
        process.env.CLIENT_SECRET,
        process.env.REDIRECT_URI || 'http://localhost:3000/auth/google/callback'
    );

    oauth2Client.setCredentials({ refresh_token: refreshToken });

    let accessToken;
    try {
        const res = await oauth2Client.getAccessToken();
        accessToken = res.token;
    } catch (err) {
        console.error("Failed to acquire access token:", err);
        jobs[jobId].status = 'error';
        return;
    }

    const transporter = nodemailer.createTransport({
        service: 'gmail',
        pool: true,
        maxConnections: 3,
        maxMessages: 100,
        auth: {
            type: 'OAuth2',
            user: senderEmail,
            clientId: process.env.CLIENT_ID,
            clientSecret: process.env.CLIENT_SECRET,
            refreshToken: refreshToken,
            accessToken: accessToken
        }
    });

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

    const helveticaFont = await basePdfDoc.embedFont(StandardFonts.Helvetica);

    const BATCH_SIZE = 3;
    for (let i = 0; i < students.length; i += BATCH_SIZE) {
        const batch = students.slice(i, i + BATCH_SIZE);

        await Promise.all(batch.map(async (bStudent) => {
            const studentName = bStudent.Name || bStudent.name || bStudent.NAME || Object.values(bStudent)[0];
            const studentEmail = bStudent.Email || bStudent.email || bStudent.EMAIL || Object.values(bStudent)[1];

            if (!studentName || !studentEmail) {
                if (jobs[jobId]) jobs[jobId].failed++;
                if (jobs[jobId]) jobs[jobId].processed++;
                return;
            }

            try {
                // Instatiate a blank document and securely copy the cached template
                const pdfDoc = await PDFDocument.create();
                const copiedPages = await pdfDoc.copyPages(basePdfDoc, [0]);
                const firstPage = copiedPages[0];
                pdfDoc.addPage(firstPage);

                // We need to re-embed the font for the cloned document instance
                const currentFont = await pdfDoc.embedFont(StandardFonts.Helvetica);

                const size = parseFloat(fontSize) || 30;
                const textWidth = currentFont.widthOfTextAtSize(studentName, size);
                const textHeight = currentFont.heightAtSize(size);

                let x = parseFloat(xPos);
                let y = parseFloat(yPos);
                if (isNaN(x)) x = (firstPage.getWidth() / 2) - (textWidth / 2);
                if (isNaN(y)) y = (firstPage.getHeight() / 2) - (textHeight / 2);

                firstPage.drawText(studentName, {
                    x,
                    y,
                    size,
                    font: currentFont,
                    color: rgb(0, 0, 0)
                });

                const pdfBytes = await pdfDoc.save();
                const fileName = `${studentName.replace(/[^a-zA-Z0-9]/g, '_')}_certificate.pdf`;
                const filePath = path.join(jobDir, fileName);
                fs.writeFileSync(filePath, pdfBytes);

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

                await transporter.sendMail(mailOptions);
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

    transporter.close();

    try {
        const zipPath = path.join(GENERATED_DIR, `${jobId}.zip`);
        const output = fs.createWriteStream(zipPath);
        const archive = archiver('zip', { zlib: { level: 9 } });

        output.on('close', () => {
            if (jobs[jobId]) {
                jobs[jobId].status = 'completed';
                jobs[jobId].zipPath = `/api/download/${jobId}`;
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

app.get('/api/status/:jobId', (req, res) => {
    const job = jobs[req.params.jobId];
    if (!job) return res.status(404).json({ error: 'Job not found' });
    res.json(job);
});

// Visitor Count Endpoint
app.get('/api/visitor-count', (req, res) => {
    visitorCount++;
    fs.promises.writeFile(COUNTER_FILE, JSON.stringify({ count: visitorCount }))
        .catch(err => console.error("Failed to save visitor count", err));
    res.json({ count: visitorCount });
});

app.get('/api/download/:jobId', (req, res) => {
    const zipPath = path.join(GENERATED_DIR, `${req.params.jobId}.zip`);
    if (!fs.existsSync(zipPath)) return res.status(404).json({ error: 'Zip not found' });
    res.download(zipPath);
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
