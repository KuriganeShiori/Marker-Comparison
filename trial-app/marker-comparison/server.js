const express = require('express');
const path = require('path');
const multer = require('multer');
const fs = require('fs').promises;
const { DataHandler } = require('./data');
const MarkerComparison = require('./comparison');
const { DocGenerator } = require('./docGenerator');
const app = express();

// Configure multer for file upload
const upload = multer({
    storage: multer.memoryStorage(),
    limits: {
        fileSize: 10 * 1024 * 1024 // 10MB limit
    }
});

// Serve static files from public directory first (for assets)
app.use(express.static(path.join(__dirname, '..', '..', 'public')));

// Then serve files from marker-comparison directory
app.use(express.static(path.join(__dirname)));

// Error handling middleware
app.use((err, req, res, next) => {
    console.error('Server Error:', err);
    res.status(500).json({
        error: err.message,
        details: process.env.NODE_ENV === 'development' ? err.stack : undefined
    });
});

// Serve static files with proper MIME types
app.use((req, res, next) => {
    if (req.path.endsWith('.js')) {
        res.type('application/javascript');
    }
    next();
});

app.use(express.json());

// API endpoints for data operations
app.post('/api/initialize', async (req, res) => {
    try {
        const dataHandler = new DataHandler();
        console.log('Initializing connection to spreadsheet...');
        const connected = await dataHandler.initialize();
        console.log('Spreadsheet connection result:', connected);
        res.json({ success: connected });
    } catch (error) {
        console.error('Initialization error:', error);
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/compare', async (req, res) => {
    try {
        const { type, params } = req.body;
        console.log('Comparison request:', { type, params });
        const dataHandler = new DataHandler();
        const comparison = new MarkerComparison(dataHandler);
        let result;
        
        switch (type) {
            case 'family':
                result = await comparison.compareFamily(params.code);
                break;
            case 'sameDay':
                result = await comparison.compareSameDay(params.code);
                break;
            case 'all':
                result = await comparison.compareAllDatabase(params.code);
                break;
            case 'two':
                result = await comparison.compareSamples(params.code1, params.code2);
                break;
            default:
                throw new Error(`Unknown comparison type: ${type}`);
        }
        
        // Ensure result is always an array
        result = Array.isArray(result) ? result : [result];
        res.json(result);
    } catch (error) {
        console.error('Comparison error:', error);
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/generate-document', async (req, res) => {
    try {
        const result = req.body;
        const { buffer, filename } = await DocGenerator.generateDocument(result);
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        const encodedFilename = encodeURIComponent(filename);
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodedFilename}`);
        res.send(buffer);
    } catch (error) {
        console.error('Document generation error:', error);
        console.error('Error details:', error.stack);
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/upload', upload.array('files'), async (req, res) => {
    try {
        const dateFolder = req.body.dateFolder;
        const files = req.files;
        const dataHandler = new DataHandler();

        console.log('Processing files:', files.length);
        // Process uploaded files
        const casesData = {};
        for (const file of files) {
            console.log('Processing file:', file.originalname);
            const [caseFolder] = file.originalname.split('/');
            if (!casesData[caseFolder]) {
                casesData[caseFolder] = [];
            }
            
            const content = file.buffer.toString('utf8');
            console.log('File content sample:', content.substring(0, 200));
            const data = dataHandler.parseTxtFile(content);
            console.log('Parsed data:', {
                code: data.code,
                name: data.name,
                markerCount: Object.keys(data.markers).length
            });
            casesData[caseFolder].push(data);
        }

        console.log('Uploading to Google Sheets...');
        // Upload to Google Sheets
        await dataHandler.uploadCasesData(dateFolder, casesData);

        res.json({ success: true });
    } catch (error) {
        console.error('Upload error:', error);
        console.error('Error stack:', error.stack);
        res.status(500).json({ 
            error: error.message,
            details: error.stack,
            type: error.constructor.name
        });
    }
});

app.get('/api/check-case', async (req, res) => {
    try {
        const { dateFolder, baseCode } = req.query;
        const dataHandler = new DataHandler();
        const exists = await dataHandler.findExistingCase(baseCode, dateFolder);
        res.json({ exists: !!exists });
    } catch (error) {
        console.error('Error checking case:', error);
        res.status(500).json({ error: error.message });
    }
});

// Serve index.html
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
}); 