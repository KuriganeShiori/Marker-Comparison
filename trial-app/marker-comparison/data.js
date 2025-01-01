const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const MarkerComparison = require('./comparison');

class DataHandler {
    constructor() {
        try {
            // Look for credentials in environment variables first
            if (process.env.GOOGLE_SHEETS_CREDENTIALS) {
                this.auth = new google.auth.GoogleAuth({
                    credentials: JSON.parse(process.env.GOOGLE_SHEETS_CREDENTIALS),
                    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
                });
            } else {
                // Fallback to file-based credentials
                const credentialsPath = path.join(__dirname, 'credentials', 'credentials.json');
                if (!fs.existsSync(credentialsPath)) {
                    throw new Error('Google Sheets credentials not found in environment or file');
                }
                this.auth = new google.auth.GoogleAuth({
                    keyFile: credentialsPath,
                    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
                });
            }
            this.sheetsApi = google.sheets({ version: 'v4', auth: this.auth });
            this.spreadsheetId = '1WEpy6eVaUoNUYqJLKfe715eW1U_RWgDkZenomBtdCrY';
            
            console.log('Spreadsheet URL:', this.getDatabaseUrl());
            console.log('Service account credentials path:', credentialsPath);
        } catch (error) {
            console.error('Error initializing DataHandler:', error);
            throw error;
        }
    }

    async initialize() {
        try {
            // Test the connection by getting spreadsheet info
            const response = await this.sheetsApi.spreadsheets.get({
                spreadsheetId: this.spreadsheetId
            });
            console.log('Successfully connected to spreadsheet:', response.data.properties.title);
            return true;
        } catch (error) {
            console.error('Failed to initialize Google Sheets connection:', error);
            throw error;
        }
    }

    async findExistingCase(baseCode, sheetName) {
        try {
            const response = await this.sheetsApi.spreadsheets.values.get({
                spreadsheetId: this.spreadsheetId,
                range: `${sheetName}!A:A`
            });

            if (response.data.values) {
                const rows = response.data.values;
                for (let i = 0; i < rows.length; i++) {
                    if (rows[i][0] && rows[i][0].startsWith(baseCode)) {
                        return i + 1; // Return 1-based row number
                    }
                }
            }
            return null;
        } catch (error) {
            console.error('Error searching for existing case:', error);
            return null;
        }
    }

    async deleteRows(startRow, numRows, sheetName) {
        try {
            // Get sheet ID first
            const response = await this.sheetsApi.spreadsheets.get({
                spreadsheetId: this.spreadsheetId
            });
            
            const sheet = response.data.sheets.find(s => 
                s.properties.title.toLowerCase() === sheetName.toLowerCase()
            );
            
            if (!sheet) {
                throw new Error(`Sheet ${sheetName} not found`);
            }

            await this.sheetsApi.spreadsheets.batchUpdate({
                spreadsheetId: this.spreadsheetId,
                resource: {
                    requests: [{
                        deleteDimension: {
                            range: {
                                sheetId: sheet.properties.sheetId,
                                dimension: 'ROWS',
                                startIndex: startRow - 1,
                                endIndex: startRow - 1 + numRows
                            }
                        }
                    }]
                }
            });
        } catch (error) {
            console.error('Error deleting rows:', error);
            throw error;
        }
    }

    async checkSpreadsheet() {
        try {
            console.log('Checking spreadsheet connection...');
            const response = await this.sheetsApi.spreadsheets.get({
                spreadsheetId: this.spreadsheetId
            });
            console.log('Successfully connected to spreadsheet:', response.data.properties.title);
            return true;
        } catch (error) {
            console.error('Spreadsheet connection error:', error);
            return false;
        }
    }

    async findOrCreateSheet(sheetName) {
        try {
            // Get all sheets
            const response = await this.sheetsApi.spreadsheets.get({
                spreadsheetId: this.spreadsheetId
            });

            // Check if sheet exists
            const sheet = response.data.sheets.find(s => 
                s.properties.title.toLowerCase() === sheetName.toLowerCase()
            );

            if (sheet) {
                console.log(`Found existing sheet: ${sheetName}`);
                return sheet.properties.title;
            }

            // Create new sheet if it doesn't exist
            console.log(`Creating new sheet: ${sheetName}`);
            const addSheetResponse = await this.sheetsApi.spreadsheets.batchUpdate({
                spreadsheetId: this.spreadsheetId,
                resource: {
                    requests: [{
                        addSheet: {
                            properties: {
                                title: sheetName
                            }
                        }
                    }]
                }
            });

            return addSheetResponse.data.replies[0].addSheet.properties.title;
        } catch (error) {
            console.error('Error managing sheets:', error);
            throw error;
        }
    }

    async uploadData(folderPath) {
        try {
            this.replaceAll = false;
            console.log('Starting data upload from:', folderPath);
            const dateFolder = path.basename(folderPath);
            
            // Find or create sheet for this date
            const sheetName = await this.findOrCreateSheet(dateFolder);
            
            const caseFolders = fs.readdirSync(folderPath)
                .filter(folder => /^V\d{4}[A-Z]/.test(folder))
                .sort();
            
            const casePairs = {};
            caseFolders.forEach(folder => {
                const baseCode = folder.slice(0, 5);
                if (!casePairs[baseCode]) {
                    casePairs[baseCode] = [];
                }
                casePairs[baseCode].push(folder);
            });

            for (const [baseCode, caseGroup] of Object.entries(casePairs)) {
                // Check if case already exists in the specific sheet
                const existingRow = await this.findExistingCase(baseCode, sheetName);
                if (existingRow && !this.replaceAll) {
                    const response = await this.electron.dialog.showMessageBox({
                        type: 'question',
                        buttons: ['Skip', 'Replace', 'Replace All'],
                        defaultId: 0,
                        title: 'Case Already Exists',
                        message: `Case ${baseCode} already exists in the database. What would you like to do?`
                    });

                    if (response.response === 0) continue; // Skip
                    if (response.response === 1 || response.response === 2) {
                        // Delete existing rows
                        await this.deleteRows(existingRow, 28, sheetName); // Header + markers + spacing
                        if (response.response === 2) {
                            this.replaceAll = true;
                        }
                    }
                } else if (existingRow && this.replaceAll) {
                    // Automatically replace if replaceAll is true
                    await this.deleteRows(existingRow, 28, sheetName);
                }

                const values = [];
                const numSamples = caseGroup.length;
                
                // Process each case to get data first
                const caseData = [];
                for (const caseFolder of caseGroup) {
                    const casePath = path.join(folderPath, caseFolder);
                    if (!fs.statSync(casePath).isDirectory()) continue;

                    const files = fs.readdirSync(casePath)
                        .filter(file => file.endsWith('.txt'));

                    for (const file of files) {
                        const filePath = path.join(casePath, file);
                        const content = fs.readFileSync(filePath, 'utf8');
                        const data = this.parseTxtFile(content);
                        caseData.push({
                            folder: caseFolder,
                            data: data
                        });
                    }
                }

                // First row: Sample names
                const nameRow = [];
                caseData.forEach((sample, index) => {
                    nameRow.push(`${sample.data.code} ${sample.data.name}`);
                    // Add empty cells to align with columns
                    nameRow.push('', '');
                    // Add single space between samples except for the last one
                    if (index < numSamples - 1) nameRow.push('');
                });

                // Second row: Column headers
                const headerRow = [];
                caseData.forEach((_, index) => {
                    headerRow.push('Marker', 'Allele 1', 'Allele 2');
                    // Add single space between samples except for the last one
                    if (index < numSamples - 1) headerRow.push('');
                });

                values.push(nameRow);
                values.push(headerRow);

                // Add marker data rows
                const markerOrder = [
                    'D3S1358', 'vWA', 'D16S539', 'CSF1PO', 'D6S1043',
                    'Yindel', 'AMEL', 'D8S1179', 'D21S11', 'D18S51',
                    'D5S818', 'D2S441', 'D19S433', 'FGA', 'D10S1248',
                    'D22S1045', 'D1S1656', 'D13S317', 'D7S820', 'Penta E',
                    'Penta D', 'TH01', 'D12S391', 'D2S1338', 'TPOX'
                ];

                markerOrder.forEach(marker => {
                    const row = [];
                    caseData.forEach((sample, index) => {
                        const values = sample.data.markers[marker] || ['', ''];
                        const formattedValues = values.map(v => {
                            const num = parseFloat(v);
                            return !isNaN(num) ? `'${v}` : v;
                        });
                        row.push(marker, formattedValues[0] || '', formattedValues[1] || '');
                        // Add single space between samples except for the last one
                        if (index < numSamples - 1) row.push('');
                    });
                    values.push(row);
                });

                // Add empty rows between cases
                values.push(Array(numSamples * 4 - (numSamples - 1)).fill(''));
                values.push(Array(numSamples * 4 - (numSamples - 1)).fill(''));
                values.push(Array(numSamples * 4 - (numSamples - 1)).fill(''));

                try {
                    await this.sheetsApi.spreadsheets.values.append({
                        spreadsheetId: this.spreadsheetId,
                        range: sheetName,
                        valueInputOption: 'USER_ENTERED',
                        insertDataOption: 'INSERT_ROWS',
                        resource: {
                            values,
                            majorDimension: 'ROWS'
                        }
                    });

                    console.log('Successfully uploaded data for case group:', baseCode);
                } catch (uploadError) {
                    console.error('Upload error response:', uploadError.response?.data);
                    console.error('Upload error status:', uploadError.response?.status);
                    throw uploadError;
                }
            }

            return true;
        } catch (error) {
            console.error('Error uploading data:', error);
            throw error;
        }
    }

    parseTxtFile(content) {
        // New implementation matching the actual format
        const sampleData = {
            markers: {},
            code: '',
            name: ''
        };

        // Split the file data into lines and filter out empty lines
        const lines = content.split('\n').filter(line => line.trim());
        
        // First pass: Process lines with two values
        for (let i = 1; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) continue;

            const parts = line.split('\t');
            
            // Get sample code and name from the first data line if not set yet
            if (!sampleData.code || !sampleData.name) {
                const sampleNameParts = parts[1].split(' ');
                sampleData.code = sampleNameParts[0];
                sampleData.name = sampleNameParts.slice(1).join(' ');
            }

            const markerName = parts[2] ? parts[2].trim() : '';
            const value1 = parts[3] ? parts[3].trim() : '';
            const value2 = parts[4] ? parts[4].trim() : '';

            // Only process if both values exist
            if (value1 && value2) {
                sampleData.markers[markerName] = [value1, value2];
            }
        }

        // Second pass: Process lines with one value
        for (let i = 1; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) continue;

            const parts = line.split('\t');
            const markerName = parts[2] ? parts[2].trim() : '';
            const value1 = parts[3] ? parts[3].trim() : '';

            // Only process if marker doesn't exist yet and has one value
            if (!sampleData.markers[markerName] && value1 && (!parts[4] || !parts[4].trim())) {
                sampleData.markers[markerName] = [value1, value1]; // Duplicate the single value
            }
        }

        return sampleData;
    }

    getDatabaseUrl() {
        return 'https://docs.google.com/spreadsheets/d/1WEpy6eVaUoNUYqJLKfe715eW1U_RWgDkZenomBtdCrY/edit#gid=0';
    }

    async getAllSheets() {
        try {
            const response = await this.sheetsApi.spreadsheets.get({
                spreadsheetId: this.spreadsheetId
            });
            return response.data.sheets.map(sheet => sheet.properties.title);
        } catch (error) {
            console.error('Error getting sheets:', error);
            throw error;
        }
    }

    async getSheetData(sheetName) {
        try {
            const response = await this.sheetsApi.spreadsheets.values.get({
                spreadsheetId: this.spreadsheetId,
                range: `${sheetName}!A:Z`
            });

            const values = response.data.values || [];
            console.log(`Raw data from sheet ${sheetName}:`, values.slice(0, 5));
            const cases = [];
            let currentCaseStartRow = -1;
            let currentCaseEndRow = -1;

            // First pass: Find case boundaries
            for (let i = 0; i < values.length; i++) {
                const row = values[i];
                
                // Skip truly empty rows
                if (!row || row.length === 0) {
                    if (currentCaseStartRow !== -1) {
                        // End of current case
                        const caseData = this.processCaseData(values, currentCaseStartRow, currentCaseEndRow);
                        if (caseData.samples.length > 0) {
                            cases.push(caseData);
                        }
                        currentCaseStartRow = -1;
                        currentCaseEndRow = -1;
                    }
                    continue;
                }

                // If we find a row starting with a case code (V####X)
                if (row[0] && typeof row[0] === 'string' && row[0].match(/^V\d{4}[A-Z]/)) {
                    if (currentCaseStartRow !== -1) {
                        // Process previous case
                        const caseData = this.processCaseData(values, currentCaseStartRow, currentCaseEndRow);
                        if (caseData.samples.length > 0) {
                            cases.push(caseData);
                        }
                    }
                    currentCaseStartRow = i;
                    console.log(`Found new case starting at row ${i}:`, row[0]);
                }

                // Update end row if we have a start row
                if (currentCaseStartRow !== -1) {
                    currentCaseEndRow = i;
                }
            }

            // Process the last case if exists
            if (currentCaseStartRow !== -1) {
                const caseData = this.processCaseData(values, currentCaseStartRow, currentCaseEndRow);
                if (caseData.samples.length > 0) {
                    cases.push(caseData);
                }
            }

            console.log(`Processed ${cases.length} cases from sheet ${sheetName}`);
            cases.forEach(caseData => {
                console.log(`Case ${caseData.baseCode} processed with ${caseData.samples.length} samples:`,
                    caseData.samples.map(s => s.code));
            });

            return cases;
        } catch (error) {
            console.error(`Error getting data from sheet ${sheetName}:`, error);
            return [];
        }
    }

    processCaseData(values, startRow, endRow) {
        // First row contains sample information
        const headerRow = values[startRow];
        console.log('Processing case header row:', headerRow);
        const baseCode = headerRow[0].substring(0, 5);
        const samples = [];

        // Process sample headers
        for (let j = 0; j < headerRow.length; j++) {
            const sampleCode = headerRow[j];
            // Check if this column contains a sample code (V####X format)
            if (sampleCode && typeof sampleCode === 'string' && sampleCode.startsWith(baseCode)) {
                console.log('Found sample in header:', sampleCode);
                const [code, ...nameParts] = sampleCode.split(' ');
                samples.push({
                    code: code,
                    name: nameParts.join(' '),
                    markers: {},
                    columnIndex: j  // Store the column index for this sample
                });
            }
        }

        console.log(`Found ${samples.length} samples in case ${baseCode}:`, samples.map(s => s.code));

        // Skip the header row and marker header row
        let markerStartRow = startRow + 2;

        // Process marker data
        for (let i = markerStartRow; i <= endRow; i++) {
            const row = values[i];
            if (!row || !row[0] || row[0] === 'Marker') continue;

            const marker = row[0];
            console.log(`Processing marker ${marker} for case ${baseCode}`);

            samples.forEach((sample) => {
                // Use the stored column index to find the correct alleles
                const sampleStartCol = sample.columnIndex;
                const allele1 = row[sampleStartCol + 1];  // Allele 1 is one column after sample name
                const allele2 = row[sampleStartCol + 2];  // Allele 2 is two columns after sample name
                if (allele1 || allele2) {
                    console.log(`Sample ${sample.code}, Marker ${marker}:`, [allele1, allele2]);
                    sample.markers[marker] = [allele1 || '', allele2 || ''];
                }
            });
        }

        // Verify marker data was properly assigned
        samples.forEach(sample => {
            console.log(`Sample ${sample.code} has ${Object.keys(sample.markers).length} markers`);
            if (Object.keys(sample.markers).length === 0) {
                console.warn(`Warning: No markers found for sample ${sample.code}`);
            }
            delete sample.columnIndex;  // Remove the temporary column index
        });

        // Additional verification for multiple samples
        if (samples.length === 1) {
            console.warn(`Warning: Only one sample found for case ${baseCode}. Expected multiple samples (A, B, C, etc.)`);
        }

        const caseData = {
            baseCode,
            samples
        };
        console.log(`Processed case ${baseCode}:`, caseData);
        return caseData;
    }

    async getAllCases() {
        try {
            const sheets = await this.getAllSheets();
            let allCases = [];

            for (const sheet of sheets) {
                if (sheet === 'Sheet1') continue;
                console.log(`Fetching data from sheet: ${sheet}`);
                const cases = await this.getSheetData(sheet);
                console.log(`Found ${cases.length} cases in sheet ${sheet}`);
                cases.forEach(caseData => {
                    console.log(`Case ${caseData.baseCode} has ${caseData.samples.length} samples:`, 
                        caseData.samples.map(s => s.code));
                });
                allCases = allCases.concat(cases);
            }

            // Debug log each case's data
            allCases.forEach(caseData => {
                console.log(`Case ${caseData.baseCode} contains samples:`, 
                    caseData.samples.map(s => ({
                        code: s.code,
                        markerCount: Object.keys(s.markers).length,
                        markers: Object.keys(s.markers)
                    }))
                );
            });

            return allCases;
        } catch (error) {
            console.error('Error getting all cases:', error);
            throw error;
        }
    }

    async uploadCasesData(dateFolder, casesData) {
        try {
            const sheetName = await this.findOrCreateSheet(dateFolder);
            console.log('Using sheet:', sheetName);

            // Sort case folders first
            const sortedFolders = Object.keys(casesData).sort((a, b) => {
                // Extract the base codes (first 5 characters)
                const codeA = a.slice(0, 5);
                const codeB = b.slice(0, 5);
                
                // Extract parts: letter (V/T), year (24), and number (45)
                const letterA = codeA[0];
                const letterB = codeB[0];
                const yearA = codeA.slice(1, 3);
                const yearB = codeB.slice(1, 3);
                const numberA = codeA.slice(3, 5);
                const numberB = codeB.slice(3, 5);

                // Compare letter first
                if (letterA !== letterB) {
                    return letterA.localeCompare(letterB);
                }
                // Then compare year
                if (yearA !== yearB) {
                    return yearA.localeCompare(yearB);
                }
                // Finally compare number
                return numberA.localeCompare(numberB);
            });

            // First, organize the data by case code
            const caseGroups = {};
            for (const caseFolder of sortedFolders) {
                const samples = casesData[caseFolder];
                const baseCode = caseFolder.slice(0, 5);
                if (!caseGroups[baseCode]) {
                    caseGroups[baseCode] = [];
                }
                caseGroups[baseCode].push(...samples);
            }

            console.log('Processing cases in order:', Object.keys(caseGroups));

            for (const baseCode of Object.keys(caseGroups)) {
                const samples = caseGroups[baseCode];
                console.log(`Processing case ${baseCode} with ${samples.length} samples`);

                // Clear any existing data for this case
                const existingRow = await this.findExistingCase(baseCode, sheetName);
                if (existingRow) {
                    console.log(`Deleting existing case at row ${existingRow}`);
                    await this.deleteRows(existingRow, 28, sheetName);
                    await new Promise(resolve => setTimeout(resolve, 1000));
                }

                // Sort all samples for this case
                samples.sort((a, b) => {
                    const suffixA = a.code.slice(-1);
                    const suffixB = b.code.slice(-1);
                    return suffixA.localeCompare(suffixB);
                });

                const values = [];
                const numSamples = samples.length;

                // Add marker header row
                const markerHeaderRow = [];
                samples.forEach((_, index) => {
                    markerHeaderRow.push('Marker', 'Allele 1', 'Allele 2');
                });
                values.push(markerHeaderRow);

                // Add sample codes and names row
                const headerRow = [];
                samples.forEach(sample => {
                    headerRow.push(`${sample.code} ${sample.name}`, '', '');
                });
                values.unshift(headerRow);  // Add at the beginning

                // Add marker data rows
                MarkerComparison.MARKER_ORDER.forEach(marker => {
                    const row = [];
                    samples.forEach((sample, index) => {
                        const values = sample.markers[marker] || ['', ''];
                        const formattedValues = values.map(v => {
                            const num = parseFloat(v);
                            return !isNaN(num) ? `'${v}` : v;
                        });
                        row.push(marker, formattedValues[0] || '', formattedValues[1] || '');
                    });
                    values.push(row);
                });

                // Add empty rows between cases
                values.push(Array(numSamples * 3).fill(''));
                values.push(Array(numSamples * 3).fill(''));

                try {
                    await this.sheetsApi.spreadsheets.values.append({
                        spreadsheetId: this.spreadsheetId,
                        range: `${sheetName}!A1`,
                        valueInputOption: 'USER_ENTERED',
                        insertDataOption: 'INSERT_ROWS',
                        resource: {
                            values,
                            majorDimension: 'ROWS'
                        }
                    });
                    console.log(`Successfully uploaded case ${baseCode}`);
                } catch (uploadError) {
                    console.error('Upload error:', uploadError);
                    throw uploadError;
                }

                await new Promise(resolve => setTimeout(resolve, 2000));
            }

            return true;
        } catch (error) {
            console.error('Error uploading cases data:', error);
            throw error;
        }
    }

    formatCaseDataForUpload(samples) {
        try {
            if (!Array.isArray(samples) || samples.length !== 2) {
                throw new Error(`Invalid samples array: expected 2 samples, got ${samples?.length}`);
            }

            const values = [];
            const markerOrder = MarkerComparison.MARKER_ORDER;

            // First row: Sample names with codes
            values.push([
                `${samples[0].code} ${samples[0].name}`, '', '',
                `${samples[1].code} ${samples[1].name}`, '', ''
            ]);

            // Second row: Marker headers
            values.push(['Marker', 'Allele 1', 'Allele 2', 'Marker', 'Allele 1', 'Allele 2']);

            // Add marker data rows
            markerOrder.forEach(marker => {
                const sample1Values = samples[0].markers[marker] || ['', ''];
                const sample2Values = samples[1].markers[marker] || ['', ''];
                const row = [
                    marker,
                    sample1Values[0] || '',
                    sample1Values[1] || '',
                    marker,
                    sample2Values[0] || '',
                    sample2Values[1] || ''
                ];
                values.push(row);
            });

            // Add empty rows for spacing
            values.push(Array(6).fill(''));
            values.push(Array(6).fill(''));

            return values;
        } catch (error) {
            console.error('Error formatting case data:', error);
            throw error;
        }
    }
}

module.exports = { DataHandler }; 