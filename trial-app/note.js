const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// Function to recursively get all txt files in directory and subdirectories
async function getTxtFiles(dirPath) {
    const files = await fs.promises.readdir(dirPath);
    let txtFiles = [];

    for (const file of files) {
        const fullPath = path.join(dirPath, file);
        const stat = await fs.promises.stat(fullPath);

        if (stat.isDirectory()) {
            // Recursively search in subdirectories
            const subDirFiles = await getTxtFiles(fullPath);
            txtFiles = txtFiles.concat(subDirFiles);
        } else if (file.endsWith('.txt')) {
            txtFiles.push(fullPath);
        }
    }

    return txtFiles;
}

// Function to read a file and extract marker data
function extractMarkersFromFile(filePath) {
    return new Promise((resolve, reject) => {
        fs.readFile(filePath, 'utf8', (err, data) => {
            if (err) {
                return reject(err);
            }
            
            const sampleData = {
                markers: {},
                code: '',
                name: ''
            };

            // Split the file data into lines and filter out empty lines
            const lines = data.split('\n').filter(line => line.trim());
            
            console.log(`\nProcessing file: ${path.basename(filePath)}`);
            console.log(`Total lines: ${lines.length}`);

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
                    console.log(`Stored two values for ${markerName}: [${value1}, ${value2}]`);
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
                    console.log(`Stored duplicated value for ${markerName}: [${value1}, ${value1}]`);
                }
            }

            // Debug: Verify all markers are present
            const expectedMarkers = [
                'D3S1358', 'vWA', 'D16S539', 'CSF1PO', 'D6S1043',
                'Yindel', 'AMEL', 'D8S1179', 'D21S11', 'D18S51',
                'D5S818', 'D2S441', 'D19S433', 'FGA', 'D10S1248',
                'D22S1045', 'D1S1656', 'D13S317', 'D7S820', 'Penta E',
                'Penta D', 'TH01', 'D12S391', 'D2S1338', 'TPOX'
            ];

            // Log all extracted markers and their values
            console.log('\nExtracted markers:');
            Object.entries(sampleData.markers).forEach(([marker, values]) => {
                console.log(`${marker}: [${values.join(', ')}]`);
            });

            const missingMarkers = expectedMarkers.filter(marker => !sampleData.markers[marker]);
            if (missingMarkers.length > 0) {
                console.log('\nWarning: Missing markers:', missingMarkers);
            }

            console.log(`\nTotal markers extracted: ${Object.keys(sampleData.markers).length}`);

            resolve(sampleData);
        });
    });
}

// Function to save markers data to Excel file
function saveToExcel(groupedByDate, outputPath) {
    // Check if file exists and load it, or create new workbook
    let wb;
    try {
        wb = XLSX.readFile(outputPath);
    } catch (err) {
        wb = XLSX.utils.book_new();
    }

    // Process each date's data
    Object.entries(groupedByDate).forEach(([date, allSamplesData]) => {
        // Group samples by first 5 characters of their code
        const groupedSamples = {};
        allSamplesData.forEach(sample => {
            const groupKey = sample.code.substring(0, 5);
            if (!groupedSamples[groupKey]) {
                groupedSamples[groupKey] = [];
            }
            groupedSamples[groupKey].push(sample);
        });

        // Convert markers data to worksheet format
        const wsData = [];
        const columnsPerSample = 4;
        
        // List of markers in desired order
        const markersList = [
            'D3S1358', 'vWA', 'D16S539', 'CSF1PO', 'D6S1043',
            'Yindel', 'AMEL', 'D8S1179', 'D21S11', 'D18S51',
            'D5S818', 'D2S441', 'D19S433', 'FGA', 'D10S1248',
            'D22S1045', 'D1S1656', 'D13S317', 'D7S820', 'Penta E',
            'Penta D', 'TH01', 'D12S391', 'D2S1338', 'TPOX'
        ];

        // Process each group
        Object.entries(groupedSamples).forEach(([groupKey, samples]) => {
            // Add group header
            const headerRow = [];
            samples.forEach(sample => {
                headerRow.push(`${sample.code} ${sample.name}`);
                headerRow.push('');
                headerRow.push('');
                headerRow.push('');
            });
            wsData.push(headerRow);
            
            // Add subheaders for each sample in group
            const subHeaderRow = [];
            samples.forEach(() => {
                subHeaderRow.push('Marker');
                subHeaderRow.push('Allele 1');
                subHeaderRow.push('Allele 2');
                subHeaderRow.push('');
            });
            wsData.push(subHeaderRow);
            
            // Add data rows for each marker
            markersList.forEach(marker => {
                const row = [];
                samples.forEach(sample => {
                    const values = sample.markers[marker] || ['', ''];
                    row.push(marker);
                    row.push(values[0]);
                    row.push(values[1]);
                    row.push('');
                });
                wsData.push(row);
            });
            
            // Add two empty rows between groups
            wsData.push(new Array(samples.length * columnsPerSample).fill(''));
            wsData.push(new Array(samples.length * columnsPerSample).fill(''));
        });
        
        // Create or update worksheet for this date
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        
        // Set column widths
        const colWidth = 12;
        const maxSamplesInGroup = Math.max(...Object.values(groupedSamples).map(group => group.length));
        const totalColumns = maxSamplesInGroup * columnsPerSample;
        const cols = Array(totalColumns).fill({ wch: colWidth });
        ws['!cols'] = cols;

        // Add or update worksheet
        XLSX.utils.book_append_sheet(wb, ws, date);
    });
    
    // Write to file
    XLSX.writeFile(wb, outputPath);
}

// Function to process all files in a directory
async function processAllFiles(inputDir) {
    try {
        // Get all subdirectories (dates)
        const dateDirs = (await fs.promises.readdir(inputDir, { withFileTypes: true }))
            .filter(dirent => dirent.isDirectory())
            .map(dirent => dirent.name);

        console.log(`Found date directories: ${dateDirs.join(', ')}`);

        // Group data by date
        const groupedByDate = {};

        // Process each date directory
        for (const dateDir of dateDirs) {
            const datePath = path.join(inputDir, dateDir);
            const txtFiles = await getTxtFiles(datePath);
            const allSamplesData = [];

            console.log(`\nProcessing directory ${dateDir}`);
            console.log(`Found ${txtFiles.length} txt files to process`);

            // Process each file in this date directory
            for (const filePath of txtFiles) {
                try {
                    const sampleData = await extractMarkersFromFile(filePath);
                    if (sampleData.code && sampleData.name) {
                        allSamplesData.push(sampleData);
                        console.log(`Processed: ${sampleData.code} - ${sampleData.name}`);
                    } else {
                        console.log(`Skipped file (no code/name found): ${filePath}`);
                    }
                } catch (fileErr) {
                    console.error(`Error processing file ${filePath}:`, fileErr);
                }
            }

            if (allSamplesData.length > 0) {
                groupedByDate[dateDir] = allSamplesData;
            }
        }
        
        // Create output Excel file path on Desktop
        const outputPath = path.join(process.env.HOME, 'Desktop', 'marker_database.xlsx');
        
        // Save to Excel
        saveToExcel(groupedByDate, outputPath);
        
        console.log('\nData has been successfully saved to:', outputPath);
        console.log('Processed dates:', Object.keys(groupedByDate).join(', '));
        
        return groupedByDate;
    } catch (err) {
        console.error('Error processing files:', err);
    }
}

// Example usage
const inputDir = path.join(__dirname, 'input_folder'); // Path to your folder containing date folders
processAllFiles(inputDir);