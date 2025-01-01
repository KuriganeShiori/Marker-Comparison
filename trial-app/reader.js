const fs = require('fs');
const Papa = require('papaparse');
const XLSX = require('xlsx');

// Step 1: Read the text file
const fileContent = fs.readFileSync('data.txt', 'utf8');

// Step 2: Parse the data (assuming it's CSV formatted)
const parsedData = Papa.parse(fileContent, { header: true }).data;

// Step 3: Convert parsed data to worksheet
const worksheet = XLSX.utils.json_to_sheet(parsedData);

// Step 4: Create a workbook and save to Excel file
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
XLSX.writeFile(workbook, 'output.xlsx');

console.log('Data successfully extracted and saved to output.xlsx');
