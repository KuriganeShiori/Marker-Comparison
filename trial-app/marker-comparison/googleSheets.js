const { google } = require('googleapis');

class GoogleSheetsDB {
    constructor() {
        // Initialize with your Google Sheets credentials
        this.auth = new google.auth.GoogleAuth({
            keyFile: 'path/to/credentials.json',
            scopes: ['https://www.googleapis.com/auth/spreadsheets'],
        });
        this.sheetsApi = google.sheets({ version: 'v4', auth: this.auth });
        this.spreadsheetId = 'YOUR_SPREADSHEET_ID';
    }

    async getData() {
        try {
            const response = await this.sheetsApi.spreadsheets.values.get({
                spreadsheetId: this.spreadsheetId,
                range: 'Sheet1!A:Z', // Adjust range as needed
            });
            return response.data.values;
        } catch (error) {
            console.error('Error fetching data:', error);
            throw error;
        }
    }

    async insertData(values) {
        try {
            await this.sheetsApi.spreadsheets.values.append({
                spreadsheetId: this.spreadsheetId,
                range: 'Sheet1!A:Z',
                valueInputOption: 'USER_ENTERED',
                resource: { values: [values] },
            });
        } catch (error) {
            console.error('Error inserting data:', error);
            throw error;
        }
    }

    async updateData(range, values) {
        try {
            await this.sheetsApi.spreadsheets.values.update({
                spreadsheetId: this.spreadsheetId,
                range,
                valueInputOption: 'USER_ENTERED',
                resource: { values: [values] },
            });
        } catch (error) {
            console.error('Error updating data:', error);
            throw error;
        }
    }
} 