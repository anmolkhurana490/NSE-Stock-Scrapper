import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';

import dotenv from 'dotenv';
dotenv.config();

const serviceAccountAuth = new JWT({
    email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    key: process.env.GOOGLE_PRIVATE_KEY,
    scopes: [
        'https://www.googleapis.com/auth/spreadsheets'
    ],
});

export async function saveDataToSheet(data, sheetId) {
    try {
        const doc = new GoogleSpreadsheet(sheetId, serviceAccountAuth);
        await doc.loadInfo();

        const sheet = doc.sheetsByIndex[0]; // Assuming you want to write to the first sheet
        await sheet.clear(); // Clear existing data

        // Set headers
        await sheet.setHeaderRow(['name', 'strength', 'weakness', 'opportunity', 'threats', 'mc_score']);

        // Add rows
        await sheet.addRows(data);

        console.log('Data saved successfully!');
    } catch (error) {
        console.error('Error saving data:', error.message);
    }
}