// excelService.js
const xlsx = require('xlsx');
const db = require('../config/db');
const predefinedColumns = [
    'Company', 'CIN', 'DATE OF REGISTRATION', 'DIN', 'DIRECTOR NAME',
    'DESIGNATION', 'Date Of Birth', 'Mobile', 'Email', 'Gender',
    'PINCODE', 'City', 'State', 'COUNTRY', 'ROC', 'CATEGORY',
    'CLASS', 'SUBCATEGORY', 'AUTHORIZED CAPITAL', 'PAIDUP CAPITAL',
    'ACTIVITY DESCRIPTION', 'DATE JOIN', 'Registered Office Address',
    'TYPE COMPANY'
];

// Function to parse and reformat dates from dd-mm-yyyy to yyyy-mm-dd
const formatDate = (dateString) => {
    if (!dateString) return null;
    const [day, month, year] = dateString.split('-');
    return `${year}-${month}-${day}`;
};

const filterExistingRecords = (data) => {
    return new Promise((resolve, reject) => {
        const uniqueIdentifiers = data.map(row => [row.DIN]);
        const query = `SELECT DIN FROM mst_company WHERE (DIN) IN (?)`;

        db.query(query, [uniqueIdentifiers], (err, results) => {
            if (err) return reject(err);

            const existingRecords = new Set(results.map(row => `${row.DIN}`));
            const filteredData = data.filter(row => !existingRecords.has(`${row.DIN}`));
            resolve(filteredData);
        });
    });
};

const insertDataInBulk = (data) => {
    return new Promise((resolve, reject) => {
        const standardizedData = data.map(row => predefinedColumns.map(col => row[col] ?? null));

        const query = `INSERT INTO mst_company (${predefinedColumns.map(col => `\`${col}\``).join(', ')}) VALUES ?`;
        db.query(query, [standardizedData], (err, result) => {
            if (err) return reject(err);
            resolve(result);
        });
    });
};

exports.processExcel = (filePath) => {
    return new Promise(async (resolve, reject) => {
        try {
            const workbook = xlsx.readFile(filePath);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const data = xlsx.utils.sheet_to_json(worksheet);

            // Reformat date columns
            data.forEach(row => {
                if (row['DATE OF REGISTRATION']) row['DATE OF REGISTRATION'] = formatDate(row['DATE OF REGISTRATION']);
                if (row['Date Of Birth']) row['Date Of Birth'] = formatDate(row['Date Of Birth']);
                if (row['DATE JOIN']) row['DATE JOIN'] = formatDate(row['DATE JOIN']);
            });

            const batchSize = 500;
            let results = [];

            for (const row of data) {
                results.push(row);

                if (results.length >= batchSize) {
                    const filteredData = await filterExistingRecords(results);
                    if (filteredData.length > 0) {
                        await insertDataInBulk(filteredData);
                    }
                    results = [];
                }
            }

            if (results.length > 0) {
                const filteredData = await filterExistingRecords(results);
                if (filteredData.length > 0) {
                    await insertDataInBulk(filteredData);
                }
            }

            resolve();
        } catch (err) {
            reject(err);
        }
    });
};
