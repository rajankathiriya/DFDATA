// excelService.js
const xlsx = require('xlsx');
const db = require('../config/db');

const predefinedColumns = [
    'GSTIN', 'Registration Date', 'PAN', 'Mobile', 'Email',
    'Legal Name', 'Trade Name', 'Business Constitution',
    'Business Nature', 'Pincode', 'Address'
];

const filterExistingRecords = (data) => {
    return new Promise((resolve, reject) => {
        const uniqueIdentifiers = data.map(row => [row.PAN, row.GSTIN]);
        const query = `SELECT PAN, GSTIN FROM mst_gst WHERE (PAN, GSTIN) IN (?)`;

        db.query(query, [uniqueIdentifiers], (err, results) => {
            if (err) return reject(err);

            const existingRecords = new Set(results.map(row => `${row.PAN}-${row.GSTIN}`));
            const filteredData = data.filter(row => !existingRecords.has(`${row.PAN}-${row.GSTIN}`));
            resolve(filteredData);
        });
    });
};

const insertDataInBulk = (data) => {
    return new Promise((resolve, reject) => {
        const standardizedData = data.map(row => predefinedColumns.map(col => row[col] ?? null));

        const query = `INSERT INTO mst_gst (${predefinedColumns.map(col => `\`${col}\``).join(', ')}) VALUES ?`;
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
