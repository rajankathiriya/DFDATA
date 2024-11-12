// excelService.js
const xlsx = require('xlsx');
const db = require('../config/db');

const ensureColumnsExist = async (columns) => {
    const existingColumns = await new Promise((resolve, reject) => {
        const query = `SHOW COLUMNS FROM mst_gst`;
        db.query(query, (err, results) => {
            if (err) return reject(err);
            resolve(results.map(column => column.Field));
        });
    });

    const newColumns = columns.filter(column => !existingColumns.includes(column));
    if (newColumns.length > 0) {
        const alterQuery = newColumns.map(col => `ADD COLUMN \`${col}\` VARCHAR(255)`).join(', ');
        await new Promise((resolve, reject) => {
            db.query(`ALTER TABLE mst_gst ${alterQuery}`, (err) => {
                if (err) return reject(err);
                resolve();
            });
        });
    }
};

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
    return new Promise(async (resolve, reject) => {
        const columns = [...new Set(data.flatMap(Object.keys))];
        const standardizedData = data.map(row => columns.map(col => row[col] ?? null));

        const query = `INSERT INTO mst_gst (${columns.map(col => `\`${col}\``).join(', ')}) VALUES ?`;
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
                    await ensureColumnsExist(Object.keys(row));
                    const filteredData = await filterExistingRecords(results);
                    if (filteredData.length > 0) {
                        await insertDataInBulk(filteredData);
                    }
                    results = [];
                }
            }

            if (results.length > 0) {
                await ensureColumnsExist(Object.keys(results[0]));
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
