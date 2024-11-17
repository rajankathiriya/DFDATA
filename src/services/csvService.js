// // excelService.js
// const xlsx = require('xlsx');
// const db = require('../config/db');

// const ensureColumnsExist = async (columns) => {
//     const existingColumns = await new Promise((resolve, reject) => {
//         const query = `SHOW COLUMNS FROM mst_gst`;
//         db.query(query, (err, results) => {
//             if (err) return reject(err);
//             resolve(results.map(column => column.Field));
//         });
//     });

//     const newColumns = columns.filter(column => !existingColumns.includes(column));
//     if (newColumns.length > 0) {
//         const alterQuery = newColumns.map(col => `ADD COLUMN \`${col}\` VARCHAR(255)`).join(', ');
//         await new Promise((resolve, reject) => {
//             db.query(`ALTER TABLE mst_gst ${alterQuery}`, (err) => {
//                 if (err) return reject(err);
//                 resolve();
//             });
//         });
//     }
// };

// const filterExistingRecords = (data) => {
//     return new Promise((resolve, reject) => {
//         const uniqueIdentifiers = data.map(row => [row.PAN, row.GSTIN]);
//         const query = `SELECT PAN, GSTIN FROM mst_gst WHERE (PAN, GSTIN) IN (?)`;

//         db.query(query, [uniqueIdentifiers], (err, results) => {
//             if (err) return reject(err);

//             const existingRecords = new Set(results.map(row => `${row.PAN}-${row.GSTIN}`));
//             const filteredData = data.filter(row => !existingRecords.has(`${row.PAN}-${row.GSTIN}`));
//             resolve(filteredData);
//         });
//     });
// };

// const insertDataInBulk = (data) => {
//     return new Promise(async (resolve, reject) => {
//         const columns = [...new Set(data.flatMap(Object.keys))];
//         const standardizedData = data.map(row => columns.map(col => row[col] ?? null));

//         const query = `INSERT INTO mst_gst (${columns.map(col => `\`${col}\``).join(', ')}) VALUES ?`;
//         db.query(query, [standardizedData], (err, result) => {
//             if (err) return reject(err);
//             resolve(result);
//         });
//     });
// };

// exports.processExcel = (filePath) => {
//     return new Promise(async (resolve, reject) => {
//         try {
//             const workbook = xlsx.readFile(filePath);
//             const sheetName = workbook.SheetNames[0];
//             const worksheet = workbook.Sheets[sheetName];
//             const data = xlsx.utils.sheet_to_json(worksheet);

//             const batchSize = 500;
//             let results = [];

//             for (const row of data) {
//                 results.push(row);

//                 if (results.length >= batchSize) {
//                     await ensureColumnsExist(Object.keys(row));
//                     const filteredData = await filterExistingRecords(results);
//                     if (filteredData.length > 0) {
//                         await insertDataInBulk(filteredData);
//                     }
//                     results = [];
//                 }
//             }

//             if (results.length > 0) {
//                 await ensureColumnsExist(Object.keys(results[0]));
//                 const filteredData = await filterExistingRecords(results);
//                 if (filteredData.length > 0) {
//                     await insertDataInBulk(filteredData);
//                 }
//             }

//             resolve();
//         } catch (err) {
//             reject(err);
//         }
//     });
// };





// // excelService.js
// const xlsx = require('xlsx');
// const db = require('../config/db');

// const predefinedColumns = [
//     'GSTIN', 'Registration Date', 'PAN', 'Mobile', 'Email',
//     'Legal Name', 'Trade Name', 'Business Constitution',
//     'Business Nature', 'Pincode', 'Address'
// ];

// // Function to parse and reformat dates from dd-mm-yyyy to yyyy-mm-dd
// const formatDate = (dateString) => {
//     if (!dateString) return null;
//     const [day, month, year] = dateString.split('-');
//     return `${year}-${month}-${day}`;
// };

// const filterExistingRecords = (data) => {
//     return new Promise((resolve, reject) => {
//         const uniqueIdentifiers = data.map(row => [row.PAN, row.GSTIN]);
//         const query = `SELECT PAN, GSTIN FROM mst_gst WHERE (PAN, GSTIN) IN (?)`;

//         db.query(query, [uniqueIdentifiers], (err, results) => {
//             if (err) return reject(err);

//             const existingRecords = new Set(results.map(row => `${row.PAN}-${row.GSTIN}`));
//             const filteredData = data.filter(row => !existingRecords.has(`${row.PAN}-${row.GSTIN}`));
//             resolve(filteredData);
//         });
//     });
// };

// const insertDataInBulk = (data) => {
//     return new Promise((resolve, reject) => {
//         const standardizedData = data.map(row => predefinedColumns.map(col => row[col] ?? null));

//         const query = `INSERT INTO mst_gst (${predefinedColumns.map(col => `\`${col}\``).join(', ')}) VALUES ?`;
//         db.query(query, [standardizedData], (err, result) => {
//             if (err) return reject(err);
//             resolve(result);
//         });
//     });
// };

// exports.processExcel = (filePath) => {
//     return new Promise(async (resolve, reject) => {
//         try {
//             const workbook = xlsx.readFile(filePath);
//             const sheetName = workbook.SheetNames[0];
//             const worksheet = workbook.Sheets[sheetName];
//             const data = xlsx.utils.sheet_to_json(worksheet);

//             // Reformat the "Registration Date" column to yyyy-mm-dd format
//             data.forEach(row => {
//                 if (row['Registration Date']) {
//                     row['Registration Date'] = formatDate(row['Registration Date']);
//                 }
//             });

//             const batchSize = 500;
//             let results = [];

//             for (const row of data) {
//                 results.push(row);

//                 if (results.length >= batchSize) {
//                     const filteredData = await filterExistingRecords(results);
//                     if (filteredData.length > 0) {
//                         await insertDataInBulk(filteredData);
//                     }
//                     results = [];
//                 }
//             }

//             if (results.length > 0) {
//                 const filteredData = await filterExistingRecords(results);
//                 if (filteredData.length > 0) {
//                     await insertDataInBulk(filteredData);
//                 }
//             }

//             resolve();
//         } catch (err) {
//             reject(err);
//         }
//     });
// };


const xlsx = require('xlsx');
const db = require('../config/db');

const predefinedColumns = [
    'GSTIN', 'Registration Date', 'PAN', 'Mobile', 'Email',
    'Legal Name', 'Trade Name', 'Business Constitution',
    'Business Nature', 'Pincode', 'Address'
];

// Function to format the date
const formatDate = (dateString) => {
    if (!dateString) return null;

    // Handle Excel serial date numbers
    if (!isNaN(dateString)) {
        const excelEpoch = new Date(1900, 0, 1); // Excel starts at 1900-01-01
        const date = new Date(excelEpoch.getTime() + (dateString - 2) * 24 * 60 * 60 * 1000); // Adjust for leap year bug
        return date.toISOString().split('T')[0];
    }

    // Handle "DD-MM-YYYY" format
    if (/^\d{2}-\d{2}-\d{4}$/.test(dateString)) {
        const [day, month, year] = dateString.split('-');
        return `${year}-${month}-${day}`; // Convert to YYYY-MM-DD
    }

    // Return if already in YYYY-MM-DD format
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateString)) {
        return dateString;
    }

    console.warn(`Unrecognized date format: ${dateString}`);
    return null;
};

// Filter records where both GSTIN and PAN exist
const filterExistingRecords = (data) => {
    return new Promise((resolve, reject) => {
        const gstinPanPairs = data.map(row => [row['GSTIN']]);
        const query = `SELECT GSTIN FROM mst_gst WHERE (GSTIN) IN (?)`;

        db.query(query, [gstinPanPairs], (err, results) => {
            if (err) return reject(err);

            const existingRecords = new Set(results.map(row => `${row.GSTIN}`));
            const filteredData = data.filter(row => !existingRecords.has(`${row['GSTIN']}`));
            resolve(filteredData);
        });
    });
};

// Insert data into the database in bulk
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

// Process the uploaded Excel file
exports.processExcel = (filePath) => {
    return new Promise(async (resolve, reject) => {
        try {
            const workbook = xlsx.readFile(filePath);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const data = xlsx.utils.sheet_to_json(worksheet);


            // Reformat the "Registration Date" column to yyyy-mm-dd format
            data.forEach(row => {
                if (row['Registration Date']) {
                    row['Registration Date'] = formatDate(row['Registration Date']);
                }
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
