const xlsx = require("xlsx");
const db = require("../config/db"); // Ensure the database connection is correctly configured

// Utility function to format dates
const formatDate = (dateString) => {
    if (!dateString) return null;

    if (!isNaN(dateString)) {
        const excelEpoch = new Date(1900, 0, 1);
        const date = new Date(excelEpoch.getTime() + (dateString - 1) * 24 * 60 * 60 * 1000);
        return date.toISOString().split("T")[0];
    }

    dateString = dateString.toString();

    if (dateString.includes("/")) {
        const parts = dateString.split("/");
        if (parts.length === 3) {
            const [part1, part2, year] = parts;
            const day = part1.length === 2 && parseInt(part1) <= 31 ? part1 : part2;
            const month = part2.length === 2 && parseInt(part2) <= 12 ? part2 : part1;
            return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
        }
    } else if (dateString.includes("-") && dateString.length === 10) {
        return dateString;
    }

    console.warn(`Date format unrecognized or missing for dateString: ${dateString}`);
    return null;
};

// Function to filter out existing records based on CIN or DIN
const filterExistingRecords = (data) => {
    return new Promise((resolve, reject) => {
        const cinDinPairs = data.map(row => [row['CIN'], row['DIN']]);
        const query = `SELECT CIN, DIN FROM mst_company WHERE (CIN, DIN) IN (?)`;

        db.query(query, [cinDinPairs], (err, results) => {
            if (err) return reject(err);

            const existingRecords = new Set(results.map(row => `${row.CIN}-${row.DIN}`));
            const filteredData = data.filter(row => !existingRecords.has(`${row['CIN']}-${row['DIN']}`));
            resolve(filteredData);
        });
    });
};

// Function to insert data in bulk into the database
const insertDataInBulk = (data) => {
    return new Promise((resolve, reject) => {
        const predefinedColumns = [
            'Company', 'CIN', 'DATE OF REGISTRATION', 'DIN', 'DIRECTOR NAME',
            'DESIGNATION', 'Date Of Birth', 'Mobile', 'Email', 'Gender', 'PINCODE',
            'City', 'State', 'COUNTRY', 'ROC', 'CATEGORY', 'CLASS', 'SUBCATEGORY',
            'AUTHORIZED CAPITAL', 'PAIDUP CAPITAL', 'ACTIVITY DESCRIPTION', 'DATE JOIN',
            'Registered Office Address', 'TYPE COMPANY'
        ];

        const standardizedData = data.map(row => predefinedColumns.map(col => row[col] ?? null));

        const query = `INSERT INTO mst_company (${predefinedColumns.map(col => `\`${col}\``).join(', ')}) VALUES ?`;
        db.query(query, [standardizedData], (err, result) => {
            if (err) return reject(err);
            resolve(result);
        });
    });
};

// Function to process the uploaded Excel file
exports.processExcel = (filePath) => {
    return new Promise(async (resolve, reject) => {
        try {
            console.log("Reading Excel file...");
            const workbook = xlsx.readFile(filePath);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const data = xlsx.utils.sheet_to_json(worksheet);

            console.log("Processing data from Excel...");

            // Reformat the "Registration Date" and "Date Join" columns to yyyy-mm-dd format
            data.forEach(row => {
                if (row['DATE OF REGISTRATION']) {
                    row['DATE OF REGISTRATION'] = formatDate(row['DATE OF REGISTRATION']);
                }
                if (row['DATE JOIN']) {
                    row['DATE JOIN'] = formatDate(row['DATE JOIN']);
                }
            });

            const batchSize = 500; // Size of each batch to insert
            let results = [];

            // Process the data in batches
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

            // Insert the remaining rows if any
            if (results.length > 0) {
                const filteredData = await filterExistingRecords(results);
                if (filteredData.length > 0) {
                    await insertDataInBulk(filteredData);
                }
            }

            console.log("File uploaded successfully.");
            resolve();
        } catch (err) {
            console.error("Error processing Excel file:", err);
            reject(err);
        }
    });
};
