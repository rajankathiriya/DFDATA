const xlsx = require("xlsx");
const db = require("../config/db"); // Ensure the database connection is correctly configured

const formatDate = (dateString) => {
    if (!dateString) return null; // Handle empty or null values

    // If the date is a number (Excel serial date format)
    if (!isNaN(dateString)) {
        const excelEpoch = new Date(1900, 0, 1); // Start date for Excel serial dates
        const date = new Date(excelEpoch.getTime() + (dateString - 1) * 24 * 60 * 60 * 1000);
        if (isNaN(date.getTime())) return null; // Return null if date is invalid
        return date.toISOString().split("T")[0]; // Format as yyyy-mm-dd
    }

    // Handle string date in mm/dd/yyyy format
    const dateParts = dateString.split("/");
    if (dateParts.length === 3) {
        const [month, day, year] = dateParts.map(part => part.padStart(2, "0"));
        if (isValidDate(day, month, year)) {
            return `${year}-${month}-${day}`; // Return formatted date as yyyy-mm-dd
        }
    }

    // Handle string date in yyyy-mm-dd format (no change needed, already in valid format)
    const isoDateParts = dateString.split("-");
    if (isoDateParts.length === 3 && isoDateParts[0].length === 4 && isoDateParts[1].length === 2 && isoDateParts[2].length === 2) {
        return dateString; // Already in yyyy-mm-dd format
    }

    console.warn(`Unrecognized or invalid date: ${dateString}`);
    return null; // Return null for invalid dates
};

// Helper function to validate the date
const isValidDate = (day, month, year) => {
    const date = new Date(`${year}-${month}-${day}`);
    return !isNaN(date.getTime()) && date.getDate() === parseInt(day) && date.getMonth() + 1 === parseInt(month);
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
            'Company', 'CIN','CEmail', 'DATE OF REGISTRATION', 'DIN', 'DIRECTOR NAME',
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

            // Reformat the "DATE OF REGISTRATION" and "DATE JOIN" columns to dd-mm-yyyy format
            data.forEach(row => {
                if (row['DATE OF REGISTRATION']) {
                    const originalDate = row['DATE OF REGISTRATION'];
                    const formattedDate = formatDate(originalDate);
                    if (!formattedDate) {
                        console.error(`Invalid 'DATE OF REGISTRATION' value in row:`, row);
                        row['DATE OF REGISTRATION'] = null; // Set to null if invalid
                    } else {
                        row['DATE OF REGISTRATION'] = formattedDate;
                    }
                } else {
                    console.warn(`Missing 'DATE OF REGISTRATION' in row:`, row);
                    row['DATE OF REGISTRATION'] = null; // Assign null if date is missing
                }

                if (row['DATE JOIN']) {
                    const originalDate = row['DATE JOIN'];
                    const formattedDate = formatDate(originalDate);
                    if (!formattedDate) {
                        console.error(`Invalid 'DATE JOIN' value in row:`, row);
                        row['DATE JOIN'] = null; // Set to null if invalid
                    } else {
                        row['DATE JOIN'] = formattedDate;
                    }
                } else {
                    row['DATE JOIN'] = null; // Assign null if date is missing
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
