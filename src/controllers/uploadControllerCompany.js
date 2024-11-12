const fs = require('fs');
const csvService = require('../services/csvServiceCompany');

exports.uploadCSV = (req, res) => {
    const filePath = req.file.path;  // Assuming multer is used for file upload
    csvService.processExcel(filePath)
        .then(() => res.status(200).send("File processed successfully"))
        .catch((error) => res.status(500).send("Error processing file: " + error.message));
};
