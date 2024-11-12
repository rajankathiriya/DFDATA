// const express = require('express');
// const multer = require('multer');
// const uploadController = require('../controllers/uploadController');

// const router = express.Router();
// const upload = multer({ dest: 'uploads/' });

// router.post('/upload', upload.single('file'), uploadController.uploadCSV);

// module.exports = router;


const express = require('express');
const multer = require('multer');
const uploadController = require('../controllers/uploadController');
const uploadControllerCompany = require('../controllers/uploadControllerCompany');
const path = require('path');

const router = express.Router();
const upload = multer({ dest: 'uploads/' });

router.get('/gst-upload', (req, res) => {
    res.sendFile(path.join(__dirname, '../../public/gst-upload.html'));
});
console.log(__dirname);


router.get('/company-upload', (req, res) => {
    res.sendFile(path.join(__dirname, '../../public/company-upload.html'));
});

router.post('/upload/gst', upload.single('file'), uploadController.uploadCSV);
router.post('/upload/company', upload.single('file'), uploadControllerCompany.uploadCSV);

module.exports = router;
