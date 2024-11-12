const express = require('express');
const uploadRoutes = require('./routes/uploadRoutes');
const path = require('path');

const app = express();
const port = 3000;

// Serve static files from "public" folder
app.use(express.static(path.join(__dirname, '../public')));

// Use the upload routes
app.use(uploadRoutes);

app.listen(port, () => {
    console.log(`Server running on http://localhost:${port}`);
});
