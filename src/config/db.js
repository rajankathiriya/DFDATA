// const mysql = require('mysql');

// // const db = mysql.createConnection({
// //     host: 'localhost',
// //     user: 'root',
// //     password: '',
// //     database: 'gsttest',
// //     port: 3306
// // });
// const db = mysql.createConnection({
//     host: 'srv1639.hstgr.io',
//     user: 'u801738259_data',
//     password: 'Rajan@123#123',
//     database: 'u801738259_data',
//     port: 3306
// });

// db.connect((err) => {
//     if (err) throw err;
//     console.log('Connected to the database');
// });

// module.exports = db;


const mysql = require('mysql');

const db = mysql.createPool({
    host: 'srv1639.hstgr.io',
    user: 'u801738259_data',
    password: 'Rajan@123#123',
    database: 'u801738259_data',
    port: 3306,
    connectionLimit: 10,  // Pool connection limit
    connectTimeout: 0,     // No timeout for initial connection
    acquireTimeout: 0,     // No timeout for acquiring connections from pool
});

db.getConnection((err, connection) => {
    if (err) {
        console.error('Database connection failed:', err.message);
        return;
    }
    console.log('Connected to the database');
    connection.release(); // Release the connection back to the pool
});

module.exports = db;
