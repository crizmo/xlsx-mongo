const xlsx2mongo = require('xlsx-mongo');
const mongoose = require('mongoose');
require('dotenv').config();
const path = require('path');

const filePath = path.join(__dirname, 'Test.xlsx');
const showConsoleMessages = true;

xlsx2mongo.init(filePath); // Initialize xlsx-mongo with the Excel file path

const collectionName = 'test';

mongoose.connect(process.env.MONGO_URL, { useNewUrlParser: true, useUnifiedTopology: true })
.then(async () => {
    // Import data from the Excel file to the specified collection
    xlsx2mongo.import(collectionName, showConsoleMessages).then(() => {
        mongoose.connection.close(); // Close the MongoDB connection after import
    });

    // Add data from the Excel file to the specified collection
    xlsx2mongo.add(collectionName, filePath, showConsoleMessages).then(() => {
        mongoose.connection.close(); // Close the MongoDB connection after adding
    });

    // Avoid running both import and add at the same time
})
.catch((err) => {
    console.error('Error:', err);
});
