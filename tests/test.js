const xlsx2mongo = require('../index');
const mongoose = require('mongoose');
require('dotenv').config()
const path = require('path');

const filePath = path.join(__dirname, 'Test.xlsx');
const showConsoleMessages = true;

xlsx2mongo.init(filePath);

const collectionName = 'test';

mongoose.connect(process.env.MONGO_URL, { useNewUrlParser: true, useUnifiedTopology: true })
.then(async () => {
    
    xlsx2mongo.import(collectionName, showConsoleMessages).then(() => {
        mongoose.connection.close();
    });
    
    xlsx2mongo.add(collectionName, filePath, showConsoleMessages).then(() => {
        mongoose.connection.close();
    });

    // Avoid running both import and add at the same time
})
.catch((err) => {
    console.error('Error:', err);
});