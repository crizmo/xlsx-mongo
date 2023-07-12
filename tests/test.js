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

    // // Add data from the Excel file to the specified collection
    xlsx2mongo.add(collectionName, filePath, showConsoleMessages).then(() => {
        mongoose.connection.close(); // Close the MongoDB connection after adding
    });

    // Export data from the specified collection to the Excel file
    const exportFilePath = path.join(__dirname, 'Export.xlsx');
    xlsx2mongo.export(collectionName, exportFilePath, showConsoleMessages).then(() => {
      mongoose.connection.close(); // Close the MongoDB connection after exporting
    });

    // Delete data from the specified collection
    xlsx2mongo.delete(collectionName, showConsoleMessages).then(() => {
        mongoose.connection.close(); // Close the MongoDB connection after deleting
    });

    // Update data from the specified collection
    //  to update single row
    const updateCriteria = { 'Name': 'efrwdawd' };
    const updateData = { $set: { 'Name': 'John Doe' } };
    xlsx2mongo.update(collectionName, updateCriteria, updateData, showConsoleMessages).then(() => {
        mongoose.connection.close(); // Close the MongoDB connection after updating
    }); 

    // to update multiple rows
    const updateCriteriaMultiple = { 'Name': 'dwgdrthg', 'Address': 'grgdrgd' };
    const updateDataMultiple = { $set: { 'Name': 'Kurizu', 'Address': 'poopy' } };
    xlsx2mongo.update(collectionName, updateCriteriaMultiple, updateDataMultiple, showConsoleMessages).then(() => {
        mongoose.connection.close(); // Close the MongoDB connection after updating
    });

    // Find data from the specified collection
    const findCriteria = { 'Name': 'Kurizu' };   
    xlsx2mongo.find(collectionName, findCriteria, showConsoleMessages).then((res) => {
        console.log(res);
        mongoose.connection.close(); // Close the MongoDB connection after finding
    });

    // Avoid running the above functions at the same time
})
.catch((err) => {
    console.error('Error:', err);
});
