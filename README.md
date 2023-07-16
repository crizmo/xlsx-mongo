<div align="center">
  <a href="https://github.com/crizmo/xlsx-mongo">
    <img src="https://cdn.discordapp.com/attachments/1126788880906080366/1126788914137546792/logo.png" alt="Logo" width="80" height="80">
  </a>

  <h3 align="center">XLSX-Mongo</h3>
  <p align="center">
    <a href="https://www.npmjs.com/package/xlsx-mongo"><img src="https://img.shields.io/npm/v/xlsx-mongo.svg?maxAge=3600&style=for-the-badge" alt="NPM version" /></a>
    <a href="https://www.npmjs.com/package/xlsx-mongo"><img src="https://img.shields.io/npm/dt/xlsx-mongo?style=for-the-badge" /></a>
    <a href="https://www.npmjs.com/package/xlsx-mongo"><img src="https://img.shields.io/npm/l/xlsx-mongo?style=for-the-badge" /></a>
  </p>
  <p align="center">
    Streamlined Import, Export, and CRUD Operations between XLSX and MongoDB
    <br />
    <a href="https://github.com/crizmo/xlsx-mongo"><strong>Explore the docs »</strong></a>
    <br />
    <br />
    <a href="https://github.com/crizmo/xlsx-mongo">View Demo</a>
    ·
    <a href="https://github.com/crizmo/xlsx-mongo/issues">Report Bug</a>
    ·
    <a href="https://github.com/crizmo/xlsx-mongo/issues">Request Feature</a>
  </p>
</div>
    
## About The Project

Seamless Integration for Importing, Exporting, and Manipulating Data between XLSX and MongoDB. <br>

## Getting Started

1. Make a mongodb database
2. Copy the connection string of the database
3. Paste the connection string in the .env file named `MONGO_URL`
4. Check env_example file for more info - <a href="https://github.com/crizmo/xlsx-mongo/blob/main/tests/.env_example">env_example</a>
5. Install the package <br>
   ```sh
   npm install xlsx-mongo
   ```
6. Require the package in your main file <br>
   ```JS
   const xlsxMongo = require('xlsx-mongo');
   ``

## Functions

Init function is required to be run before any other function. <br>
```javascript
xlsxMongo.init(filePath);
```

Import data from excel file to the specified mongodb collection. <br>
```javascript
xlsxMongo.import(collectionName, showConsoleMessages);
```

Export data from the specified mongodb collection to excel file. <br>
```javascript
const exportFilePath = path.join(__dirname, 'Export.xlsx');
xlsx2mongo.export(collectionName, exportFilePath, showConsoleMessages)
```

Add data from excel file to the specified mongodb collection. <br>
```javascript
xlsxMongo.add(collectionName, filePath, showConsoleMessages);
```

Insert data to the specified mongodb collection. <br>
```javascript
const insertData = { 'Name': 'Kurizu', 'Address': 'poopy' };
xlsx2mongo.insert(collectionName, insertData, showConsoleMessages)
```

Delete data from the specified mongodb collection. <br>
```javascript
xlsxMongo.delete(collectionName, showConsoleMessages);
```

Update data from the specified mongodb collection. <br>
```javascript
xlsx2mongo.update(collectionName, updateCriteria, updateData, showConsoleMessages)
```

Find data from the specified mongodb collection. <br>
```javascript
xlsx2mongo.find(collectionName, findCriteria, showConsoleMessages)
```

Find data with projection from the specified mongodb collection. <br>
```javascript
const findCriteriaPro = { 'Name': 'Kurizu' };
const projection = { 'Name': 1 };
xlsx2mongo.findWithProjection(collectionName, findCriteriaPro, projection, showConsoleMessages)
```

Replace data with new excel file <br>
```javascript
const replaceFilePath = path.join(__dirname, 'Replace.xlsx');
xlsx2mongo.replace(collectionName, replaceFilePath, showConsoleMessages)
```

Check env_example file for more info - <a href="https://github.com/crizmo/xlsx-mongo/blob/main/tests/test.js">env_example</a>

## Usage

```javascript
const xlsx2mongo = require('xlsx-mongo');
const mongoose = require('mongoose');
require('dotenv').config()
const path = require('path');

const filePath = path.join(__dirname, 'Test2.xlsx');
const showConsoleMessages = false;

xlsx2mongo.init(filePath);

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

    // Insert data to the specified collection
    const insertData = { 'Name': 'Kurizu', 'Address': 'poopy' };
    xlsx2mongo.insert(collectionName, insertData, showConsoleMessages).then(() => {
        mongoose.connection.close(); // Close the MongoDB connection after adding
    });
    // If you want to insert multiple rows, you can use .add() method instead

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

    // Find data with projection
    const findCriteriaPro = { 'Name': 'Kurizu' };
    const projection = { 'Name': 1 };
    xlsx2mongo.findWithProjection(collectionName, findCriteriaPro, projection, showConsoleMessages).then((res) => {
        console.log(res);
        mongoose.connection.close(); // Close the MongoDB connection after finding
    });

    // Replace data with new excel file
    const replaceFilePath = path.join(__dirname, 'Replace.xlsx');
    xlsx2mongo.replace(collectionName, replaceFilePath, showConsoleMessages).then(() => {
        mongoose.connection.close(); // Close the MongoDB connection after replacing
    });

    // Avoid running the above functions at the same time
    // you can remove the .then() if you don't want to close the connection
})
.catch((err) => {
    console.error('Error:', err);
});
```

## For more information on how to use it visit

- [Github](https://github.com/crizmo/xlsx-mongo)
- [Example](https://github.com/crizmo/xlsx-mongo/tree/main/tests)

## License

Distributed under the MIT License. See `LICENSE.txt` for more information.

## Contact
Package Made by: `kurizu.taz` on discord <br>
Github - [https://github.com/crizmo/xlsx-mongo](https://github.com/crizmo/xlsx-mongo)