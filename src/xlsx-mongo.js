const xlsx = require('xlsx');
const mongoose = require('mongoose');

let workbook, sheetName, worksheet, jsonData, columnNames
let schema = {}

const SetupInit = (filePath) => {
    workbook = xlsx.readFile(filePath);
    sheetName = workbook.SheetNames[0];
    worksheet = workbook.Sheets[sheetName];
    jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    columnNames = jsonData.shift();

    columnNames.forEach((columnName) => {
        schema[columnName] = String;
    });
}

const SetupSchema = (collectionName) => {
    const Schema = mongoose.Schema;
    const schemaObj = new Schema(schema);
    const Model = mongoose.model(collectionName, schemaObj) || mongoose.model(collectionName, schemaObj, new Schema({ collection: collectionName }));
    return Model;
}

const UploadData = async (collectionName, showConsoleMessages = true) => {
    const YourModel = SetupSchema(collectionName);
    showConsoleMessages
        ? console.log('Connected to MongoDB') && console.time('import')
        : null;

    if (await YourModel.countDocuments().exec() === 0) {
        showConsoleMessages
            ? console.log('No data found, inserting data...')
            : null;
        await YourModel.insertMany(
            jsonData.map((row) => {
                const doc = {};
                columnNames.forEach((columnName, index) => {
                    doc[columnName] = row[index] || null;
                });
                return doc;
            })
        );
    } else {
        showConsoleMessages
            ? console.log('Data already exists. Nothing to do.')
            : null;
    }

    showConsoleMessages
        ? console.log('Data imported successfully.') && console.timeEnd('import')
        : null;
};

const AddData = async (collectionName, filePath, showConsoleMessages = true) => {
    SetupInit(filePath);
    const YourModel = SetupSchema(collectionName);
    showConsoleMessages
        ? console.log('Connected to MongoDB') && console.time('import')
        : null;

    const existingDataCount = await YourModel.countDocuments().exec();
    if (existingDataCount === 0) {
        showConsoleMessages
            ? console.log('No data found, inserting data...')
            : null;

        await YourModel.insertMany(
            jsonData.map((row) => {
                const doc = {};
                columnNames.forEach((columnName, index) => {
                    doc[columnName] = row[index] || null;
                });
                return doc;
            })
        );
    } else {
        showConsoleMessages
            ? console.log('Data already exists, adding new data...')
            : null;
        await YourModel.collection.insertMany(
            jsonData.map((row) => {
                const doc = {};
                columnNames.forEach((columnName, index) => {
                    doc[columnName] = row[index] || null;
                });
                return doc;
            })
        );
    }

    showConsoleMessages
        ? console.log('Data added successfully to an existing collection.') && console.timeEnd('import')
        : null;
        
};


module.exports.init = SetupInit;
module.exports.upload = UploadData;
module.exports.add = AddData;
