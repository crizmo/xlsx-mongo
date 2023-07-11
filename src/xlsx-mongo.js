const ExcelJS = require('exceljs');
const mongoose = require('mongoose');

let workbook, worksheet, jsonData, columnNames ;
let schema = {}

const SetupInit = async (filePath) => {
    workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    worksheet = workbook.worksheets[0];
    jsonData = [];

    worksheet.eachRow((row) => {
        const rowData = [];
        row.eachCell((cell) => {
            rowData.push(cell.value);
        });
        jsonData.push(rowData);
    });

    columnNames = jsonData.shift();
    columnNames.forEach((columnName) => {
        schema[columnName] = String;
    });
};


const SetupSchema = (collectionName) => {
    let YourModel;
    try {
        YourModel = mongoose.model(collectionName);
    } catch (error) {
        const Schema = mongoose.Schema;
        const schemaObj = new Schema(schema);
        YourModel = mongoose.model(collectionName, schemaObj);
    }
    return YourModel;
};

const ImportData = async (collectionName, showConsoleMessages = true) => {
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
    await SetupInit(filePath);
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
module.exports.import = ImportData;
module.exports.add = AddData;
