const ExcelJS = require('exceljs');
const mongoose = require('mongoose');

let workbook, worksheet, jsonData, columnNames;
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

const ExportData = async (collectionName, filePath, showConsoleMessages = true) => {
    const YourModel = SetupSchema(collectionName);
    showConsoleMessages
        ? console.log('Connected to MongoDB') && console.time('export')
        : null;

    const docs = await YourModel.find().exec();
    const rows = docs.map((doc) => {
        const row = [];
        columnNames.forEach((columnName) => {
            row.push(doc[columnName]);
        });
        return row;
    });
    rows.unshift(columnNames);

    const newWorkbook = new ExcelJS.Workbook();
    const newWorksheet = newWorkbook.addWorksheet('Sheet1');
    newWorksheet.addRows(rows);
    await newWorkbook.xlsx.writeFile(filePath);

    showConsoleMessages
        ? console.log('Data exported successfully.') && console.timeEnd('export')
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

const DeleteData = async (collectionName, showConsoleMessages = true) => {
    const YourModel = SetupSchema(collectionName);
    showConsoleMessages
        ? console.log('Connected to MongoDB') && console.time('delete')
        : null;

    await YourModel.deleteMany({});
    showConsoleMessages
        ? console.log('Data deleted successfully.') && console.timeEnd('delete')
        : null;
};

const UpdateData = async (collectionName, updateCriteria, updateData, showConsoleMessages = true) => {
    const YourModel = SetupSchema(collectionName);
    showConsoleMessages
        ? console.log('Connected to MongoDB') && console.time('update')
        : null;

    const result = await YourModel.updateMany(updateCriteria, updateData);

    showConsoleMessages
        ? console.log(`${result.modifiedCount} document(s) updated.`) && console.timeEnd('update')
        : null;
};

const FindData = async (collectionName, findCriteria, showConsoleMessages = true) => {
    const YourModel = SetupSchema(collectionName);
    showConsoleMessages
        ? console.log('Connected to MongoDB') && console.time('find')
        : null;

    const result = await YourModel.find(findCriteria);

    showConsoleMessages
        ? console.log(`${result.length} document(s) found.`) && console.timeEnd('find')
        : null;
    return result;
};

module.exports = {
    init: SetupInit,
    import: ImportData,
    export: ExportData,
    add: AddData,
    delete: DeleteData,
    update: UpdateData,
    find: FindData
};  