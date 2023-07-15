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
        ? console.log('Connected to MongoDB')
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
        ? console.log('Data imported successfully.')
        : null;
};

const ExportData = async (collectionName, filePath, showConsoleMessages = true) => {
    const YourModel = SetupSchema(collectionName);
    showConsoleMessages
        ? console.log('Connected to MongoDB')
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
        ? console.log('Data exported successfully.')
        : null;
};

const AddData = async (collectionName, filePath, showConsoleMessages = true) => {
    await SetupInit(filePath);
    const YourModel = SetupSchema(collectionName);
    showConsoleMessages
        ? console.log('Connected to MongoDB')
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
        ? console.log('Data added successfully to an existing collection.')
        : null;
};

const InsertData = async (collectionName, insertData, showConsoleMessages = true) => {
    const YourModel = SetupSchema(collectionName);
    showConsoleMessages
        ? console.log('Connected to MongoDB')
        : null;

    await YourModel.collection.insertOne(insertData);

    showConsoleMessages
        ? console.log('Data inserted successfully.')
        : null;
};

const DeleteData = async (collectionName, showConsoleMessages = true) => {
    const YourModel = SetupSchema(collectionName);
    showConsoleMessages
        ? console.log('Connected to MongoDB')
        : null;

    await YourModel.deleteMany({});
    showConsoleMessages
        ? console.log('Data deleted successfully.')
        : null;
};

const UpdateData = async (collectionName, updateCriteria, updateData, showConsoleMessages = true) => {
    const YourModel = SetupSchema(collectionName);
    showConsoleMessages
        ? console.log('Connected to MongoDB')
        : null;

    const result = await YourModel.updateMany(updateCriteria, updateData);

    showConsoleMessages
        ? console.log(`${result.modifiedCount} document(s) updated.`)
        : null;
};

const FindData = async (collectionName, findCriteria, showConsoleMessages = true) => {
    const YourModel = SetupSchema(collectionName);
    showConsoleMessages
        ? console.log('Connected to MongoDB')
        : null;

    const result = await YourModel.find(findCriteria);

    showConsoleMessages
        ? console.log(`${result.length} document(s) found.`)
        : null;
    return result;
};

const FindDataWithProjection = async (collectionName, findCriteria, projection, showConsoleMessages = true) => {
    const YourModel = SetupSchema(collectionName);
    showConsoleMessages
        ? console.log('Connected to MongoDB')
        : null;
        
    const result = await YourModel.find(findCriteria, projection);

    showConsoleMessages
        ? console.log(`${result.length} document(s) found.`)
        : null;
    return result;
};

const ReplaceData = async (collectionName, replaceFilePath, showConsoleMessages = true) => {
    await SetupInit(replaceFilePath);
    const YourModel = SetupSchema(collectionName);
    showConsoleMessages
        ? console.log('Connected to MongoDB')
        : null;

    await YourModel.deleteMany({});
    await YourModel.insertMany(
        jsonData.map((row) => {
            const doc = {};
            columnNames.forEach((columnName, index) => {
                doc[columnName] = row[index] || null;
            });
            return doc;
        })
    );
    
    showConsoleMessages
        ? console.log('Data replaced successfully.')
        : null;
};

module.exports = {
    init: SetupInit,
    import: ImportData,
    export: ExportData,
    add: AddData,
    insert: InsertData,
    delete: DeleteData,
    update: UpdateData,
    find: FindData,
    findWithProjection: FindDataWithProjection,
    replace: ReplaceData
};