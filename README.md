<a name="readme-top"></a>

<br />
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
    Import / Add data from xlsx file to mongodb
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
    

<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
    </li>
    <li>
      <a href="#getting-started">Getting Started</a>
      <ul>
        <li><a href="#functions">Functions</a></li>
        <li><a href="#usage">Usage</a></li>
      </ul>
    </li>
    <li><a href="#roadmap">Roadmap</a></li>
    <li><a href="#contributing">Contributing</a></li>
    <li><a href="#license">License</a></li>
    <li><a href="#contact">Contact</a></li>
  </ol>
</details>


<!-- ABOUT THE PROJECT -->
## About The Project

Import / Add data from xlsx file to mongodb
- How it works <br>
Excel file is converted to json and then the json is imported to mongodb. <br>
Importing to mongodb is done using mongoose. <br>
- Why use it <br>
It is useful when you have a lot of data in excel file and you want to import it to mongodb. <br>
- How to use it <br>
Check the usage section for more info. <br>

<p align="right">(<a href="#readme-top">back to top</a>)</p>

## Getting Started

1. Make a mongodb database
2. Copy the connection string of the database
3. Paste the connection string in the .env file named `MONGO_URL`
4. Run the desired function
5. Check env_example file for more info - <a href="/tests/.env_example">env_example</a>
6. Install the required packages - `mongoose , exceljs , dotenv , path`
7. Install the package <br>
   ```sh
   npm install xlsx-mongo
   ```
8. Require the package in your main file <br>
   ```JS
   const xlsxMongo = require('xlsx-mongo');
   ```
<p align="right">(<a href="#readme-top">back to top</a>)</p>

## Functions

```javascript

const xlsx2mongo = require('xlsx-mongo');
const mongoose = require('mongoose');
require('dotenv').config()

const path = require('path');
const filePath = path.join(__dirname, 'Test2.xlsx'); // Path to the excel file

xlsx2mongo.init(filePath);
// Init function is used to setup the file path

const collectionName = 'test';
// Collection name is the name of the collection in which you want to import the data

// Set MONGO_URL in .env file to your mongodb connection string
mongoose.connect(process.env.MONGO_URL, { useNewUrlParser: true, useUnifiedTopology: true }) 
.then(async () => {
    
    xlsx2mongo.import(collectionName, showConsoleMessages)
    .then(() => {
        mongoose.connection.close();
    }); // you can remove the .then() if you don't want to close the connection
    
    // Import function imports the entire excel file to the database

    xlsx2mongo.add(collectionName, filePath, showConsoleMessages)
    .then(() => {
        mongoose.connection.close();
    }); // you can remove the .then() if you don't want to close the connection
    
    // Add function adds the data from the excel file to the database
    
})
.catch((err) => {
    console.error('Error:', err);
});

```
Check env_example file for more info - <a href="/tests/.env_example">env_example</a>

<p align="right">(<a href="#readme-top">back to top</a>)</p>

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
    
    xlsx2mongo.import(collectionName, showConsoleMessages)
    .then(() => {
        mongoose.connection.close();
    }); // you can remove the .then() if you don't want to close the connection

    xlsx2mongo.add(collectionName, filePath, showConsoleMessages)
    .then(() => {
        mongoose.connection.close();
    }); // you can remove the .then() if you don't want to close the connection
})
.catch((err) => {
    console.error('Error:', err);
});
```

Note: 
1. Remember to set the MONGO_URL in .env file
2. The collection name is the name of the collection in which you want to import the data
3. The file path is the path to the excel file
4. The showConsoleMessages is a boolean value which is used to show the console messages or not
5. The init function is used to setup the file path
6. The import function imports the entire excel file to the database 
[ It will add duplicate data if the data already exists in the database ]
7. The schema of the data is generated dynamically based on the first row of the excel file 
[ The first row of the excel file should contain the names of the columns ]

<p align="right">(<a href="#readme-top">back to top</a>)</p>

## For more information on how to use it visit

- [Github](https://github.com/crizmo/xlsx-mongo)
- [Example](https://github.com/crizmo/xlsx-mongo/tree/main/tests)

<p align="right">(<a href="#readme-top">back to top</a>)</p>

## Roadmap

- [ ] Add Changelog
- [ ] Add Tests
- [ ] Add Additional Templates w/ Examples
- [ ] Add More Functions
- [ ] Documentation
    - [ ] Website
    - [ ] Examples
    - [ ] Wiki
- [ ] Add Support for more platforms

See the [open issues](https://github.com/crizmo/xlsx-mongo/issues) for a full list of proposed features (and known issues).

<p align="right">(<a href="#readme-top">back to top</a>)</p>

## Contributing

Contributions are what make the open source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".
Don't forget to give the project a star! Thanks again!

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

<p align="right">(<a href="#readme-top">back to top</a>)</p>


<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE.txt` for more information.
<p align="right">(<a href="#readme-top">back to top</a>)</p>

## Contact
Package Made by: `kurizu.taz` on discord <br>
Github - [https://github.com/crizmo/xlsx-mongo](https://github.com/crizmo/xlsx-mongo)