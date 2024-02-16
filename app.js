/* eslint-disable no-console */
const express = require('express');

const app = express();
const cors = require('cors');
const mongoose = require('mongoose');
const bodyParser = require('body-parser');
const cookieParser = require('cookie-parser');
const Excel = require('exceljs');
const fileUpload = require('express-fileupload');

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(cookieParser());
app.use(fileUpload());

const { PORT = 3001, MONGODB = 'mongodb://127.0.0.1:27017/test' } = process.env;

mongoose.connect(MONGODB);

const urlList = [
  'http://localhost:3000',
  'https://tafontend.online',
  'http://tafontend.online/',
];
app.use(
  cors({
    origin: urlList,
    credentials: true,
    secure: true,
  }),
);

const workbook = new Excel.Workbook();
app.post('/', (req, res) => {
  workbook.xlsx.load(req.files.file.data).then(() => {
    const worksheet = workbook.getWorksheet(1);
    const cell = (lit, num) => worksheet.getCell(lit + num);
    console.log(cell('A', 1).value);
    res.send({ 1: cell('A', 1).value });
  });
});

app.use('/users', require('./routes/user'));
// app.use('/dangerGroup', require('./routes/dangerGroup'));
// app.use('/danger', require('./routes/danger'));
// app.use('/dangerEvent', require('./routes/dangerEvent'));
app.use('/update', require('./routes/update'));
app.use('/info', require('./routes/info'));
app.use('/enterprise', require('./routes/enterprise'));
app.use('/tabels', require('./routes/tabels'));

app.use((err, req, res, next) => {
  const { statusCode = 500, message } = err;
  console.log(err);
  res.status(statusCode).send({
    message: statusCode === 500 ? 'На сервере произошла ошибка' : err,
  });
  next();
});

// const workbook = new Excel.Workbook();
// workbook.getWorksheet();
// workbook.xlsx
//   .readFile('Базовая таблица.xlsx')
//   .then(() => {
//     const worksheet = workbook.getWorksheet(1);
//     const arr = [];
//     const { lastRow } = worksheet;

//     const cell = (lit, num) => worksheet.getCell(lit + num);

//     for (let startRow = 2; startRow <= lastRow.number; startRow += 1) {
//       const obj = { SIZ: [] };
//       const siz = {};
//       if (cell('A', startRow).value) {
//         obj.type = cell('A', startRow).value;
//         arr.push(obj);
//       }
//       if (!cell('A', startRow).value) {
//         const lastObj = arr.at(-1);
//         siz.t = cell('F', startRow).value;
//         siz.a = cell('G', startRow).value;
//         lastObj.SIZ.push(siz);
//       }
//     }

//     // let excelTitles = [];
//     // const excelData = [];

//     // // excel to json converter (only the first sheet)
//     // workbook.worksheets[0].eachRow((row, rowNumber) => {
//     //   // rowNumber 0 is empty
//     //   if (rowNumber > 0) {
//     //     // get values from row
//     //     const rowValues = row.values;
//     //     // remove first element (extra without reason)
//     //     rowValues.shift();
//     //     // titles row
//     //     if (rowNumber === 1) excelTitles = rowValues;
//     //     // table data
//     //     else {
//     //       // create object with the titles and the row values (if any)
//     //       const rowObject = {};
//     //       for (let i = 0; i < excelTitles.length; i++) {
//     //         const title = excelTitles[i];
//     //         const value = rowValues[i] ? rowValues[i] : '';
//     //         rowObject[title] = value;
//     //       }
//     //       excelData.push(rowObject);
//     //     }
//     //   }
//     // });
//     // console.log(excelData);
//   })
//   .catch();

app.listen(PORT, () => {
  console.log(`Слушаем порт ${PORT}`);
});
