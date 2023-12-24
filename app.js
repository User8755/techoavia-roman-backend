/* eslint-disable no-console */
const express = require('express');

const app = express();
const mongoose = require('mongoose');
const bodyParser = require('body-parser');
const cookieParser = require('cookie-parser');

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(cookieParser());

const { PORT = 3001, MONGODB = 'mongodb://127.0.0.1:27017/test' } = process.env;

mongoose.connect(MONGODB);

const urlList = ['http://127.0.0.1:3000', 'https://tafontend.online'];

app.use((req, res, next) => {
  const { origin } = req.headers;
  const { method } = req;

  const requestHeaders = req.headers['access-control-request-headers'];
  const DEFAULT_ALLOWED_METHODS = 'GET,HEAD,PUT,PATCH,POST,DELETE';

  if (urlList.includes(origin)) {
    res.header('Access-Control-Allow-Origin', origin);
  }

  // res.header('Access-Control-Allow-Origin', '*');
  if (method === 'OPTIONS') {
    res.header('Access-Control-Allow-Methods', DEFAULT_ALLOWED_METHODS);
    res.header('Access-Control-Allow-Headers', requestHeaders);
    return res.end();
  }

  return next();
});

// const ExcelJS = require('exceljs');

// const fileName = './list.xlsx';

// const workbook = new ExcelJS.Workbook();

// workbook.xlsx
//   .readFile(fileName)
//   .then(() => {})
//   .catch((e) => console.log(e));
// workbook.xlsx.readFile(fileName).then(() => {
//   const worksheet = workbook.getWorksheet('Лист1');
//   const dobCol = worksheet.getColumn(2);

//   // dobCol.eachCell((cell) => {
//   //   // console.log(`sheet.getCell('${cell.address}').style = `);
//   //   // console.log(cell.style);
//   //   worksheet.getCell(
//   //     cell.address,
//   //   ).value = `getCell(${cell.address}).style = {${cell.style}}`;
//   // });
//   // const row = worksheet.getRow(27);
//   // row.eachCell((cell, colNumber) => {
//   //   if (cell.value !== null) {
//   //     console.log(`sheet.getCell('${cell.address}').value='${cell.value}'`);
//   //   }
//   // });
//   // row.eachCell((cell) => {
//   //   console.log(`sheet.mergeCells('${cell.model.address}', '${cell.model.master}')`)
//   // });
//   dobCol.eachCell((cell) => {
//     // console.log(`sheet.getCell('${cell.address}').style = `);
//     //console.log(cell._row);

//   });
//   console.log( worksheet)
//   workbook.xlsx.writeFile('filename1.xlsx').catch((e) => console.log(e));
// });

app.use('/users', require('./routes/user'));
app.use('/dangerGroup', require('./routes/dangerGroup'));
app.use('/danger', require('./routes/danger'));
app.use('/dangerEvent', require('./routes/dangerEvent'));

app.use((err, req, res, next) => {
  const { statusCode = 500, message } = err;
  res.status(statusCode).send({
    message: statusCode === 500 ? 'На сервере произошла ошибка' : message,
  });
  next();
});

app.listen(PORT, () => {
  console.log(`Слушаем порт ${PORT}`);
});
