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
app.use('/value', require('./routes/enterpriceValue'));

app.use((err, req, res, next) => {
  const { statusCode = 500, message } = err;
  console.log(err);
  res.status(statusCode).send({
    message: statusCode === 500 ? 'На сервере произошла ошибка' : message,
  });
  next();
});

app.listen(PORT, () => {
  console.log(`Слушаем порт ${PORT}`);
});
