/* eslint-disable no-console */
const express = require('express');

const app = express();
const cors = require('cors');
const mongoose = require('mongoose');
const bodyParser = require('body-parser');
const cookieParser = require('cookie-parser');
const fileUpload = require('express-fileupload');

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(cookieParser());
app.use(fileUpload());

const { PORT = 3001, MONGODB = 'mongodb://127.0.0.1:27017/test' } = process.env;

try {
  mongoose.connect(MONGODB);
  console.log('успех');
} catch (e) {
  console.log(e);
}

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
  })
);

app.use('/users', require('./routes/user'));
// app.use('/dangerGroup', require('./routes/dangerGroup'));
// app.use('/danger', require('./routes/danger'));
// app.use('/dangerEvent', require('./routes/dangerEvent'));
app.use('/branch', require('./routes/branch'));
app.use('/info', require('./routes/info'));
app.use('/enterprise', require('./routes/enterprise'));
app.use('/tabels', require('./routes/tabels'));
app.use('/value', require('./routes/enterpriceValue'));
app.use('/logs', require('./routes/logs'));
app.use('/update', require('./routes/update'));
app.use('/work-place', require('./routes/workPlace'));
app.use('/data', require('./routes/proff767'));

app.use((err, req, res, next) => {
  const { statusCode = 500, message } = err;
  console.log(err);
  res.status(statusCode).send({
    message:
      statusCode === 500 ? `На сервере произошла ошибка ${err}` : message,
  });
  next();
});

app.listen(PORT, () => {
  console.log(`Слушаем порт ${PORT}`);
});
