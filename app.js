/* eslint-disable no-console */
const express = require('express');

const app = express();
const cors = require('cors');
const mongoose = require('mongoose');
const bodyParser = require('body-parser');
const cookieParser = require('cookie-parser');

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(cookieParser());

const { PORT = 3001, MONGODB = 'mongodb://127.0.0.1:27017/test' } = process.env;

mongoose.connect(MONGODB);

const urlList = ['http://localhost:3000', 'https://tafontend.online', 'http://tafontend.online/'];
// app.use(
//   cors({
//     origin: urlList,
//     credentials: true,
//     secure: false,
//   }), https://tafontend.online
// );
app.use((req, res, next) => {
  const { origin } = req.headers;
  const { method } = req;
  console.log(origin);
  const requestHeaders = req.headers['access-control-request-headers'];
  const DEFAULT_ALLOWED_METHODS = 'GET,HEAD,PUT,PATCH,POST,DELETE';
  console.log(urlList.includes(origin));
  if (urlList.includes(origin)) {
    res.header('Access-Control-Allow-Origin', origin);
    res.header('Access-Control-Allow-Credentials', true);
  }

  // res.header('Access-Control-Allow-Origin', '*');
  if (method === 'OPTIONS') {
    res.header('Access-Control-Allow-Methods', DEFAULT_ALLOWED_METHODS);
    res.header('Access-Control-Allow-Headers', requestHeaders);
    return res.end();
  }

  return next();
});

app.use('/users', require('./routes/user'));
// app.use('/dangerGroup', require('./routes/dangerGroup'));
// app.use('/danger', require('./routes/danger'));
// app.use('/dangerEvent', require('./routes/dangerEvent'));
app.use('/update', require('./routes/update'));
app.use('/info', require('./routes/info'));
app.use('/enterprise', require('./routes/enterprise'));

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
