// eslint-disable-next-line import/no-extraneous-dependencies
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const User = require('../models/user');
const ConflictError = require('../errors/ConflictError');
const BadRequestError = require('../errors/BadRequestError');
const Unauthorized = require('../errors/Unauthorized');
const NotFoundError = require('../errors/NotFound');

const { NODE_ENV, JWT_SECRET } = process.env;

module.exports.createUsers = (req, res, next) => {
  const { name, family, email } = req.body;
  bcrypt
    .hash(req.body.password, 4) // для теста пароль 4 символа
    .then((hash) => User.create({
      name,
      family,
      email,
      password: hash,
    }))
    .then((user) => {
      res.send({
        name: user.name,
        family: user.family,
        email: user.email,
      });
    })
    .catch((err) => {
      if (err.code === 11000) {
        next(new ConflictError('Пользователь с таким email зарегистрирован'));
      } else if (err.name === 'ValidationError') {
        next(new BadRequestError('Произошла ошибка, проверьте email и пароль'));
      } else {
        next(err);
      }
    });
};

module.exports.login = (req, res, next) => {
  const { email, password } = req.body;
  User.findOne({ email }).select('+password')
    .then((user) => {
      if (!user) {
        throw new Unauthorized('Проверьте email и пароль');
      }
      return bcrypt.compare(password, user.password)
        .then((matched) => {
          const token = jwt.sign({ _id: user._id }, NODE_ENV === 'production' ? JWT_SECRET : 'dev-secret', { expiresIn: '7d' });
          if (!matched) {
            throw new Unauthorized('Проверьте email и пароль');
          }
          res.cookie('key', token, {
            maxAge: 3600000 * 24 * 7, sameSite: 'None', secure: true, httpOnly: true,
          }).end();
        });
    })
    .catch((err) => {
      next(err);
    });
};

module.exports.getUsersСurrent = (req, res, next) => {
  User.findById({ _id: req.user._id })
    .then((user) => {
      if (user) {
        res.send(user);
      } else {
        next(new NotFoundError('Пользователь с данным Id не найден'));
      }
    })
    .catch((err) => {
      if (err.name === 'CastError') {
        next(new BadRequestError('Неверный Id пользователя'));
      } else {
        next(err);
      }
    });
};
