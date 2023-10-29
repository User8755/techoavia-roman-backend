// eslint-disable-next-line import/no-extraneous-dependencies
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const User = require('../models/user');
const ConflictError = require('../errors/ConflictError');
const BadRequestError = require('../errors/BadRequestError');
const Unauthorized = require('../errors/Unauthorized');

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
  console.log(res);
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
          res.send({ token });
        });
    })
    .catch((err) => {
      next(err);
    });
};
