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
  const {
    name, family, email, branch, login, role,
  } = req.body;
  let newrole;
  if (role === 'Администратор филиала') {
    newrole = 'admin';
  } else if (role === 'Пользователь') {
    newrole = 'user';
  } else if (role === 'Нет') {
    newrole = 'none';
  } else if (role === 'Супер') {
    newrole = 'sadmin';
  }

  bcrypt
    .hash(req.body.password, 4) // для теста пароль 4 символа
    .then((hash) => User.create({
      name,
      family,
      email,
      password: hash,
      branch,
      login,
      role: newrole,
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
        next(
          new ConflictError(
            'Пользователь с таким логином или email зарегистрирован',
          ),
        );
      } else if (err.name === 'ValidationError') {
        next(new BadRequestError('Произошла ошибка, проверьте логин и пароль'));
      } else {
        next(err);
      }
    });
};

module.exports.login = (req, res, next) => {
  const { login, password } = req.body;
  User.findOne({ login })
    .select('+password')
    .then((user) => {
      if (!user) {
        throw new Unauthorized('Проверьте логин и пароль');
      }
      return bcrypt.compare(password, user.password).then((matched) => {
        const token = jwt.sign(
          { _id: user._id },
          NODE_ENV === 'production' ? JWT_SECRET : 'dev-secret',
          { expiresIn: '7d' },
        );
        if (!matched) {
          throw new Unauthorized('Проверьте логин и пароль');
        }
        res.send({ key: token });
      });
    })
    .catch((err) => {
      next(err);
    });
};

module.exports.getUsersСurrent = (req, res, next) => {
  console.log(req.user)
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

module.exports.getAllUsers = (req, res, next) => {
  const { authorization } = req.headers;
  const token = authorization.replace('Bearer ', '');
  const payload = jwt.verify(
    token,
    NODE_ENV === 'production' ? JWT_SECRET : 'dev-secret',
  );
  User.findById({ _id: payload._id })
    .then((user) => {
      if (user.role === 'admin') {
        User.find({ branch: user.branch })
          .then((u) => {
            res.send(u);
          })
          .catch((e) => next(e));
      } else if (user.role === 'sadmin') {
        User.find({})
          .then((u) => {
            const sortedUserBranch = u.sort((a, b) => {
              const nameA = a.branch.toLowerCase();
              const nameB = b.branch.toLowerCase();
              if (nameA < nameB) return -1;
              // сортируем строки по возрастанию
              if (nameA > nameB) return 1;
              return 0; // Никакой сортировки
            });
            res.send(sortedUserBranch);
          })
          .catch((e) => next(e));
      } else {
        next(new NotFoundError('Пользователь с данным Id не найден'));
      }
    })
    .catch((e) => next(e));
};

module.exports.updateProfile = (req, res, next) => {
  const { authorization } = req.headers;
  const token = authorization.replace('Bearer ', '');
  const payload = jwt.verify(
    token,
    NODE_ENV === 'production' ? JWT_SECRET : 'dev-secret',
  );
  bcrypt
    .hash(req.body.password.password, 4) // Пофискисть это дерьмо
    .then((hash) => User.findByIdAndUpdate(
      payload._id,
      { password: hash },
      { new: true, runValidators: true },
    ))

    .then((user) => {
      if (!user) {
        next(new NotFoundError('Пользователь с данным Id не найден'));
      } else {
        res.send(user);
      }
    })
    .catch((e) => next(e));
};
