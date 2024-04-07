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
  bcrypt
    .hash(req.body.password, 4) // для теста пароль 4 символа
    .then((hash) => User.create({
      name,
      family,
      email,
      password: hash,
      branch,
      login,
      role,
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
          {
            _id: user._id,
            role: user.role,
            name: `${user.name} ${user.family}`,
          },
          NODE_ENV === 'production' ? JWT_SECRET : 'dev-secret',
          { expiresIn: '8h' },
        );
        if (!matched) {
          throw new Unauthorized('Проверьте логин и пароль');
        }
        res
          .cookie('key', token, {
            sameSite: 'none',
            maxAge: 3600000,
            httpOnly: true,
            secure: true,
          })
          .send({ key: token });
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

module.exports.getAllUsers = (req, res, next) => {
  User.findById({ _id: req.user._id })
    .then((user) => {
      if (!user) {
        next(new NotFoundError('Пользователь с данным Id не найден'));
      }
      User.find({})
        .then((allUsers) => res.send(allUsers))
        .catch((e) => next(e));
    })
    .catch((e) => next(e));
};

module.exports.getAllUsersBranch = (req, res, next) => {
  User.findById({ _id: req.user._id })
    .then((user) => {
      if (user) {
        User.find({ branch: user.branch })
          .then((u) => {
            res.send(u);
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

module.exports.newUserRole = (req, res, next) => {
  const { id } = req.body;
  User.findById(req.user._id)
    .then((u) => {
      if (!u.role.includes('user')) {
        next(new ConflictError('У Вас нет доступа'));
      }

      User.findByIdAndUpdate(
        id,
        { $addToSet: { role: req.body.role } },
        { new: true, runValidators: true },
      )
        .then((user) => {
          if (!user) {
            next(new NotFoundError('Пользователь с данным Id не найден'));
          } else {
            res.send(user);
          }
        })
        .catch((e) => next(e));
    })
    .catch((e) => next(e));
};

module.exports.delUserRole = (req, res, next) => {
  const { id } = req.body;

  User.findById(req.user._id)
    .then((u) => {
      if (!u.role.includes('user')) {
        next(new ConflictError('У Вас нет доступа'));
      }

      User.findByIdAndUpdate(
        id,
        { $pull: { role: req.body.role } },
        { new: true },
      )
        .then((user) => {
          if (!user) {
            next(new NotFoundError('Пользователь с данным Id не найден'));
          } else {
            res.send(user);
          }
        })
        .catch((e) => next(e));
    })
    .catch((e) => next(e));
};
