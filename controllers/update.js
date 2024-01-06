const Branch = require('../models/branch');
const BadRequestError = require('../errors/BadRequestError');
const ConflictError = require('../errors/ConflictError');

module.exports.createBranch = (req, res, next) => {
  const { branch } = req.body;
  Branch.create({ branch })
    .then((i) => res.send(i))
    .catch((e) => {
      if (e.code === 11000) {
        next(next(new ConflictError('Запись уже существует')));
      }
      if (e.name === 'ValidationError') {
        next(new BadRequestError('Недопустимая запись'));
      } else {
        next(e);
      }
    });
};

module.exports.getBranch = (req, res, next) => {
  Branch.find({})
    .then((i) => res.send(i))
    .catch((e) => next(e));
};
