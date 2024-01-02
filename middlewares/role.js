const User = require('../models/user');

const ConflictError = require('../errors/ConflictError');

module.exports = (req, res, next) => {
  User.findById({ _id: req.query._id })
    .then((user) => {
      if (user.role === 'admin') {
        next();
      } else {
        throw new ConflictError('Недоcтаточно прав доустпа');
      }
    })
    .catch((err) => {
      next(err);
    });
};
