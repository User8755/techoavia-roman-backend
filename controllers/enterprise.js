const Enterprise = require('../models/enterprise');
const ConflictError = require('../errors/ConflictError');

module.exports.createEnterprise = (req, res, next) => {
  const {
    enterprise, inn, kpp, order,
  } = req.body;
  Enterprise.create({
    enterprise,
    inn,
    kpp,
    order,
    owner: req.user._id,
  })
    .then((i) => res.send(i))
    .catch((e) => {
      if (e.code === 11000) {
        next(new ConflictError('Данный номер договра уже существует'));
      }
      next(e);
    });
};

module.exports.getEnterprisesUser = (req, res, next) => {
  Enterprise.find({ owner: req.user._id })
    .then((i) => {
      res.send(i);
    })
    .catch((e) => next(e));
};

module.exports.getEnterprisesAccessUser = (req, res, next) => {
  Enterprise.find({ access: req.user._id })
    .then((i) => {
      res.send(i);
    })
    .catch((e) => next(e));
};
module.exports.getCurrentEnterprise = (req, res, next) => {
  Enterprise.findOne({ _id: req.params.id })
    .then((i) => {
      if (
        i.owner.toString() === req.user._id
        || i.access.includes(req.user._id)
      ) {
        res.send(i);
      } else {
        next(new ConflictError('Нет доступа'));
      }
    })
    .catch((e) => next(e));
};

module.exports.updateAccess = (req, res, next) => {
  const { user } = req.body;
  Enterprise.findByIdAndUpdate(
    req.params.id,
    { $push: { access: user } },
    { new: true },
  )
    .then((i) => {
      res.send(i);
    })
    .catch((e) => next(e));
};

module.exports.updateCloseAccess = (req, res, next) => {
  const { user } = req.body;
  Enterprise.findByIdAndUpdate(
    req.params.id,
    { $pull: { access: user } },
    { new: true },
  )
    .then((i) => {
      res.send(i);
    })
    .catch((e) => next(e));
};
