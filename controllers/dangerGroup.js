/* eslint-disable no-console */
const dangerGroup = require('../models/dangerGroup');
const NotFoundError = require('../errors/NotFound');

module.exports.getDangerGroup = (req, res) => {
  dangerGroup
    .find({})
    .then((item) => res.send(item))
    .catch((err) => console.log(err));
};

module.exports.createDangerGroup = (req, res) => {
  const { label, dangerID } = req.body;
  dangerGroup
    .create({ label, dangerID })
    .then((item) => res.send(item))
    .catch((err) => {
      if (err.code === 11000 && err.keyPattern.label === 1) {
        res.status(409).send({ message: 'Не врено заполнено название' });
      } else if (err.code === 11000 && err.keyPattern.dangerID === 1) {
        res.status(409).send({ message: 'Не врено заполнено Id' });
      } else { console.log(err); }
    });
};

module.exports.delDangerGroup = (req, res, next) => {
  dangerGroup
    .findById(req.params.id)
    .then((data) => {
      if (!data) {
        throw new NotFoundError('не найдено');
      } else if (data._id.toString() === req.params.id) {
        dangerGroup
          .findByIdAndRemove(req.params.id)
          .then((item) => res.send(item))
          .catch((err) => next(err));
      }
    })
    .catch((err) => next(err));
};
