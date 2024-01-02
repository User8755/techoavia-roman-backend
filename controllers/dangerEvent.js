/* eslint-disable no-console */
const dangerEvent = require('../models/dangerEvent');

module.exports.getDangerEvent = (req, res) => {
  dangerEvent
    .find({})
    .then((item) => res.send(item))
    .catch((err) => console.log(err));
};

module.exports.createDangerEvent = (req, res, next) => {
  const { label, groupId, dependence } = req.body;
  dangerEvent
    .create({ label, groupId, dependence })
    .then((item) => res.send(item))
    .catch((err) => {
      if (err.code === 11000 && err.keyPattern.label === 1) {
        res.status(409).send({ message: 'Не врено заполнено название' });
      } else if (err.code === 11000 && err.keyPattern.dangerID === 1) {
        res.status(409).send({ message: 'Не врено заполнено Id' });
      } else {
        next(err);
      }
    });
};
