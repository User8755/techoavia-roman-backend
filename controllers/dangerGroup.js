const dangerGroup = require('../models/dangerGroup');

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
      console.log(err), console.log(err.keyPattern.label);
      if (err.code === 11000 && err.keyPattern.label === 1) {
        res.status(409).send({ message: 'Не врено заполнено название' });
      }
      if (err.code === 11000 && err.keyPattern.dangerID === 1) {
        res.status(409).send({ message: 'Не врено заполнено Id' });
      }
    });
};
