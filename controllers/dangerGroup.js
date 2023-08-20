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
    .catch(() => res.status(500).send({ message: 'Произошла ошибка' }));
};
