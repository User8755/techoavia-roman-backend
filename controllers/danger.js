const danger = require('../models/danger');

module.exports.getDanger = (req, res) => {
  danger
    .find({})
    .then((item) => res.send(item))
    // eslint-disable-next-line no-console
    .catch((err) => console.log(err));
};

module.exports.createDanger = (req, res) => {
  const { dependence, groupId, label } = req.body;
  danger
    .create({ label, groupId, dependence })
    .then((item) => res.send(item))
    .catch(() => res.status(500).send({ message: 'Произошла ошибка' }));
};
