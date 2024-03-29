const Info = require('../models/info');

module.exports.createInfo = (req, res, next) => {
  const { info } = req.body.data;
  Info.create({ info })
    .then((i) => res.status(201).send(i))
    .catch((e) => next(e));
};

module.exports.getBranch = (req, res, next) => {
  Info.find({})
    .limit(1)
    .sort({ $natural: -1 })
    .then((i) => res.send(i))
    .catch((e) => next(e));
};
