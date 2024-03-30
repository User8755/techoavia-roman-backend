const logs = require('../models/logs');

module.exports.getLogs = (req, res, next) => {
  logs.find().limit(10).sort({ $natural: -1 })
    .then((l) => res.send(l))
    .catch((e) => next(e));
};
