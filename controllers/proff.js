const Proff767 = require('../models/proff767');
module.exports.getAllproff767 = (req, res, next) => {
  Proff767.find({})
    .then((i) => res.send(i))
    .catch((e) => next(e));
};
