const Branch = require('../models/branch');

module.exports.createBranch = (req, res, next) => {
  const { branch } = req.body;
  Branch.create({ branch })
    .then((i) => res.send(i))
    .catch((e) => {
      if (e.code === 11000) {
        res.status(409).send('Такая запись уже существует');
      } else {
        next(e);
      }
    });
};

module.exports.getBranch = (req, res, next) => {
  Branch.find({})
    .then((i) => res.send(i))
    .catch((e) => next(e));
};
