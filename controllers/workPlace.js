const Enterprise = require('../models/enterprise');
const Value = require('../models/value');
const workPlace = require('../models/workPlace');
const BadRequestError = require('../errors/BadRequestError');

module.exports.getUniqWorkPlace = (req, res, next) => {
  workPlace
    .find({ enterpriseId: req.params.id })
    .then((i) => {
      // const arr = [];
      // i.forEach((doc) => {
      //   if (!arr.some((u) => u.num === doc.num)) {
      //     arr.push(doc);
      //   }
      // });
      res.send(i);
    })
    .catch((e) => next(e));
};

module.exports.getCurentWorkPlace = (req, res, next) => {
  Value.find({ enterpriseId: req.params.id, num: req.body.curent })
    .then((i) => {
      res.send(i);
    })
    .catch((e) => next(e));
};

module.exports.createWorkPlace = (req, res, next) => {
  try {
    Enterprise.findById(req.params.id).then((ent) => {
      if (!ent) {
        next(new NotFound('Предприятие не найдено'));
      }

      if (
        !ent.owner.toString() === req.user._id ||
        ent.access.includes(req.user._id)
      ) {
        next(new BadRequestError('У Вас нет доступа'));
      }

      if (!req.body.proff.length > 0 && !req.body.job.length > 0)
        next(new BadRequestError('Не заполнены данные о профессии'));
      req.body.enterpriseId = req.params.id;

      workPlace
        .findOne({ enterpriseId: req.params.id, num: req.body.num })
        .then((i) => {
          if (i) {
            next(new BadRequestError(`Рабочее место ${i.num} уже созданно`));
          } else {
            workPlace
              .create(req.body)
              .then((i) => res.send(i))
              .catch((e) => next(e));
          }
        });
    });
  } catch (e) {
    next(e);
  }
};
