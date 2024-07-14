const Excel = require('exceljs');
const Enterprise = require('../models/enterprise');
const Value = require('../models/value');
const logs = require('../models/logs');
const NotFound = require('../errors/NotFound');
const ConflictError = require('../errors/ConflictError');
const BadRequestError = require('../errors/BadRequestError');

const workbook = new Excel.Workbook();

module.exports.updateValue = (req, res, next) => {
  workbook.xlsx
    .load(req.files.file.data)
    .then(() => {
      const worksheet = workbook.getWorksheet(1);
      const cell = (lit, num) => worksheet.getCell(lit + num);

      if (cell('AU', 1).value !== 'Отметка о выполнении') {
        next(new NotFound('Не верный файл'));
      }

      Enterprise.findById(req.params.id)
        .then((enterprise) => {
          if (!enterprise) {
            next(new NotFound('Предприятие не найдено'));
          }

          if (enterprise.owner.toString() !== req.user._id) next(new ConflictError('У Вас нет доступа'));
          if (!req.data) next();

          Value.deleteMany({ enterpriseId: req.params.id })
            .then(() => {
              req.data.forEach((item) => {
                Value.create(item)
                  .then(() => {
                    res.end();
                  })
                  .catch((e) => {
                    if (e.name === 'ValidationError') {
                      console.log(e);
                      next(new BadRequestError('Не все обязательные поля заполнены'));
                    } else {
                      next(e);
                    }
                  });
              });
            })
            .catch((e) => next(e));
          logs
            .create({
              action: `Пользователь ${req.user.name} обновил(а) записи для предприятия ${enterprise.enterprise}`,
              userId: req.user._id,
              enterpriseId: enterprise._id,
            })
            .catch((e) => next(e));
        })
        .catch((e) => next(e));
    })
    .catch((i) => next(i));
};

module.exports.newValue = (req, res, next) => {
  Enterprise.findById(req.params.id)
    .then((enterprise) => {
      if (!enterprise) {
        next(new NotFound('Предприятие не найдено'));
      }
      Value.create(req.body)
        .then((data) => {
          res.send(data);
          logs
            .create({
              action: `Пользователь ${req.user.name} создал(а) запись для предприятия ${enterprise.enterprise}`,
              userId: req.user._id,
              enterpriseId: enterprise._id,
            })
            .catch((e) => next(e));
        })
        .catch((e) => next(e));
    })
    .catch((e) => next(e));
};

module.exports.getValueEnterprise = (req, res, next) => {
  Value.count({ enterpriseId: req.params.id })
    .then((i) => {
      res.send(String(i));
    })
    .catch((e) => next(e));
};
