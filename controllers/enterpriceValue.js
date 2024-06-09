const Excel = require('exceljs');
const Enterprise = require('../models/enterprise');
const Value = require('../models/value');
const logs = require('../models/logs');
const NotFound = require('../errors/NotFound');
const ConflictError = require('../errors/ConflictError');

const workbook = new Excel.Workbook();

module.exports.updateValue = (req, res, next) => {
  const handleStyleString = (valueStr) => {
    if (typeof valueStr === 'string') {
      const str = valueStr.charAt(0).toUpperCase() + valueStr.substr(1);
      return str.trim();
    }
    if (typeof valueStr === 'number') {
      return valueStr;
    }
    return '';
  };

  workbook.xlsx
    .load(req.files.file.data)
    .then(() => {
      const worksheet = workbook.getWorksheet(1);
      const cell = (lit, num) => worksheet.getCell(lit + num);
      const arr = [];
      const { lastRow } = worksheet;

      if (cell('AU', 1).value !== 'Отметка о выполнении') {
        next(new NotFound('Не верный файл'));
      }

      Enterprise.findById(req.params.id)
        .then((enterprise) => {
          if (!enterprise) {
            next(new NotFound('Предприятие не найдено'));
          }

          if (enterprise.owner.toString() !== req.user._id) next(new ConflictError('У Вас нет доступа'));

          for (let startRow = 2; startRow <= lastRow.number; startRow += 1) {
            const newObj = { proffSIZ: [] };
            const siz = {};
            if (cell('C', startRow).value) {
              newObj.proffId = handleStyleString(cell('B', startRow).value) || '';
              newObj.num = handleStyleString(cell('C', startRow).value) || '';
              newObj.proff = handleStyleString(cell('D', startRow).value) || '';
              newObj.job = handleStyleString(cell('E', startRow).value) || '';
              newObj.subdivision = handleStyleString(cell('F', startRow).value) || '';
              newObj.obj = handleStyleString(cell('J', startRow).value) || '';
              newObj.source = handleStyleString(cell('K', startRow).value) || '';
              newObj.dangerID = handleStyleString(cell('L', startRow).value) || '';
              newObj.danger = handleStyleString(cell('M', startRow).value) || '';
              newObj.dangerGroupId = handleStyleString(cell('N', startRow).value) || '';
              newObj.dangerGroup = handleStyleString(cell('O', startRow).value) || '';
              newObj.dangerEventID = handleStyleString(cell('P', startRow).value) || '';
              newObj.dangerEvent = handleStyleString(cell('Q', startRow).value) || '';
              newObj.heaviness = handleStyleString(cell('R', startRow).value) || '';
              newObj.probability = handleStyleString(cell('S', startRow).value) || '';
              newObj.ipr = handleStyleString(cell('T', startRow).value) || '';
              newObj.risk = handleStyleString(cell('U', startRow).value) || '';
              newObj.acceptability = handleStyleString(cell('V', startRow).value) || '';
              newObj.riskAttitude = handleStyleString(cell('W', startRow).value) || '';
              newObj.typeSIZ = handleStyleString(cell('X', startRow).value) || '';
              newObj.speciesSIZ = handleStyleString(cell('Y', startRow).value) || '';
              newObj.issuanceRate = handleStyleString(cell('Z', startRow).value) || '';
              newObj.additionalMeans = handleStyleString(cell('AA', startRow)).value || '';
              newObj.AdditionalIssuanceRate = handleStyleString(cell('AB', startRow).value) || '';
              newObj.standart = handleStyleString(cell('AC', startRow).value) || '';
              newObj.OperatingLevel = handleStyleString(cell('AD', startRow).value) || '';
              newObj.commit = handleStyleString(cell('AE', startRow).value) || '';
              newObj.danger776Id = handleStyleString(cell('AF', startRow).value) || '';
              newObj.danger776 = handleStyleString(cell('AG', startRow).value) || '';
              newObj.dangerEvent776Id = handleStyleString(cell('AH', startRow).value) || '';
              newObj.dangerEvent776 = handleStyleString(cell('AI', startRow).value) || '';
              newObj.riskManagementID = handleStyleString(cell('AJ', startRow).value) || '';
              newObj.riskManagement = handleStyleString(cell('AK', startRow).value) || '';
              newObj.heaviness1 = handleStyleString(cell('AL', startRow).value) || '';
              newObj.probability1 = handleStyleString(cell('AM', startRow).value) || '';
              newObj.ipr1 = handleStyleString(cell('AN', startRow).value) || '';
              newObj.risk1 = handleStyleString(cell('AO', startRow).value) || '';
              newObj.acceptability1 = handleStyleString(cell('AP', startRow).value) || '';
              newObj.riskAttitude1 = handleStyleString(cell('AQ', startRow).value) || '';
              newObj.existingRiskManagement = handleStyleString(cell('AR', startRow).value) || '';
              newObj.periodicity = handleStyleString(cell('AS', startRow).value) || '';
              newObj.responsiblePerson = handleStyleString(cell('AT', startRow).value) || '';
              newObj.completionMark = handleStyleString(cell('AU', startRow).value) || '';
              newObj.numWorkers = handleStyleString(cell('AV', startRow).value) || '';
              newObj.equipment = handleStyleString(cell('AW', startRow).value) || '';
              newObj.materials = handleStyleString(cell('AX', startRow).value) || '';
              newObj.laborFunction = handleStyleString(cell('AY', startRow).value) || '';
              newObj.code = handleStyleString(cell('AZ', startRow).value);
              newObj.enterpriseId = req.params.id;

              arr.push(newObj);
            }

            if (!cell('A', startRow).value && cell('G', startRow).value) {
              const lastObj = arr.at(-1);
              siz.type = cell('G', startRow).value;
              siz.vid = cell('H', startRow).value;
              siz.norm = cell('I', startRow).value;
              lastObj.proffSIZ.push(siz);
            }
          }

          Value.deleteMany({ enterpriseId: req.params.id })
            .then(() => {
              arr.forEach((item) => {
                Value.create(item)
                  .then(() => {
                    res.end();
                  })
                  .catch((e) => next(e));
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
