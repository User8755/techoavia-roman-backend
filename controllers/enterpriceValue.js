const Excel = require('exceljs');
const Enterprise = require('../models/enterprise');
const Value = require('../models/value');
const NotFound = require('../errors/NotFound');

const workbook = new Excel.Workbook();

module.exports.updateValue = (req, res, next) => {
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

          for (let startRow = 2; startRow <= lastRow.number; startRow += 1) {
            const newObj = { SIZ: [] };
            const siz = {};
            if (cell('A', startRow).value) {
              // obj.type = cell('A', startRow).value;
              newObj.proffId = cell('B', startRow).value;
              newObj.num = cell('C', startRow).value;
              newObj.proff = cell('D', startRow).value;
              newObj.job = cell('E', startRow).value;
              newObj.subdivision = cell('F', startRow).value;
              newObj.obj = cell('J', startRow).value;
              newObj.source = cell('K', startRow).value;
              newObj.dangerID = cell('L', startRow).value;
              newObj.danger = cell('M', startRow).value;
              newObj.dangerGroupId = cell('N', startRow).value;
              newObj.dangerGroup = cell('O', startRow).value;
              newObj.dangerEventID = cell('P', startRow).value;
              newObj.dangerEvent = cell('Q', startRow).value;
              newObj.heaviness = cell('R', startRow).value;
              newObj.probability = cell('S', startRow).value;
              newObj.ipr = cell('T', startRow).value;
              newObj.risk = cell('U', startRow).value;
              newObj.acceptability = cell('V', startRow).value;
              newObj.riskAttitude = cell('W', startRow).value;
              newObj.typeSIZ = cell('X', startRow).value;
              newObj.speciesSIZ = cell('Y', startRow).value;
              newObj.issuanceRate = cell('Z', startRow).value;
              newObj.additionalMeans = cell('AA', startRow).value;
              newObj.AdditionalIssuanceRate = cell('AB', startRow).value;
              newObj.standart = cell('AC', startRow).value;
              newObj.OperatingLevel = cell('AD', startRow).value;
              newObj.commit = cell('AE', startRow).value;
              newObj.danger776Id = cell('AF', startRow).value;
              newObj.danger776 = cell('AG', startRow).value;
              newObj.dangerEvent776Id = cell('AH', startRow).value;
              newObj.dangerEvent776 = cell('AI', startRow).value;
              newObj.riskManagementID = cell('AJ', startRow).value;
              newObj.riskManagement = cell('AK', startRow).value;
              newObj.heaviness1 = cell('AL', startRow).value;
              newObj.probability1 = cell('AM', startRow).value;
              newObj.ipr1 = cell('AN', startRow).value;
              newObj.risk1 = cell('AO', startRow).value;
              newObj.acceptability1 = cell('AP', startRow).value;
              newObj.riskAttitude1 = cell('AQ', startRow).value;
              newObj.existingRiskManagement = cell('AR', startRow).value;
              newObj.periodicity = cell('AS', startRow).value;
              newObj.responsiblePerson = cell('AT', startRow).value;
              newObj.completionMark = cell('AU', startRow).value;
              newObj.numWorkers = cell('AV', startRow).value;
              newObj.enterpriseId = req.params.id;

              arr.push(newObj);
            }

            if (!cell('A', startRow).value) {
              const lastObj = arr.at(-1);
              siz.type = cell('G', startRow).value;
              siz.vid = cell('H', startRow).value;
              siz.norm = cell('I', startRow).value;

              lastObj.SIZ.push(siz);
            }
          }

          Value.deleteMany({ enterpriseId: req.params.id })
            .then(() => {
              arr.forEach((item) => {
                Value.create(item)
                  .then(() => res.end())
                  .catch((e) => next(e));
              });
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
      console.log(req.body)
      Value.create(req.body)
        .then((data) => res.send(data))
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
