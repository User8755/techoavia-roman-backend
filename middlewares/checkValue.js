const Excel = require('exceljs');
const NotFound = require('../errors/NotFound');
const BadRequestError = require('../errors/BadRequestError');

module.exports = (req, res, next) => {
  const workbook = new Excel.Workbook();
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

  if (!req.files) next(new BadRequestError('Не верный файл'));

  const arr = [];
  workbook.xlsx
    .load(req.files.file.data)
    .then(() => {
      const worksheet = workbook.getWorksheet(1);
      const cell = (lit, num) => worksheet.getCell(lit + num);

      const { lastRow } = worksheet;

      if (cell('AU', 1).value !== 'Отметка о выполнении') {
        next(new NotFound('Не верный файл'));
      }
      let newData;
      try {
        for (let startRow = 2; startRow <= lastRow.number; startRow += 1) {
          const newObj = { proffSIZ: [] };
          const siz = {};
          if (typeof cell('C', startRow).value === 'string' || 'number') {
            newObj.proffId = handleStyleString(cell('B', startRow).value);
            newObj.num = handleStyleString(cell('C', startRow).value);
            newObj.proff = handleStyleString(cell('D', startRow).value);
            newObj.job = handleStyleString(cell('E', startRow).value);
            newObj.subdivision = handleStyleString(cell('F', startRow).value);
            newObj.obj = handleStyleString(cell('J', startRow).value);
            newObj.source = handleStyleString(cell('K', startRow).value);
            newObj.dangerID = handleStyleString(cell('L', startRow).value);
            newObj.danger = handleStyleString(cell('M', startRow).value);
            newObj.dangerGroupId = handleStyleString(cell('N', startRow).value);
            newObj.dangerGroup = handleStyleString(cell('O', startRow).value);
            newObj.dangerEventID = handleStyleString(cell('P', startRow).value);
            newObj.dangerEvent = handleStyleString(cell('Q', startRow).value);
            newObj.heaviness = handleStyleString(cell('R', startRow).value);
            newObj.probability = handleStyleString(cell('S', startRow).value);
            newObj.ipr = handleStyleString(cell('T', startRow).value);
            newObj.risk = handleStyleString(cell('U', startRow).value);
            newObj.acceptability = handleStyleString(cell('V', startRow).value);
            newObj.riskAttitude = handleStyleString(cell('W', startRow).value);
            newObj.typeSIZ = handleStyleString(cell('X', startRow).value);
            newObj.speciesSIZ = handleStyleString(cell('Y', startRow).value);
            newObj.issuanceRate = handleStyleString(cell('Z', startRow).value);
            newObj.additionalMeans = handleStyleString(
              cell('AA', startRow).value,
            );
            newObj.AdditionalIssuanceRate = handleStyleString(
              cell('AB', startRow).value,
            );
            newObj.standart = handleStyleString(cell('AC', startRow).value);
            newObj.OperatingLevel = handleStyleString(
              cell('AD', startRow).value,
            );
            newObj.commit = handleStyleString(cell('AE', startRow).value);
            newObj.danger776Id = handleStyleString(cell('AF', startRow).value);
            newObj.danger776 = handleStyleString(cell('AG', startRow).value);
            newObj.dangerEvent776Id = handleStyleString(
              cell('AH', startRow).value,
            );
            newObj.dangerEvent776 = handleStyleString(
              cell('AI', startRow).value,
            );
            newObj.riskManagementID = handleStyleString(
              cell('AJ', startRow).value,
            );
            newObj.riskManagement = handleStyleString(
              cell('AK', startRow).value,
            );
            newObj.heaviness1 = handleStyleString(cell('AL', startRow).value);
            newObj.probability1 = handleStyleString(cell('AM', startRow).value);
            newObj.ipr1 = handleStyleString(cell('AN', startRow).value);
            newObj.risk1 = handleStyleString(cell('AO', startRow).value);
            newObj.acceptability1 = handleStyleString(
              cell('AP', startRow).value,
            );
            newObj.riskAttitude1 = handleStyleString(
              cell('AQ', startRow).value,
            );
            newObj.existingRiskManagement = handleStyleString(
              cell('AR', startRow).value,
            );
            newObj.periodicity = handleStyleString(cell('AS', startRow).value);
            newObj.responsiblePerson = handleStyleString(
              cell('AT', startRow).value,
            );
            newObj.completionMark = handleStyleString(
              cell('AU', startRow).value,
            );
            newObj.numWorkers = handleStyleString(cell('AV', startRow).value);
            newObj.equipment = handleStyleString(cell('AW', startRow).value);
            newObj.materials = handleStyleString(cell('AX', startRow).value);
            newObj.laborFunction = handleStyleString(
              cell('AY', startRow).value,
            );
            newObj.code = handleStyleString(cell('AZ', startRow).value);
            newObj.enterpriseId = req.params.id;
            if (!newObj.source) {
              next(new BadRequestError('Не все поля "Источник" заполнены'));
              break;
            }
            if (!newObj.num) {
              next(
                new BadRequestError(
                  'Не все поля "Номер рабочего места" заполнены',
                ),
              );
              break;
            }
            if (!newObj.obj) {
              next(new BadRequestError('Не все поля "ОБЪЕКТ" заполнены'));
              break;
            }
            if (!newObj.numWorkers) {
              next(
                new BadRequestError('Не все поля "Кол-во работников" заполнены'),
              );
              break;
            }
            // if (!newObj.risk) {
            //   next(
            //     new BadRequestError('Не все поля "Уровень риска" заполнены'),
            //   );
            //   break;
            // }
            arr.push(newObj);
          }
          if (typeof cell('C', startRow).value !== 'string' || 'number' && cell('G', startRow).value) {
            const lastObj = arr.at(-1);
            siz.type = cell('G', startRow).value;
            siz.vid = cell('H', startRow).value;
            siz.norm = cell('I', startRow).value;
            lastObj.proffSIZ.push(siz);
          }
        }

        newData = arr;
      } catch (e) {
        throw new BadRequestError('Ошибка заполенения таблици');
      }
      req.data = newData;
      next();
    })
    .catch((e) => next(e));
};
