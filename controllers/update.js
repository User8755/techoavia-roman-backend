const Excel = require('exceljs');
const Proff767 = require('../models/proff767');
const TypeSiz = require('../models/typeSiz');
const BadRequestError = require('../errors/BadRequestError');
const workbook = new Excel.Workbook();
// Прикложение 1 приказ 767
module.exports.UpdatesProff767 = (req, res, next) => {
  const handleStyleString = (valueStr) => {
    if (typeof valueStr === 'string') {
      const str = valueStr.charAt(0).toUpperCase() + valueStr.substring(1);
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

      try {
        for (let startRow = 2; startRow <= lastRow.number; startRow += 1) {
          const newObj = {};
          newObj.proffId = Number(cell('A', startRow).value);
          newObj.proff = handleStyleString(cell('B', startRow).value);
          newObj.typeSIZ = handleStyleString(cell('C', startRow).value);
          newObj.speciesSIZ = handleStyleString(cell('D', startRow).value);
          newObj.issuanceRate = handleStyleString(cell('E', startRow).value);

          newObj.markerBase = handleStyleString(cell('F', startRow).value);
          newObj.markerRubber = handleStyleString(cell('G', startRow).value);
          newObj.markerSlip = handleStyleString(cell('H', startRow).value);
          newObj.markerPuncture = handleStyleString(cell('I', startRow).value);
          newObj.markerGlovesAbrasion = handleStyleString(
            cell('J', startRow).value
          );
          newObj.markerGlovesCut = handleStyleString(cell('K', startRow).value);
          newObj.markerGlovesPuncture = handleStyleString(
            cell('L', startRow).value
          );
          newObj.markerWinterShoes = handleStyleString(
            cell('M', startRow).value
          );
          newObj.markerWinterclothes = handleStyleString(
            cell('N', startRow).value
          );
          newObj.markerHierarchyOfClothing = handleStyleString(
            cell('O', startRow).value
          );
          newObj.markerHierarchyOfShoes = handleStyleString(
            cell('P', startRow).value
          );
          newObj.markerHierarchyOfGloves = handleStyleString(
            cell('Q', startRow).value
          );
          newObj.markerTypeSiz = handleStyleString(cell('R', startRow).value);
          newObj.markerMarkerTypeSiz = handleStyleString(
            cell('S', startRow).value
          );

          if (!newObj.proffId) {
            next(new BadRequestError('Не все поля заполнены'));
            break;
          }
          if (!newObj.proff) {
            next(new BadRequestError('Не все поля заполнены'));
            break;
          }
          if (!newObj.speciesSIZ) {
            next(new BadRequestError('Не все поля заполнены'));
            break;
          }
          if (!newObj.typeSIZ) {
            next(new BadRequestError('Не все поля заполнены'));
            break;
          }
          if (!newObj.issuanceRate) {
            next(new BadRequestError('Не все поля заполнены'));
            break;
          }
          arr.push(newObj);
        }
      } catch (e) {
        throw new BadRequestError('Ошибка заполенения таблици');
      }
      Proff767.deleteMany({})
        .then(() => {
          for (i of arr) {
            Proff767.create(i)
              .then(() => res.end())
              .catch((e) => next(e));
          }
        })
        .catch((e) => next(e));
    })
    .catch((e) => next(e));
};

module.exports.UpdatesTypeSiz = (req, res, next) => {
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

      try {
        for (let startRow = 2; startRow <= lastRow.number; startRow += 1) {
          const newObj = {};
          newObj.dependence = handleStyleString(cell('A', startRow).value);
          newObj.label = handleStyleString(cell('B', startRow).value);
          newObj.speciesSIZ = handleStyleString(cell('C', startRow).value);
          newObj.issuanceRate = handleStyleString(cell('D', startRow).value);
          newObj.additionalMeans = handleStyleString(cell('E', startRow).value);
          newObj.AdditionalIssuanceRate = handleStyleString(
            cell('F', startRow).value
          );
          newObj.standart = handleStyleString(cell('G', startRow).value);
          newObj.OperatingLevel = handleStyleString(cell('H', startRow).value);

          newObj.markerBase = handleStyleString(cell('I', startRow).value);
          newObj.markerRubber = handleStyleString(cell('J', startRow).value);
          newObj.markerSlip = handleStyleString(cell('K', startRow).value);
          newObj.markerPuncture = handleStyleString(cell('L', startRow).value);
          newObj.markerGlovesAbrasion = handleStyleString(
            cell('M', startRow).value
          );
          newObj.markerGlovesCut = handleStyleString(cell('N', startRow).value);
          newObj.markerGlovesPuncture = handleStyleString(
            cell('O', startRow).value
          );
          newObj.markerWinterShoes = handleStyleString(
            cell('P', startRow).value
          );
          newObj.markerWinterclothes = handleStyleString(
            cell('Q', startRow).value
          );
          newObj.markerHierarchyOfClothing = handleStyleString(
            cell('R', startRow).value
          );
          newObj.markerHierarchyOfShoes = handleStyleString(
            cell('S', startRow).value
          );
          newObj.markerHierarchyOfGloves = handleStyleString(
            cell('T', startRow).value
          );
          newObj.markerTypeSiz = handleStyleString(cell('U', startRow).value);
          newObj.markerMarkerTypeSiz = handleStyleString(
            cell('V', startRow).value
          );

          if (!newObj.dependence || !newObj.label || !newObj.speciesSIZ) {
            next(new BadRequestError('Не все поля заполнены'));
            break;
          }
          arr.push(newObj);
        }
      } catch (e) {
        throw new BadRequestError('Ошибка заполенения таблици');
      }
      TypeSiz.deleteMany({})
        .then(() => {
          for (i of arr) {
            TypeSiz.create(i)
              .then(() => res.end())
              .catch((e) => next(e));
          }
        })
        .catch((e) => next(e));
    })
    .catch((e) => next(e));
};
