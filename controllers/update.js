const Excel = require('exceljs');
const Proff767 = require('../models/proff767');
const BadRequestError = require('../errors/BadRequestError');

module.exports.UpdatesProff767 = (req, res, next) => {
  const workbook = new Excel.Workbook();
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
          newObj.proff = cell('B', startRow).value;
          newObj.vid = cell('C', startRow).value;
          newObj.type = cell('D', startRow).value;
          newObj.norm = cell('E', startRow).value;

          if (!newObj.proffId) {
            next(new BadRequestError('Не все поля заполнены'));
            break;
          }
          if (!newObj.proff) {
            next(new BadRequestError('Не все поля заполнены'));
            break;
          }
          if (!newObj.vid) {
            next(new BadRequestError('Не все поля заполнены'));
            break;
          }
          if (!newObj.type) {
            next(new BadRequestError('Не все поля заполнены'));
            break;
          }
          if (!newObj.norm) {
            next(new BadRequestError('Не все поля заполнены'));
            break;
          }
          arr.push(newObj);
        }
      } catch (e) {
        throw new BadRequestError('Ошибка заполенения таблици');
      }
      Proff767.deleteMany({}).then(() => {
        arr.forEach((i) => {
          Proff767.create(i).then(() => res.end()).catch((e) => next(e));
        });
      }).catch((e) => next(e));
    }).catch((e) => next(e));
};
