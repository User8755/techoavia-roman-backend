/* eslint-disable no-mixed-operators */
/* eslint-disable no-underscore-dangle */
/* eslint-disable no-param-reassign */
/* eslint-disable no-return-assign */
const Excel = require('exceljs');
const Value = require('../models/value');
const Enterprise = require('../models/enterprise');
const NotFound = require('../errors/NotFound');
const convertValues = require('../forNormTable');
const logs = require('../models/logs');

const darkGeen = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FF00B050' },
};

const green = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FF92D050' },
};

const yellow = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFFFF00' },
};

const orange = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFFC000' },
};

const red = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFF0000' },
};

const style = {
  border: {
    left: { style: 'thin' },
    right: { style: 'thin' },
    bottom: { style: 'thin' },
    top: { style: 'thin' },
  },
  alignment: {
    horizontal: 'left',
    vertical: 'top',
    wrapText: 'true',
  },
};
// Базовая таблица
module.exports.createBaseTabel = (req, res, next) => {
  Enterprise.findById(req.params.id).then((ent) => {
    if (!ent) {
      next(new NotFound('Предприятие не найдено'));
    }
    if (
      ent.owner.toString() === req.user._id
      || ent.access.includes(req.user._id)
    ) {
      Value.find({ enterpriseId: req.params.id })
        .then((el) => {
          const workbook = new Excel.Workbook();
          const sheet = workbook.addWorksheet('sheet');
          sheet.columns = [
            { header: '№ п/п', key: 'number', width: 9 },
            {
              header: 'Код профессии (при наличии)',
              key: 'proffId',
              width: 20,
            },
            { header: 'Номер рабочего места', key: 'num', width: 20 },
            {
              header: 'Профессия (Приказ 767н приложения 1):',
              key: 'proff',
              width: 20,
            },
            { header: 'Профессия', key: 'job', width: 20 },
            { header: 'Подразделение', key: 'subdivision', width: 20 },
            { header: 'Тип средства защиты', key: 'type', width: 20 },
            {
              header:
                'Наименование специальной одежды, специальной обуви и других средств индивидуальной защиты',
              key: 'vid',
              width: 20,
            },
            {
              header:
                'Нормы выдачи на год (период) (штуки, пары, комплекты, мл)',
              key: 'norm',
              width: 20,
            },
            { header: 'ОБЪЕКТ', key: 'obj', width: 20 },
            { header: 'Источник', key: 'source', width: 20 },
            { header: 'ID группы опасностей', key: 'dangerID', width: 20 },
            { header: 'Группа опасности', key: 'danger', width: 25 },
            { header: 'Опасность, ID 767', key: 'dangerGroupId', width: 17 },
            { header: 'Опасности', key: 'dangerGroup', width: 25 },
            {
              header: 'Опасное событие, текст 767',
              key: 'dangerEventID',
              width: 25,
            },
            { header: 'Опасное событие', key: 'dangerEvent', width: 25 },
            { header: 'Тяжесть', key: 'heaviness', width: 8 },
            { header: 'Вероятность', key: 'probability', width: 12 },
            { header: 'ИПР', key: 'ipr', width: 5 },
            { header: 'Уровень риска', key: 'risk', width: 20 },
            { header: 'Приемлемость', key: 'acceptability', width: 20 },
            { header: 'Отношение к риску', key: 'riskAttitude', width: 20 },
            { header: 'Тип СИЗ', key: 'typeSIZ', width: 20 },
            { header: 'Вид СИЗ', key: 'speciesSIZ', width: 40 },
            {
              header:
                'Нормы выдачи средств индивидуальной защиты на год (штуки, пары, комплекты, мл)',
              key: 'issuanceRate',
              width: 20,
            },
            { header: 'ДОП средства', key: 'additionalMeans', width: 20 },
            {
              header:
                'Нормы выдачи средств индивидуальной защиты, выдаваемых дополнительно, на год (штуки, пары, комплекты, мл)',
              key: 'AdditionalIssuanceRate',
              width: 20,
            },
            { header: 'Стандарты (ГОСТ, EN)', key: 'standart', width: 20 },
            { header: 'Экспл.уровень', key: 'OperatingLevel', width: 20 },
            { header: 'Комментарий', key: 'commit', width: 20 },
            { header: 'ID опасности 776н', key: 'danger776Id', width: 20 },
            { header: 'Опасности 776н', key: 'danger776', width: 20 },
            {
              header: 'ID опасного события 776н',
              key: 'dangerEvent776Id',
              width: 20,
            },
            {
              header: 'Опасное событие 776н',
              key: 'dangerEvent776',
              width: 20,
            },
            { header: 'ID мер управления', key: 'riskManagementID', width: 20 },
            {
              header: 'Меры управления/контроля профессиональных рисков',
              key: 'riskManagement',
              width: 20,
            },
            { header: 'Тяжесть', key: 'heaviness1', width: 8 },
            { header: 'Вероятность', key: 'probability1', width: 12 },
            { header: 'ИПР', key: 'ipr1', width: 5 },
            { header: 'Уровень риска1', key: 'risk1', width: 20 },
            { header: 'Приемлемость1', key: 'acceptability1', width: 20 },
            { header: 'Отношение к риску1', key: 'riskAttitude1', width: 20 },
            {
              header: 'Существующие меры упр-я рисками',
              key: 'existingRiskManagement',
              width: 20,
            },
            { header: 'Периодичность', key: 'periodicity', width: 20 },
            {
              header: 'Ответственное лицо',
              key: 'responsiblePerson',
              width: 20,
            },
            {
              header: 'Отметка о выполнении',
              key: 'completionMark',
              width: 20,
            },
            { header: 'Кол-во работников', key: 'numWorkers', width: 20 },
            { header: 'Оборудование', key: 'equipment', width: 20 },
            { header: 'Материалы', key: 'materials', width: 20 },
            { header: 'Трудовая функция', key: 'laborFunction', width: 20 },
            { header: 'Код ОК-016-94:', key: 'code', width: 20 },
          ];
          let i = 1;
          el.forEach((item) => {
            item.number = i;
            sheet.addRow(item);

            if (item.proffSIZ) {
              item.proffSIZ.forEach((SIZ) => sheet.addRow(SIZ));
            }
            i += 1;
          });
          sheet.autoFilter = 'A1:AZ1';

          res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          );
          res.setHeader(
            'Content-Disposition',
            'attachment; filename="Workbook.xlsx"',
          );

          workbook.xlsx
            .write(res)
            .then(() => {
              res.end();
            })
            .catch((err) => {
              res.setHeader('content-type', 'application/json');
              next(err);
            });
        })
        .catch((i) => {
          next(i);
        });
    }
    logs
      .create({
        action: `Пользователь ${req.user.name} выгрузил(а) таблицу Базовая таблица  ${ent.enterprise}`,
        userId: req.user._id,
        enterpriseId: ent._id,
      })
      .catch((e) => next(e));
  });
};
const workbook = new Excel.Workbook();
// Нормы выдачи
module.exports.createNormTabel = (req, res, next) => {
  Enterprise.findById(req.params.id).then((ent) => {
    if (!ent) {
      next(new NotFound('Предприятие не найдено'));
    }
    if (
      ent.owner.toString() === req.user._id
      || ent.access.includes(req.user._id)
    ) {
      Value.find({ enterpriseId: req.params.id })
        .then((el) => {
          const fileName = 'normSIZ.xlsx';
          workbook.xlsx
            .readFile(fileName)
            .then((e) => {
              const entName = `Нормы выдачи средств индивидуальной защиты (далее — СИЗ) в ${ent.enterprise} (наименование подразделения, организации)
                в соответствии с требованиями приказов Минтруда от 29 октября 2021 г.
                №767н «Об утверждении единых типовых норм (далее – ЕТН) выдачи СИЗ и смывающих средств»,
                №766н «Об утверждении правил обеспечения работников средствами индивидуальной защиты и смывающими средствами»
                (далее - приказ №766н)`;

              const sheet = e.getWorksheet('Лист1');

              const cell = (c, i) => sheet.getCell(c + i);
              let startRow = 11;
              sheet.getCell('A5').value = entName;
              el.forEach((item, index) => {
                const handleFilterTypeSIZ = convertValues.find(
                  (i) => i.typeSIZ === item.typeSIZ,
                );

                const stringProff = item.proffId
                  ? `${item.proffId}. ${item.proff}. ${item.subdivision}`
                  : `${item.num}. ${item.job}. ${item.subdivision}.`;
                if (item.typeSIZ) {
                  cell('A', startRow).value = index + 1;
                  cell('B', startRow).value = stringProff;
                  cell('C', startRow).value = `${item.typeSIZ}`;
                  cell('D', startRow).value = !handleFilterTypeSIZ
                    ? `${item.speciesSIZ} ${
                      item.OperatingLevel ? `${item.OperatingLevel}` : ''
                    }  ${item.standart ? `${item.standart}` : ''}`
                    : `${item.speciesSIZ} ${handleFilterTypeSIZ.forTable}  ${
                      item.OperatingLevel ? `${item.OperatingLevel}` : ''
                    }  ${item.standart ? `${item.standart}` : ''}`;
                  cell('E', startRow).value = item.issuanceRate;
                  cell(
                    'F',
                    startRow,
                  ).value = `${item.dangerEventID}, Приложения 2 Приказа 767н`;
                  cell('G', startRow).value = item.dangerGroupId;
                  cell('H', startRow).value = item.dangerGroup;
                  cell('I', startRow).value = item.dangerEventID;
                  cell('J', startRow).value = item.dangerEvent;
                  // стили
                  cell('A', startRow).style = style;
                  cell('B', startRow).style = style;
                  cell('C', startRow).style = style;
                  cell('D', startRow).style = style;
                  cell('E', startRow).style = style;
                  cell('F', startRow).style = style;
                  cell('G', startRow).style = style;
                  cell('H', startRow).style = style;
                  cell('I', startRow).style = style;
                  cell('J', startRow).style = style;
                  startRow += 1;
                  sheet.insertRow(startRow);
                  if (item.additionalMeans) {
                    cell('D', startRow).value = item.additionalMeans;
                    cell('E', startRow).value = item.AdditionalIssuanceRate;
                    // стили
                    cell('A', startRow).style = style;
                    cell('B', startRow).style = style;
                    cell('C', startRow).style = style;
                    cell('D', startRow).style = style;
                    cell('E', startRow).style = style;
                    cell('F', startRow).style = style;
                    cell('G', startRow).style = style;
                    cell('H', startRow).style = style;
                    cell('I', startRow).style = style;
                    cell('J', startRow).style = style;
                    startRow += 1;
                    sheet.insertRow(startRow);
                  }
                  if (item.proffSIZ) {
                    item.proffSIZ.forEach((SIZ) => {
                      cell('D', startRow).value = SIZ.vid;
                      cell('E', startRow).value = SIZ.norm;
                      cell(
                        'F',
                        startRow,
                      ).value = `Пункт ${item.proffId} Приложения 1 Приказа 767н`;
                      cell('G', startRow).value = item.dangerGroupId;
                      cell('H', startRow).value = item.dangerGroup;
                      cell('I', startRow).value = item.dangerEventID;
                      cell('J', startRow).value = item.dangerEvent;
                      // стили
                      cell('A', startRow).style = style;
                      cell('B', startRow).style = style;
                      cell('C', startRow).style = style;
                      cell('D', startRow).style = style;
                      cell('E', startRow).style = style;
                      cell('F', startRow).style = style;
                      cell('G', startRow).style = style;
                      cell('H', startRow).style = style;
                      cell('I', startRow).style = style;
                      cell('J', startRow).style = style;
                      startRow += 1;
                      sheet.insertRow(startRow);
                    });
                  }
                }
              });
              sheet.autoFilter = 'A10:J10';
              res.setHeader(
                'Content-Type',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
              );
              res.setHeader(
                'Content-Disposition',
                `attachment; filename="${Date.now()}_My_Workbook.xlsx"`,
              );

              workbook.xlsx
                .write(res)
                .then(() => {
                  res.end();
                })
                .catch((err) => next(err));
            })
            .catch((e) => next(e));
        })
        .catch((i) => {
          next(i);
        });
    }
    logs
      .create({
        action: `Пользователь ${req.user.name} выгрузил(а) таблицу норма выдачи СИЗ  ${ent.enterprise}`,
        userId: req.user._id,
        enterpriseId: ent._id,
      })
      .catch((e) => next(e));
  });
};
// Карты опаснстей
module.exports.createMapOPRTabel = (req, res, next) => {
  Enterprise.findById(req.params.id)
    .then((ent) => {
      if (!ent) {
        next(new NotFound('Предприятие не найдено'));
      }
      if (
        ent.owner.toString() === req.user._id
        || ent.access.includes(req.user._id)
      ) {
        Value.find({ enterpriseId: req.params.id })
          .sort({ ipr: -1 })
          .then((el) => {
            const uniq = el.reduce((accumulator, value) => {
              if (accumulator.every((item) => !(item.num === value.num))) accumulator.push(value);
              return accumulator;
            }, []);

            const fileName = 'mapOPR.xlsx';
            workbook.xlsx
              .readFile(fileName)
              .then((e) => {
                const sheet = workbook.getWorksheet('Лист1');

                uniq.forEach((w) => {
                  const sheet1 = workbook.addWorksheet(w.num);
                  sheet1.getCell('B30').fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFFF0000' },
                  };
                });

                workbook.worksheets.forEach((ss) => {
                  const newSheet = e.getWorksheet(ss.name);
                  newSheet.getColumn(1).width = 2.33203125;
                  newSheet.getColumn(2).width = 10.66;
                  newSheet.getColumn(3).width = 10.5;
                  newSheet.getColumn(4).width = 20.5;
                  newSheet.getColumn(5).width = 9.6640625;
                  newSheet.getColumn(6).width = 23;
                  newSheet.getColumn(7).width = 23;
                  newSheet.getColumn(8).width = 23;
                  newSheet.getColumn(9).width = 37;
                  newSheet.getColumn(10).width = 13.83203125;
                  newSheet.getColumn(11).width = 14.33203125;
                  newSheet.getColumn(12).width = 18.83203125;
                  newSheet.getColumn(13).width = 21.5;
                  newSheet.getColumn(14).width = 18.140625;
                  for (let a = 1; a <= 30; a += 1) {
                    if (ss.name !== 'Лист1') {
                      const sheetRow2 = sheet.getRow(a);
                      newSheet.getRow(2).height = sheetRow2.height;
                      sheetRow2.eachCell(
                        { includeEmpty: true },
                        (sourceCell) => {
                          const targetCell = newSheet.getCell(
                            sourceCell.address,
                          );

                          // style
                          targetCell.style = sourceCell.style;
                          targetCell.height = sheetRow2.height;
                          // value
                          targetCell.value = sourceCell.value;

                          // merge cell
                          const range = `${
                            sourceCell.model.master || sourceCell.address
                          }:${targetCell.address}`;
                          newSheet.unMergeCells(range);
                          newSheet.mergeCells(range);
                        },
                      );
                    }
                  }
                  const numFilter = el.filter(
                    (filterEl) => filterEl.num === ss.name,
                  );

                  let i = 30;
                  const Ncell = (c) => newSheet.getCell(c);

                  numFilter.forEach((elem, index) => {
                    if (index === 0) {
                      Ncell('F14').value = elem.subdivision;
                      Ncell('G8').value = elem.proff || elem.job;
                      Ncell('H18').value = elem.numWorkers;
                      Ncell('H19').value = elem.equipment;
                      Ncell('H20').value = elem.materials;
                      Ncell('H21').value = elem.laborFunction;
                      Ncell('M8').value = elem.code || elem.proffId;
                    }

                    Ncell('C2').value = ent.enterprise;
                    Ncell('E12').value = elem.num;
                    Ncell('H5').value = elem.num;
                    Ncell(`B${i}`).value = index + 1;
                    Ncell(`C${i}`).value = elem.danger776Id || elem.dangerGroupId;
                    Ncell(`D${i}`).value = elem.danger776 || elem.dangerGroup;
                    Ncell(`E${i}`).value = elem.dangerEvent776Id || elem.dangerEventID;
                    Ncell(`F${i}`).value = elem.dangerEvent776 || elem.dangerEvent;
                    Ncell(`G${i}`).value = elem.obj;
                    Ncell(`H${i}`).value = elem.source;
                    Ncell(`I${i}`).value = elem.existingRiskManagement;
                    Ncell(`J${i}`).value = elem.probability;
                    Ncell(`K${i}`).value = elem.heaviness;
                    Ncell(`L${i}`).value = elem.ipr;
                    Ncell(`M${i}`).value = elem.riskAttitude;
                    Ncell(`N${i}`).value = elem.acceptability;
                    Ncell(`B${i}`).style = style;
                    Ncell(`C${i}`).style = style;
                    Ncell(`D${i}`).style = style;
                    Ncell(`E${i}`).style = style;
                    Ncell(`F${i}`).style = style;
                    Ncell(`G${i}`).style = style;
                    Ncell(`H${i}`).style = style;
                    Ncell(`I${i}`).style = style;
                    Ncell(`J${i}`).style = style;
                    Ncell(`K${i}`).style = style;
                    Ncell(`L${i}`).style = style;
                    Ncell(`M${i}`).style = style;
                    Ncell(`N${i}`).style = style;

                    if (elem.ipr <= 2) {
                      Ncell(`L${i}`).style = {
                        ...(Ncell(`L${i}`).style || {}),
                        fill: darkGeen,
                      };
                      Ncell(`M${i}`).style = {
                        ...(Ncell(`M${i}`).style || {}),
                        fill: darkGeen,
                      };
                    }
                    if (elem.ipr >= 3 && elem.ipr <= 6) {
                      Ncell(`L${i}`).style = {
                        ...(Ncell(`L${i}`).style || {}),
                        fill: green,
                      };
                      Ncell(`M${i}`).style = {
                        ...(Ncell(`M${i}`).style || {}),
                        fill: green,
                      };
                    }
                    if (elem.ipr >= 8 && elem.ipr <= 12) {
                      Ncell(`L${i}`).style = {
                        ...(Ncell(`L${i}`).style || {}),
                        fill: yellow,
                      };
                      Ncell(`M${i}`).style = {
                        ...(Ncell(`M${i}`).style || {}),
                        fill: yellow,
                      };
                    }
                    if (elem.ipr >= 15 && elem.ipr <= 16) {
                      Ncell(`L${i}`).style = {
                        ...(Ncell(`L${i}`).style || {}),
                        fill: orange,
                      };
                      Ncell(`M${i}`).style = {
                        ...(Ncell(`M${i}`).style || {}),
                        fill: orange,
                      };
                    }
                    if (elem.ipr >= 20) {
                      Ncell(`L${i}`).style = {
                        ...(Ncell(`L${i}`).style || {}),
                        fill: red,
                      };
                      Ncell(`M${i}`).style = {
                        ...(Ncell(`M${i}`).style || {}),
                        fill: red,
                      };
                    }

                    i += 1;
                    sheet.insertRow(i);
                  });

                  const styleFooterTitle = {
                    font: {
                      bold: true,
                      size: 12,
                      name: 'Arial',
                      family: 2,
                    },
                    fill: { type: 'pattern', pattern: 'none' },
                    alignment: { horizontal: 'left' },
                  };
                  const styleFooterSubTitle = {
                    font: { size: 12, name: 'Arial', family: 2 },
                    fill: { type: 'pattern', pattern: 'none' },
                    alignment: { horizontal: 'right', vertical: 'top' },
                  };
                  const styleBorder = {
                    border: {
                      bottom: { style: 'thin' },
                    },
                  };
                  const job = '(должность)';
                  const signature = '(подпись)';
                  const date = '(дата)';
                  const FIO = '(Ф.И.О.)';
                  const last = newSheet.lastRow;
                  Ncell(`B${last.number + 3}`).value = '3. Рекомендации работникам:';
                  Ncell(`B${last.number + 7}`).value = '4. Комиссия по оценке профессиональных рисков:';
                  Ncell(`B${last.number + 21}`).value = 'С результатами оценки профессиональных рисков на рабочем месте ознакомлен(ы):';
                  Ncell(`B${last.number + 3}`).style = styleFooterTitle;
                  Ncell(`B${last.number + 7}`).style = styleFooterTitle;
                  Ncell(`B${last.number + 21}`).style = styleFooterTitle;
                  Ncell(`E${last.number + 9}`).value = 'Председатель комиссии:';
                  Ncell(`E${last.number + 12}`).value = 'Члены комиссии:';
                  Ncell(`E${last.number + 9}`).style = styleFooterSubTitle;
                  Ncell(`E${last.number + 12}`).style = styleFooterSubTitle;
                  Ncell(`F${last.number + 10}`).style = styleBorder;
                  Ncell(`G${last.number + 10}`).style = styleBorder;
                  Ncell(`I${last.number + 10}`).style = styleBorder;
                  Ncell(`K${last.number + 10}`).style = styleBorder;
                  Ncell(`N${last.number + 10}`).style = styleBorder;
                  Ncell(`F${last.number + 11}`).value = job;
                  Ncell(`F${last.number + 14}`).value = job;
                  Ncell(`F${last.number + 16}`).value = job;
                  Ncell(`F${last.number + 13}`).style = styleBorder;
                  Ncell(`G${last.number + 13}`).style = styleBorder;
                  Ncell(`I${last.number + 13}`).style = styleBorder;
                  Ncell(`K${last.number + 13}`).style = styleBorder;
                  Ncell(`N${last.number + 13}`).style = styleBorder;
                  Ncell(`F${last.number + 15}`).style = styleBorder;
                  Ncell(`G${last.number + 15}`).style = styleBorder;
                  Ncell(`I${last.number + 15}`).style = styleBorder;
                  Ncell(`K${last.number + 15}`).style = styleBorder;
                  Ncell(`N${last.number + 15}`).style = styleBorder;

                  Ncell(`I${last.number + 11}`).value = FIO;
                  Ncell(`K${last.number + 11}`).value = signature;
                  Ncell(`N${last.number + 11}`).value = date;

                  Ncell(`I${last.number + 14}`).value = FIO;
                  Ncell(`K${last.number + 14}`).value = signature;
                  Ncell(`N${last.number + 14}`).value = date;

                  Ncell(`I${last.number + 16}`).value = FIO;
                  Ncell(`K${last.number + 16}`).value = signature;
                  Ncell(`N${last.number + 16}`).value = date;
                  Ncell(`I${last.number + 15}`).style = styleBorder;
                  Ncell(`K${last.number + 15}`).style = styleBorder;
                  Ncell(`N${last.number + 15}`).style = styleBorder;

                  Ncell(`I${last.number + 29}`).value = FIO;
                  Ncell(`K${last.number + 29}`).value = signature;
                  Ncell(`N${last.number + 29}`).value = date;
                  Ncell(`I${last.number + 28}`).style = styleBorder;
                  Ncell(`K${last.number + 28}`).style = styleBorder;
                  Ncell(`N${last.number + 28}`).style = styleBorder;

                  Ncell(`I${last.number + 26}`).value = FIO;
                  Ncell(`K${last.number + 26}`).value = signature;
                  Ncell(`N${last.number + 26}`).value = date;
                  Ncell(`I${last.number + 25}`).style = styleBorder;
                  Ncell(`K${last.number + 25}`).style = styleBorder;
                  Ncell(`N${last.number + 25}`).style = styleBorder;

                  Ncell(`I${last.number + 32}`).value = FIO;
                  Ncell(`K${last.number + 32}`).value = signature;
                  Ncell(`N${last.number + 32}`).value = date;
                  Ncell(`I${last.number + 31}`).style = styleBorder;
                  Ncell(`K${last.number + 31}`).style = styleBorder;
                  Ncell(`N${last.number + 31}`).style = styleBorder;

                  Ncell(`I${last.number + 35}`).value = FIO;
                  Ncell(`K${last.number + 35}`).value = signature;
                  Ncell(`N${last.number + 35}`).value = date;
                  Ncell(`I${last.number + 34}`).style = styleBorder;
                  Ncell(`K${last.number + 34}`).style = styleBorder;
                  Ncell(`N${last.number + 34}`).style = styleBorder;

                  Ncell(`I${last.number + 38}`).value = FIO;
                  Ncell(`K${last.number + 38}`).value = signature;
                  Ncell(`N${last.number + 38}`).value = date;
                  Ncell(`I${last.number + 37}`).style = styleBorder;
                  Ncell(`K${last.number + 37}`).style = styleBorder;
                  Ncell(`N${last.number + 37}`).style = styleBorder;
                });

                sheet.getColumn(1).width = 2.33203125;
                sheet.getColumn(2).width = 10.66;
                sheet.getColumn(3).width = 10.5;
                sheet.getColumn(4).width = 20.5;
                sheet.getColumn(5).width = 9.6640625;
                sheet.getColumn(6).width = 23;
                sheet.getColumn(7).width = 23;
                sheet.getColumn(8).width = 23;
                sheet.getColumn(9).width = 37;
                sheet.getColumn(10).width = 13.83203125;
                sheet.getColumn(11).width = 14.33203125;
                sheet.getColumn(12).width = 18.83203125;
                sheet.getColumn(13).width = 21.5;

                workbook.removeWorksheet(1);

                res.setHeader(
                  'Content-Type',
                  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                );
                res.setHeader(
                  'Content-Disposition',
                  `attachment; filename="${Date.now()}_My_Workbook.xlsx"`,
                );
                workbook.xlsx
                  .write(res)
                  .then(() => {
                    res.end();
                  })
                  .catch((err) => next(err));
              })
              .catch((e) => next(e));
          });
      }
      logs
        .create({
          action: `Пользователь ${req.user.name} выгрузил(а) таблицу карта опасностей  ${ent.enterprise}`,
          userId: req.user._id,
          enterpriseId: ent._id,
        })
        .catch((e) => next(e));
    })
    .catch((e) => next(e));
};

// Меры управления
module.exports.createListOfMeasuresTabel = (req, res, next) => {
  Enterprise.findById(req.params.id)
    .then((ent) => {
      if (!ent) {
        next(new NotFound('Предприятие не найдено'));
      }
      if (
        ent.owner.toString() === req.user._id
        || ent.access.includes(req.user._id)
      ) {
        Value.find({ enterpriseId: req.params.id })
          .sort({ ipr: -1 })
          .then((el) => {
            const fileName = 'ListOfMeasures.xlsx';
            workbook.xlsx
              .readFile(fileName)
              .then((e) => {
                const sheet = e.getWorksheet('TDSheet');
                const arr = [];
                el.forEach((i) => {
                  const obj = {};
                  if (i.dangerEventID !== '') {
                    if (
                      !arr.find(
                        (n) => n.dangerEventID === i.dangerEventID
                          && n.obj.toLocaleLowerCase().trim()
                            === i.obj.toLocaleLowerCase().trim()
                          && n.source.toLocaleLowerCase().trim()
                            === i.source.toLocaleLowerCase().trim()
                          && n.ipr === i.ipr
                          && n.ipr1 === i.ipr1,
                      )
                    ) {
                      obj.dangerGroupId = i.dangerGroupId;
                      obj.dangerGroup = `${i.dangerGroup} (Приказ 767н)`;
                      obj.dangerEventID = i.dangerEventID;
                      obj.dangerEvent = i.dangerEvent;
                      obj.obj = i.obj;
                      obj.source = i.source;
                      obj.ipr = i.ipr;
                      obj.ipr1 = i.ipr1;
                      obj.riskManagement = i.riskManagement;
                      obj.periodicity = i.periodicity;
                      obj.completionMark = i.completionMark;
                      obj.periodicity = i.periodicity;
                      obj.completionMark = i.completionMark;
                      obj.probability = i.probability;
                      obj.probability1 = i.probability1;
                      obj.heaviness = i.heaviness;
                      obj.heaviness1 = i.heaviness1;
                      arr.push(obj);
                    }
                  }
                  if (i.dangerEventID === '') {
                    if (
                      !arr.find(
                        (n) => n.dangerEvent776Id === i.dangerEvent776Id
                          && n.obj.toLocaleLowerCase().trim()
                            === i.obj.toLocaleLowerCase().trim()
                          && n.source.toLocaleLowerCase().trim()
                            === i.source.toLocaleLowerCase().trim()
                          && n.ipr === i.ipr
                          && n.ipr1 === i.ipr1,
                      )
                    ) {
                      obj.danger776 = `${i.danger776} (Приказ776н)`;
                      obj.danger776Id = i.danger776Id;
                      obj.dangerEvent776Id = i.dangerEvent776Id;
                      obj.dangerEvent776 = i.dangerEvent776;
                      obj.obj = i.obj;
                      obj.source = i.source;
                      obj.ipr = i.ipr;
                      obj.ipr1 = i.ipr1;
                      obj.riskManagement = i.riskManagement;
                      obj.periodicity = i.periodicity;
                      obj.completionMark = i.completionMark;
                      obj.probability = i.probability;
                      obj.probability1 = i.probability1;
                      obj.heaviness = i.heaviness;
                      obj.heaviness1 = i.heaviness1;
                      arr.push(obj);
                    }
                  }
                });
                arr.forEach((i) => {
                  const numArr = [];
                  if (
                    el.filter((n) => {
                      if (
                        n.dangerEventID === i.dangerEventID
                        && n.obj.toLocaleLowerCase().trim()
                          === i.obj.toLocaleLowerCase().trim()
                        && n.source.toLocaleLowerCase().trim()
                          === i.source.toLocaleLowerCase().trim()
                        && n.ipr === i.ipr
                        && n.ipr1 === i.ipr1
                      ) {
                        if (!numArr.includes(n.num)) numArr.push(n.num);
                        let numResult = '';
                        numArr.forEach((nu) => {
                          numResult += `${nu}; `;
                        });
                        i.num = numResult;
                      }
                    })
                  );
                  if (
                    el.filter((n) => {
                      if (
                        n.dangerEvent776Id === i.dangerEvent776Id
                        && n.obj.toLocaleLowerCase().trim()
                          === i.obj.toLocaleLowerCase().trim()
                        && n.source.toLocaleLowerCase().trim()
                          === i.source.toLocaleLowerCase().trim()
                        && n.ipr === i.ipr
                        && n.ipr1 === i.ipr1
                      ) {
                        if (!numArr.includes(n.num)) numArr.push(n.num);
                        let numResult = '';
                        numArr.forEach((nu) => {
                          numResult += `${nu}; `;
                        });
                        i.num = numResult;
                      }
                    })
                  );
                });

                let line = 21;
                const cell = (c) => sheet.getCell(c + line);
                sheet.getCell('C15').value = ent.enterprise;
                arr.forEach((i) => {
                  cell('A').value = line - 20;
                  cell('B').value = i.danger776Id || i.dangerGroupId;
                  cell('C').value = i.danger776 || i.dangerGroup;
                  cell('D').value = i.dangerEvent776Id || i.dangerEventID;
                  cell('E').value = i.dangerEvent776 || i.dangerEvent;
                  cell('F').value = i.obj;
                  cell('G').value = i.source;
                  cell('H').value = i.num;
                  cell('I').value = i.riskManagement;
                  cell('J').value = i.periodicity;
                  cell('L').value = i.completionMark;
                  cell('M').value = i.probability;
                  cell('N').value = i.probability1;
                  cell('O').value = i.heaviness;
                  cell('P').value = i.heaviness1;
                  cell('Q').value = i.ipr;
                  cell('R').value = i.ipr1;

                  cell('A').style = style;
                  cell('B').style = style;
                  cell('C').style = style;
                  cell('D').style = style;
                  cell('E').style = style;
                  cell('F').style = style;
                  cell('G').style = style;
                  cell('H').style = style;
                  cell('I').style = style;
                  cell('J').style = style;
                  cell('L').style = style;
                  cell('M').style = style;
                  cell('N').style = style;
                  cell('O').style = style;
                  cell('P').style = style;
                  cell('Q').style = style;
                  cell('R').style = style;
                  cell('K').style = style;

                  if (i.ipr <= 2) {
                    cell('Q').style = {
                      ...(cell('Q').style || {}),
                      fill: darkGeen,
                    };
                  }
                  if (i.ipr >= 3 && i.ipr <= 6) {
                    cell('Q').style = {
                      ...(cell('Q').style || {}),
                      fill: green,
                    };
                  }
                  if (i.ipr >= 8 && i.ipr <= 12) {
                    cell('Q').style = {
                      ...(cell('Q').style || {}),
                      fill: yellow,
                    };
                  }
                  if (i.ipr >= 15 && i.ipr <= 16) {
                    cell('Q').style = {
                      ...(cell('Q').style || {}),
                      fill: orange,
                    };
                  }
                  if (i.ipr >= 20) {
                    cell('Q').style = {
                      ...(cell('Q').style || {}),
                      fill: red,
                    };
                  }

                  if (i.ipr1 <= 2) {
                    cell('R').style = {
                      ...(cell('R').style || {}),
                      fill: darkGeen,
                    };
                  }
                  if (i.ipr1 >= 3 && i.ipr1 <= 6) {
                    cell('R').style = {
                      ...(cell('R').style || {}),
                      fill: green,
                    };
                  }
                  if (i.ipr1 >= 8 && i.ipr1 <= 12) {
                    cell('R').style = {
                      ...(cell('R').style || {}),
                      fill: yellow,
                    };
                  }
                  if (i.ipr1 >= 15 && i.ipr1 <= 16) {
                    cell('R').style = {
                      ...(cell('R').style || {}),
                      fill: orange,
                    };
                  }
                  if (i.ipr1 >= 20) {
                    cell('R').style = {
                      ...(cell('R').style || {}),
                      fill: red,
                    };
                  }

                  line += 1;
                  sheet.insertRow(line);
                });

                res.setHeader(
                  'Content-Type',
                  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                );
                res.setHeader(
                  'Content-Disposition',
                  `attachment; filename="${Date.now()}_My_Workbook.xlsx"`,
                );
                workbook.xlsx
                  .write(res)
                  .then(() => {
                    res.end();
                  })
                  .catch((err) => next(err));
              })
              .catch((e) => next(e));
          });
      }
      logs
        .create({
          action: `Пользователь ${req.user.name} выгрузил(а) таблицу Меры управления без СИЗ  ${ent.enterprise}`,
          userId: req.user._id,
          enterpriseId: ent._id,
        })
        .catch((e) => next(e));
    })
    .catch((e) => next(e));
};

module.exports.createListHazardsTable = (req, res, next) => {
  Enterprise.findById(req.params.id).then((ent) => {
    if (!ent) {
      next(new NotFound('Предприятие не найдено'));
    }
    if (
      ent.owner.toString() === req.user._id
      || ent.access.includes(req.user._id)
    ) {
      Value.find({ enterpriseId: req.params.id })
        .then((el) => {
          const fileName = 'Реестр опасностей.xlsx';
          workbook.xlsx
            .readFile(fileName)
            .then((e) => {
              const sheet = e.getWorksheet(1);

              const cellA16 = sheet.getCell('A16');
              const cellB16 = sheet.getCell('B16');
              const cellC16 = sheet.getCell('C16');
              const cellD16 = sheet.getCell('D16');
              const cellE16 = sheet.getCell('E16');

              const border = {
                border: {
                  left: { style: 'thin' },
                  right: { style: 'thin' },
                  bottom: { style: 'thin' },
                  top: { style: 'thin' },
                },
                alignment: {
                  horizontal: 'center',
                  vertical: 'middle',
                  wrapText: 'true',
                },
                font: {
                  size: 8,
                  bold: true,
                  name: 'Arial',
                },
              };

              const textRotation = {
                border: {
                  left: { style: 'thin' },
                  right: { style: 'thin' },
                  bottom: { style: 'thin' },
                  top: { style: 'thin' },
                },
                alignment: {
                  horizontal: 'center',
                  vertical: 'middle',
                  wrapText: 'true',
                  textRotation: 'vertical',
                },
                font: {
                  size: 8,
                  bold: true,
                  name: 'Arial',
                },
              };

              cellA16.style = border;
              cellC16.style = border;
              cellE16.style = border;

              cellB16.style = textRotation;
              cellD16.style = textRotation;

              cellA16.value = '№ п/п';
              cellB16.value = '№ опасности';
              cellC16.value = 'Наименование опасности';
              cellD16.value = '№ опасного события';
              cellE16.value = 'Наименование опасного события';

              sheet.getColumn(1).width = 6;
              sheet.getColumn(2).width = 8;
              sheet.getColumn(3).width = 20;
              sheet.getColumn(4).width = 8;
              sheet.getColumn(5).width = 20;
              let i = 17;
              let col = 6;

              const table2 = {};

              const uniq = el.reduce((accumulator, currentValue) => {
                if (
                  accumulator.every(
                    (item) => !(
                      item.dangerEvent776Id
                          === currentValue.dangerEvent776Id
                        && item.dangerEventID === currentValue.dangerEventID
                    ),
                  )
                ) accumulator.push(currentValue);
                return accumulator;
              }, []);

              const resProff = el.filter(
                ({ num }) => !table2[num] && (table2[num] = 1),
              );

              uniq.forEach((e1, index) => {
                sheet.getCell(`A${i}`).value = index + 1;
                sheet.getCell(`B${i}`).value = e1.danger776Id || e1.dangerGroupId;
                sheet.getCell(`C${i}`).value = e1.danger776 || e1.dangerGroup;
                sheet.getCell(`D${i}`).value = e1.dangerEvent776Id || e1.dangerEventID;
                sheet.getCell(`E${i}`).value = e1.dangerEvent776 || e1.dangerEvent;

                sheet.getCell(`A${i}`).style = style;
                sheet.getCell(`B${i}`).style = style;
                sheet.getCell(`C${i}`).style = style;
                sheet.getCell(`D${i}`).style = style;
                sheet.getCell(`E${i}`).style = style;
                i += 1;
              });
              const rowAddress = [];

              resProff
                .sort((a, b) => {
                  const nameA = Number(a.num);
                  const nameB = Number(b.num);
                  if (nameA > nameB) return 1;
                  if (nameA < nameB) return -1;
                  return 0;
                })
                .forEach((element) => {
                  const currentCell = sheet.getColumn(col).letter;
                  rowAddress.push(currentCell + 16);

                  sheet.getCell(currentCell + 16).value = element.num;
                  sheet.getCell(currentCell + 16).style = style;
                  sheet.getCell(currentCell + 16).width = 20;
                  col += 1;
                });

              rowAddress.forEach((address) => {
                const filterJobValue = el.filter(
                  (element) => element.num === sheet.getCell(address).value,
                );

                filterJobValue.forEach((v) => {
                  let colStr = i - 1;
                  while (colStr >= 17) {
                    sheet.getCell(
                      sheet.getCell(address)._column.letter + colStr,
                    ).style = style;
                    if (
                      sheet.getCell(`D${colStr}`).value
                        === v.dangerEvent776Id
                      && sheet.getCell(`D${colStr}`).value !== null
                    ) {
                      sheet.getCell(
                        sheet.getCell(address)._column.letter + colStr,
                      ).value = '+';
                    } else if (
                      sheet.getCell(`D${colStr}`).value === v.dangerEventID
                      && sheet.getCell(`D${colStr}`).value !== null
                    ) {
                      sheet.getCell(
                        sheet.getCell(address)._column.letter + colStr,
                      ).value = '+';
                    }
                    colStr -= 1;
                  }
                });
              });
              res.setHeader(
                'Content-Type',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
              );
              res.setHeader(
                'Content-Disposition',
                `attachment; filename="${Date.now()}_My_Workbook.xlsx"`,
              );
              workbook.xlsx
                .write(res)
                .then(() => {
                  res.end();
                })
                .catch((err) => next(err));
            })
            .catch((e) => next(e));
        })
        .catch((e) => next(e));
    }
    logs
      .create({
        action: `Пользователь ${req.user.name} выгрузил(а) таблицу Перечень идентифицированных опасностей  ${ent.enterprise}`,
        userId: req.user._id,
        enterpriseId: ent._id,
      })
      .catch((e) => next(e));
  });
};
// План-график мер
module.exports.createPlanTimetable = (req, res, next) => {
  Enterprise.findById(req.params.id).then((ent) => {
    if (!ent) {
      next(new NotFound('Предприятие не найдено'));
    }
    if (
      ent.owner.toString() === req.user._id
      || ent.access.includes(req.user._id)
    ) {
      Value.find({ enterpriseId: req.params.id })
        .then((el) => {
          const fileName = 'План-график.xlsx';
          workbook.xlsx
            .readFile(fileName)
            .then((e) => {
              const sheet = e.getWorksheet(1);
              const sheetTwo = e.getWorksheet(2);

              const cell = (c) => sheet.getCell(c);
              const cellSheetTwo = (c) => sheetTwo.getCell(c);
              sheet.autoFilter = 'A15:L15';
              sheetTwo.autoFilter = 'A3:L3';
              cell('B10').value = ent.enterprise;
              let start = 16;
              const arr = [];
              el.forEach((i) => {
                const obj = {};
                if (i.dangerEventID !== '') {
                  if (
                    !arr.find(
                      (n) => n.dangerEventID === i.dangerEventID
                        && n.obj.toLocaleLowerCase().trim()
                          === i.obj.toLocaleLowerCase().trim()
                        && n.source.toLocaleLowerCase().trim()
                          === i.source.toLocaleLowerCase().trim(),
                    )
                  ) {
                    obj.dangerGroupId = i.dangerGroupId;
                    obj.dangerGroup = i.dangerGroup;
                    obj.dangerEventID = i.dangerEventID;
                    obj.dangerEvent = i.dangerEvent;
                    obj.obj = i.obj;
                    obj.source = i.source;
                    obj.riskManagement = i.riskManagement;
                    obj.periodicity = i.periodicity;
                    obj.completionMark = i.completionMark;
                    obj.periodicity = i.periodicity;
                    obj.completionMark = i.completionMark;
                    obj.riskManagement = i.riskManagement;
                    obj.responsiblePerson = i.responsiblePerson;
                    obj.typeSIZ = `Выдавать: ${i.typeSIZ}`;
                    obj.issuanceRate = i.issuanceRate;
                    obj.existingRiskManagement = i.existingRiskManagement;
                    arr.push(obj);
                  }
                }
                if (i.dangerEventID === '') {
                  if (
                    !arr.find(
                      (n) => n.dangerEvent776Id === i.dangerEvent776Id
                        && n.obj.toLocaleLowerCase().trim()
                          === i.obj.toLocaleLowerCase().trim()
                        && n.source.toLocaleLowerCase().trim()
                          === i.source.toLocaleLowerCase().trim(),
                    )
                  ) {
                    obj.danger776 = i.danger776;
                    obj.danger776Id = i.danger776Id;
                    obj.dangerEvent776Id = i.dangerEvent776Id;
                    obj.dangerEvent776 = i.dangerEvent776;
                    obj.obj = i.obj;
                    obj.source = i.source;
                    obj.riskManagement = i.riskManagement;
                    obj.periodicity = i.periodicity;
                    obj.completionMark = i.completionMark;
                    obj.probability = i.probability;
                    obj.probability1 = i.probability1;
                    obj.heaviness = i.heaviness;
                    obj.heaviness1 = i.heaviness1;
                    obj.riskManagement = i.riskManagement;
                    obj.responsiblePerson = i.responsiblePerson;
                    obj.typeSIZ = `Выдавать: ${i.typeSIZ}`;
                    obj.issuanceRate = i.issuanceRate;
                    obj.existingRiskManagement = i.existingRiskManagement;
                    arr.push(obj);
                  }
                }
              });

              arr.forEach((i) => {
                const numArr = [];
                el.filter(
                  (n) => n.dangerEventID === i.dangerEventID
                    && n.obj.toLocaleLowerCase().trim()
                      === i.obj.toLocaleLowerCase().trim()
                    && n.source.toLocaleLowerCase().trim()
                      === i.source.toLocaleLowerCase().trim(),
                ).forEach((nu) => {
                  if (!numArr.includes(nu.num)) numArr.push(nu.num);
                  let numResult = '';
                  numArr.forEach((num) => {
                    numResult += `${num}; `;
                  });
                  i.num = numResult;
                });
                el.filter(
                  (n) => n.dangerEvent776Id === i.dangerEvent776Id
                    && n.obj.toLocaleLowerCase().trim()
                      === i.obj.toLocaleLowerCase().trim()
                    && n.source.toLocaleLowerCase().trim()
                      === i.source.toLocaleLowerCase().trim(),
                ).forEach((nu) => {
                  if (!numArr.includes(nu.num)) numArr.push(nu.num);
                  let numResult = '';
                  numArr.forEach((num) => {
                    numResult += `${num}; `;
                  });
                  i.num = numResult;
                });
              });

              arr.forEach((value) => {
                cell(`A${start}`).value = start - 15;
                cell(`B${start}`).value = value.danger776Id || value.dangerGroupId;
                cell(`C${start}`).value = value.danger776 || value.dangerGroup;
                cell(`D${start}`).value = value.dangerEvent776Id || value.dangerEventID;
                cell(`E${start}`).value = value.dangerEvent776 || value.dangerEvent;
                cell(`F${start}`).value = value.obj;
                cell(`G${start}`).value = value.source;
                cell(`H${start}`).value = value.num;
                cell(
                  `I${start}`,
                ).value = `${value.riskManagement} \n ${value.existingRiskManagement}`;
                cell(`J${start}`).value = value.periodicity;
                cell(`K${start}`).value = value.responsiblePerson;
                cell(`L${start}`).value = value.completionMark;

                cell(`A${start}`).style = style;
                cell(`B${start}`).style = style;
                cell(`C${start}`).style = style;
                cell(`D${start}`).style = style;
                cell(`E${start}`).style = style;
                cell(`F${start}`).style = style;
                cell(`G${start}`).style = style;
                cell(`H${start}`).style = style;
                cell(`I${start}`).style = style;
                cell(`J${start}`).style = style;
                cell(`K${start}`).style = style;
                cell(`L${start}`).style = style;
                start += 1;
              });

              let tableTwoStart = 4;

              arr.forEach((value, index) => {
                cellSheetTwo(`A${tableTwoStart}`).value = index + 1;
                cellSheetTwo(`B${tableTwoStart}`).value = value.dangerGroupId;
                cellSheetTwo(`C${tableTwoStart}`).value = value.dangerGroup;
                cellSheetTwo(`D${tableTwoStart}`).value = value.dangerEventID;
                cellSheetTwo(`E${tableTwoStart}`).value = value.dangerEvent;
                cellSheetTwo(`F${tableTwoStart}`).value = value.obj;
                cellSheetTwo(`G${tableTwoStart}`).value = value.source;
                cellSheetTwo(`H${tableTwoStart}`).value = value.num;
                cellSheetTwo(`I${tableTwoStart}`).value = value.typeSIZ;
                cellSheetTwo(`J${tableTwoStart}`).value = value.issuanceRate;
                cellSheetTwo(`K${tableTwoStart}`).value = value.responsiblePerson;
                cellSheetTwo(`L${tableTwoStart}`).value = value.completionMark;

                cellSheetTwo(`A${tableTwoStart}`).style = style;
                cellSheetTwo(`B${tableTwoStart}`).style = style;
                cellSheetTwo(`C${tableTwoStart}`).style = style;
                cellSheetTwo(`D${tableTwoStart}`).style = style;
                cellSheetTwo(`E${tableTwoStart}`).style = style;
                cellSheetTwo(`F${tableTwoStart}`).style = style;
                cellSheetTwo(`G${tableTwoStart}`).style = style;
                cellSheetTwo(`H${tableTwoStart}`).style = style;
                cellSheetTwo(`I${tableTwoStart}`).style = style;
                cellSheetTwo(`J${tableTwoStart}`).style = style;
                cellSheetTwo(`K${tableTwoStart}`).style = style;
                cellSheetTwo(`L${tableTwoStart}`).style = style;
                tableTwoStart += 1;
              });
              res.setHeader(
                'Content-Type',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
              );
              res.setHeader(
                'Content-Disposition',
                `attachment; filename="${Date.now()}_My_Workbook.xlsx"`,
              );

              workbook.xlsx
                .write(res)
                .then(() => {
                  res.end();
                })
                .catch((err) => next(err));
            })
            .catch((e) => next(e));
        })
        .catch((err) => next(err));
    }
    logs
      .create({
        action: `Пользователь ${req.user.name} выгрузил(а) таблицу План-график мер  ${ent.enterprise}`,
        userId: req.user._id,
        enterpriseId: ent._id,
      })
      .catch((e) => next(e));
  });
};

module.exports.createRegisterHazards = (req, res, next) => {
  Enterprise.findById(req.params.id).then((ent) => {
    if (!ent) {
      next(new NotFound('Предприятие не найдено'));
    }
    if (
      ent.owner.toString() === req.user._id
      || ent.access.includes(req.user._id)
    ) {
      Value.find(
        { enterpriseId: req.params.id },
        {
          source: true,
          dangerEventID: true,
          dangerEvent776Id: true,
          dangerGroup: true,
          danger776: true,
          dangerEvent: true,
          dangerGroupId: true,
          danger776Id: true,
          numWorkers: true,
          ipr: true,
          risk: true,
          dangerEvent776: true,
          num: true,
        },
      )
        .then((el) => {
          const fileName = 'Реестр оцененных опасностей_ИОУПР.xlsx';
          workbook.xlsx
            .readFile(fileName)
            .then((e) => {
              const sheet = e.getWorksheet(1);
              const sheetDiagramma = e.getWorksheet(2);
              const diagrammaValue = {
                veryLow: 0,
                low: 0,
                mid: 0,
                height: 0,
                critical: 0,
              };

              sheet.getCell('G10').value = ent.enterprise;

              // Получаем уникальные источники
              const uniqSource = [];
              el.forEach((i) => {
                if (!uniqSource.includes(i.source)) uniqSource.push(i.source);
              });
              const arr = [];
              uniqSource.forEach((item) => {
                const filter = el.filter((i) => i.source === item);
                filter.forEach((o) => {
                  if (
                    !arr.some(
                      (l) => l.source.toLocaleLowerCase() === o.source.toLocaleLowerCase()
                        && l.dangerEventID === o.dangerEventID,
                    )
                  ) {
                    arr.push({
                      risk: o.risk,
                      source: o.source,
                      dangerEventID: o.dangerEventID,
                      dangerEvent776Id: o.dangerEvent776Id,
                      dangerGroup: o.dangerGroup,
                      danger776: o.danger776,
                      dangerEvent: o.dangerEvent,
                      dangerEvent776: o.dangerEvent776,
                      dangerGroupId: o.dangerGroupId,
                      danger776Id: o.danger776Id,
                      numWorkers: 0, // Кол-во работников
                      countWorkPlaces: 0, // Кол-во рабочих мест
                      // Риски кол-во рабочих мест
                      veryLowPlace: 0, // Незначительный
                      lowPlace: 0, // Низкий
                      midPlace: 0, // Средний
                      highPlace: 0, // Высокий
                      criticalPlace: 0, // Критический
                      // Риски кол-во работников
                      veryLowWorker: 0, // Незначительный
                      lowWorker: 0, // Низкий
                      midWorker: 0, // Средний
                      highWorker: 0, // Высокий
                      criticalWorker: 0, // Критический
                      IPR: 0,
                    });
                  }
                });
              });

              arr.forEach((i) => {
                const countWorker = [];
                el.forEach((c) => {
                  if (c.dangerEventID) {
                    if (c.source.toLocaleLowerCase() === i.source.toLocaleLowerCase()
                      && c.dangerEventID === i.dangerEventID) {
                      if (!countWorker.includes(c.num)) {
                        if (!countWorker.includes(c.num)) {
                          countWorker.push(c.num);
                          i.numWorkers += Number(c.numWorkers);
                          i.IPR += c.ipr;
                        }
                      }
                      i.countWorkPlaces = countWorker.length;
                      switch (c.risk) {
                        case 'Незначительный':
                          i.veryLowWorker += Number(c.numWorkers);
                          i.veryLowPlace += 1;
                          break;
                        case 'Низкий':
                          i.lowWorker += Number(c.numWorkers);
                          i.lowPlace += 1;
                          break;
                        case 'Средний':
                          i.midWorker += Number(c.numWorkers);
                          i.midPlace += 1;
                          break;
                        case 'Высокий':
                          i.highWorker += Number(c.numWorkers);
                          i.highPlace += 1;
                          break;
                        case '':
                          break;
                        default:
                          i.criticalWorker += Number(c.numWorkers);
                          i.criticalPlace += 1;
                      }
                    }
                  }
                  if (!c.dangerEventID) {
                    if (c.source.toLocaleLowerCase() === i.source.toLocaleLowerCase()
                      && c.dangerEvent776Id === i.dangerEvent776Id) {
                      if (!countWorker.includes(c.num)) {
                        countWorker.push(c.num);
                        i.numWorkers += Number(c.numWorkers);
                        i.IPR += c.ipr;
                      }
                      i.countWorkPlaces = countWorker.length;
                      switch (c.risk) {
                        case 'Незначительный':
                          i.veryLowWorker += Number(c.numWorkers);
                          i.veryLowPlace += 1;
                          break;
                        case 'Низкий':
                          i.lowWorker += Number(c.numWorkers);
                          i.lowPlace += 1;
                          break;
                        case 'Средний':
                          i.midWorker += Number(c.numWorkers);
                          i.midPlace += 1;
                          break;
                        case 'Высокий':
                          i.highWorker += Number(c.numWorkers);
                          i.highPlace += 1;
                          break;
                        case '':
                          break;
                        default:
                          i.criticalWorker += Number(c.numWorkers);
                          i.criticalPlace += 1;
                      }
                    }
                  }
                });
              });

              arr
                .sort((a, b) => {
                  const nameA = a.IPR;
                  const nameB = b.IPR;
                  if (nameA > nameB) return -1;
                  if (nameA < nameB) return 1;
                  return 0;
                })
                .forEach((item, index) => {
                  const startRow = index + 15;
                  // Работники
                  const a = item.veryLowWorker
                    + item.lowWorker
                    + item.midWorker
                    + item.highWorker
                    + item.criticalWorker;
                  if (item.veryLowWorker !== 0) {
                    item.vlp = item.veryLowWorker / a;
                  } else {
                    item.vlp = 0;
                  }
                  if (item.lowWorker !== 0) {
                    item.lp = item.lowWorker / a;
                  } else {
                    item.lp = 0;
                  }
                  if (item.midWorker !== 0) {
                    item.mp = item.midWorker / a;
                  } else {
                    item.mp = 0;
                  }
                  if (item.highWorker !== 0) {
                    item.hp = item.highWorker / a;
                  } else {
                    item.hp = 0;
                  }
                  if (item.criticalWorker !== 0) {
                    item.cp = item.criticalWorker / a;
                  } else {
                    item.cp = 0;
                  }
                  item.vl = Math.round(item.vlp * item.numWorkers);
                  item.l = Math.round(item.lp * item.numWorkers);
                  item.m = Math.round(item.mp * item.numWorkers);
                  item.h = Math.round(item.hp * item.numWorkers);
                  item.c = Math.round(item.cp * item.numWorkers);
                  // Рабочие места
                  const b = item.veryLowPlace
                    + item.lowPlace
                    + item.midPlace
                    + item.highPlace
                    + item.criticalPlace;
                  if (item.veryLowPlace !== 0) {
                    item.pvlp = item.veryLowPlace / b;
                  } else {
                    item.pvlp = 0;
                  }
                  if (item.lowPlace !== 0) {
                    item.plp = item.lowPlace / b;
                  } else {
                    item.plp = 0;
                  }
                  if (item.midPlace !== 0) {
                    item.pmp = item.midPlace / b;
                  } else {
                    item.pmp = 0;
                  }
                  if (item.highPlace !== 0) {
                    item.php = item.highPlace / b;
                  } else {
                    item.php = 0;
                  }
                  if (item.criticalPlace !== 0) {
                    item.pcp = item.criticalPlace / b;
                  } else {
                    item.pcp = 0;
                  }

                  item.pvl = Math.round(item.pvlp * item.countWorkPlaces);
                  item.pl = Math.round(item.plp * item.countWorkPlaces);
                  item.pm = item.pmp > 0 && item.pmp < 1
                    ? 1
                    : Math.round(item.pmp * item.countWorkPlaces);
                  item.ph = Math.round(item.php * item.countWorkPlaces);
                  item.pc = Math.round(item.pcp * item.countWorkPlaces);
                  sheet.getCell(`A${startRow}`).value = index + 1;
                  sheet.getCell(`B${startRow}`).value = item.source;
                  sheet.getCell(`C${startRow}`).value = item.dangerGroupId || item.danger776Id;
                  sheet.getCell(`F${startRow}`).value = item.dangerGroup || item.danger776;
                  sheet.getCell(`L${startRow}`).value = item.dangerEventID || item.dangerEvent776Id;
                  sheet.getCell(`O${startRow}`).value = item.dangerEvent || item.dangerEvent776;
                  sheet.getCell(
                    `R${startRow}`,
                  ).value = `${item.numWorkers}/${item.countWorkPlaces}`;
                  sheet.getCell(
                    `T${startRow}`,
                  ).value = `${item.vl}/${item.pvl}`;
                  sheet.getCell(`X${startRow}`).value = `${item.l}/${item.pl}`;
                  sheet.getCell(`AA${startRow}`).value = `${item.m}/${item.pm}`;
                  sheet.getCell(`AD${startRow}`).value = `${item.h}/${item.ph}`;
                  sheet.getCell(`AG${startRow}`).value = `${item.c}/${item.pc}`;
                  sheet.getCell(`AI${startRow}`).value = item.IPR;
                  sheet.getCell(`A${startRow}`).style = style;
                  sheet.getCell(`B${startRow}`).style = style;
                  sheet.getCell(`C${startRow}`).style = style;
                  sheet.getCell(`F${startRow}`).style = style;
                  sheet.getCell(`L${startRow}`).style = style;
                  sheet.getCell(`O${startRow}`).style = style;
                  sheet.getCell(`R${startRow}`).style = style;
                  sheet.getCell(`T${startRow}`).style = style;
                  sheet.getCell(`X${startRow}`).style = style;
                  sheet.getCell(`AA${startRow}`).style = style;
                  sheet.getCell(`AD${startRow}`).style = style;
                  sheet.getCell(`AG${startRow}`).style = style;
                  sheet.getCell(`AI${startRow}`).style = style;
                  sheet.mergeCells(`C${startRow} : E${startRow}`);
                  sheet.mergeCells(`F${startRow} : K${startRow}`);
                  sheet.mergeCells(`L${startRow} : N${startRow}`);
                  sheet.mergeCells(`O${startRow} : Q${startRow}`);
                  sheet.mergeCells(`R${startRow} : S${startRow}`);
                  sheet.mergeCells(`T${startRow} : W${startRow}`);
                  sheet.mergeCells(`X${startRow} : Z${startRow}`);
                  sheet.mergeCells(`AA${startRow} : AC${startRow}`);
                  sheet.mergeCells(`AD${startRow} : AF${startRow}`);
                  sheet.mergeCells(`AG${startRow} : AH${startRow}`);
                  sheet.mergeCells(`AI${startRow} : AJ${startRow}`);
                  if (sheet.getCell(`T${startRow}`).value !== '0/0') {
                    diagrammaValue.veryLow += 1;
                    sheet.getCell(`T${startRow}`).style = {
                      ...(sheet.getCell(`T${startRow}`).style || {}),
                      fill: darkGeen,
                    };
                  }
                  if (sheet.getCell(`X${startRow}`).value !== '0/0') {
                    diagrammaValue.low += 1;
                    sheet.getCell(`X${startRow}`).style = {
                      ...(sheet.getCell(`X${startRow}`).style || {}),
                      fill: green,
                    };
                  }
                  if (sheet.getCell(`AA${startRow}`).value !== '0/0') {
                    diagrammaValue.mid += 1;
                    sheet.getCell(`AA${startRow}`).style = {
                      ...(sheet.getCell(`AA${startRow}`).style || {}),
                      fill: yellow,
                    };
                  }
                  if (sheet.getCell(`AD${startRow}`).value !== '0/0') {
                    diagrammaValue.height += 1;
                    sheet.getCell(`AD${startRow}`).style = {
                      ...(sheet.getCell(`AD${startRow}`).style || {}),
                      fill: orange,
                    };
                  }
                  if (sheet.getCell(`AG${startRow}`).value !== '0/0') {
                    diagrammaValue.critical += 1;
                    sheet.getCell(`AG${startRow}`).style = {
                      ...(sheet.getCell(`AG${startRow}`).style || {}),
                      fill: red,
                    };
                  }
                  sheet.insertRow(index + 16);
                });
              sheetDiagramma.getCell('B3').value = diagrammaValue.veryLow;
              sheetDiagramma.getCell('B4').value = diagrammaValue.low;
              sheetDiagramma.getCell('B5').value = diagrammaValue.mid;
              sheetDiagramma.getCell('B6').value = diagrammaValue.height;
              sheetDiagramma.getCell('B7').value = diagrammaValue.critical;

              res.setHeader(
                'Content-Type',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
              );
              res.setHeader(
                'Content-Disposition',
                `attachment; filename="${Date.now()}_My_Workbook.xlsx"`,
              );

              workbook.xlsx
                .write(res)
                .then(() => {
                  res.end();
                })
                .catch((err) => next(err));
            })
            .catch((e) => next(e));
        })
        .catch((err) => next(err));
    }
    logs
      .create({
        action: `Пользователь ${req.user.name} выгрузил(а) таблицу Реестр оцененных опасностей_ИОУПР  ${ent.enterprise}`,
        userId: req.user._id,
        enterpriseId: ent._id,
      })
      .catch((e) => next(e));
  });
};
