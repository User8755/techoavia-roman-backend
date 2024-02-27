/* eslint-disable no-underscore-dangle */
/* eslint-disable no-param-reassign */
/* eslint-disable no-return-assign */
const Excel = require('exceljs');
const Value = require('../models/value');
const Enterprise = require('../models/enterprise');
const NotFound = require('../errors/NotFound');

module.exports.createBaseTabel = (req, res, next) => {
  Value.find({ enterpriseId: req.params.id })
    .then((el) => {
      const workbook = new Excel.Workbook();
      const sheet = workbook.addWorksheet('sheet');
      sheet.columns = [
        { header: '№ п/п', key: 'number', width: 9 },
        { header: 'Код профессии (при наличии)', key: 'proffId', width: 20 },
        { header: 'Номер рабочего места', key: 'num', width: 20 },
        { header: 'Профессия', key: 'proff', width: 20 },
        { header: 'Должность', key: 'job', width: 20 },
        { header: 'Подразделение', key: 'subdivision', width: 20 },
        { header: 'Тип средства защиты', key: 'type', width: 20 },
        {
          header:
            'Наименование специальной одежды, специальной обуви и других средств индивидуальной защиты',
          key: 'vid',
          width: 20,
        },
        {
          header: 'Нормы выдачи на год (период) (штуки, пары, комплекты, мл)',
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
        { header: 'Опасное событие 776н', key: 'dangerEvent776', width: 20 },
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
        { header: 'Ответственное лицо', key: 'responsiblePerson', width: 20 },
        { header: 'Отметка о выполнении', key: 'completionMark', width: 20 },
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
      sheet.autoFilter = 'A1:AU29';
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
    .catch((i) => {
      next(i);
    });
};
const workbook = new Excel.Workbook();

module.exports.createNormTabel = (req, res, next) => {
  Value.find({ enterpriseId: req.params.id })
    .then((el) => {
      const fileName = 'normSIZ.xlsx';
      workbook.xlsx
        .readFile(fileName)
        .then((e) => {
          const style = {
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
          };
          const sheet = e.getWorksheet('Лист1');
          const cell = (c, i) => sheet.getCell(c + i);

          let startRow = 11;
          el.forEach((item) => {
            cell('A', startRow).value = item.proffId;
            cell(
              'B',
              startRow,
            ).value = `${item.num}. ${item.job}. ${item.subdivision}.`;
            cell('C', startRow).value = item.typeSIZ === null ? '' : `${item.typeSIZ}`;
            cell('D', startRow).value = item.typeSIZ === null
              ? ''
              : `${item.typeSIZ} \n ${item.speciesSIZ} \n ${item.standart} \n ${item.OperatingLevel}`;
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
            if (item.proffSIZ) {
              item.proffSIZ.forEach((SIZ) => {
                cell('A', startRow).value = item.proffId;
                cell('D', startRow).value = SIZ.vid;
                cell('E', startRow).value = SIZ.norm;
                cell('F', startRow).value = 'Пункт 1 Приложения 1 Приказа 767н';
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
};

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
        Value.find({ enterpriseId: req.params.id }).then((el) => {
          const uniq = el.reduce((accumulator, value) => {
            if (accumulator.every((item) => !(item.num === value.num))) accumulator.push(value);
            return accumulator;
          }, []);

          const fileName = 'mapOPR.xlsx';
          workbook.xlsx
            .readFile(fileName)
            .then((e) => {
              const style = {
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
              };

              const sheet = e.getWorksheet('Лист1');

              uniq.forEach((w) => {
                e.addWorksheet(w.num);
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
                    sheetRow2.eachCell({ includeEmpty: true }, (sourceCell) => {
                      const targetCell = newSheet.getCell(sourceCell.address);

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
                    });
                  }
                }
                const numFilter = el.filter(
                  (filterEl) => filterEl.num === ss.name,
                );
                let i = 30;
                const Ncell = (c) => newSheet.getCell(c);
                numFilter.forEach((elem) => {
                  Ncell('C2').value = ent.enterprise;
                  Ncell('E12').value = elem.num;
                  Ncell('H5').value = elem.num;
                  Ncell(`B${i}`).value = elem.num;
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
    })
    .catch((e) => next(e));
};

module.exports.createListOfMeasuresTabel = (req, res, next) => {
  Value.find({ enterpriseId: req.params.id }).then((el) => {
    const fileName = 'ListOfMeasures.xlsx';
    workbook.xlsx
      .readFile(fileName)
      .then((e) => {
        const style = {
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
        };
        const sheet = e.getWorksheet('TDSheet');

        let line = 21;
        const cell = (c) => sheet.getCell(c + line);

        el.forEach((i) => {
          cell('A').value = line - 20;
          cell('B').value = i.danger776Id || i.dangerGroupId;
          cell('C').value = typeof i.danger776 === 'string'
            ? `${i.danger776} (Приказ776н)`
            : `${i.dangerGroup} (Приказ 767н)`;
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
};

module.exports.createListHazardsTable = (req, res, next) => {
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
          cellB16.value = '№ опасности*';
          cellC16.value = 'Наименование опасности';
          cellD16.value = '№ опасного события*';
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
                  item.dangerEvent776Id === currentValue.dangerEvent776Id
                    && item.dangerEventID === currentValue.dangerEventID
                ),
              )
            ) accumulator.push(currentValue);
            return accumulator;
          }, []);

          const resProff = el.filter(
            ({ num }) => !table2[num] && (table2[num] = 1),
          );

          uniq.forEach((e1) => {
            sheet.getCell(`A${i}`).value = e1.number;
            sheet.getCell(`B${i}`).value = e1.danger776Id || e1.dangerGroupId;
            sheet.getCell(`C${i}`).value = e1.danger776 || e1.dangerGroup;
            sheet.getCell(`D${i}`).value = e1.dangerEvent776Id || e1.dangerEventID;
            sheet.getCell(`E${i}`).value = e1.dangerEvent776 || e1.dangerEvent;

            sheet.getCell(`A${i}`).style = border;
            sheet.getCell(`B${i}`).style = border;
            sheet.getCell(`C${i}`).style = border;
            sheet.getCell(`D${i}`).style = border;
            sheet.getCell(`E${i}`).style = border;
            i += 1;
          });
          const rowAddress = [];

          resProff.forEach((element) => {
            const currentCell = sheet.getColumn(col).letter;
            rowAddress.push(currentCell + 16);

            sheet.getCell(currentCell + 16).value = element.num;
            sheet.getCell(currentCell + 16).style = border;
            sheet.getCell(currentCell + 16).width = 20;
            col += 1;
          });

          rowAddress.forEach((address) => {
            const filterJobValue = el.filter(
              (element) => element.num === sheet.getCell(address).value,
            );

            filterJobValue.forEach((v) => {
              let colStr = i;
              while (colStr >= 17) {
                if (sheet.getCell(`D${colStr}`).value === v.dangerEvent776Id) {
                  sheet.getCell(
                    sheet.getCell(address)._column.letter + colStr,
                  ).value = '+';
                } else if (
                  sheet.getCell(`D${colStr}`).value === v.dangerEventID
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
};

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
              const border = {
                border: {
                  left: { style: 'medium' },
                  top: { style: 'medium' },
                  bottom: { style: 'medium' },
                  right: { style: 'medium' },
                },
                alignment: {
                  horizontal: 'center',
                  vertical: 'middle',
                  wrapText: 'true',
                },
              };
              const sheet = e.getWorksheet(1);
              let start = 16;
              const cell = (c) => sheet.getCell(c);
              cell('B10').value = ent.enterprise;
              el.forEach((value) => {
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

                cell(`A${start}`).style = border;
                cell(`B${start}`).style = border;
                cell(`C${start}`).style = border;
                cell(`D${start}`).style = border;
                cell(`E${start}`).style = border;
                cell(`F${start}`).style = border;
                cell(`G${start}`).style = border;
                cell(`H${start}`).style = border;
                cell(`I${start}`).style = border;
                cell(`J${start}`).style = border;
                cell(`K${start}`).style = border;
                cell(`L${start}`).style = border;
                start += 1;
                sheet.insertRow(start);
              });
              let tableTwoStart = start + 5;
              el.forEach((value, index) => {
                cell(`A${tableTwoStart}`).value = index + 1;
                cell(`B${tableTwoStart}`).value = value.dangerGroupId;
                cell(`C${tableTwoStart}`).value = value.dangerGroup;
                cell(`D${tableTwoStart}`).value = value.dangerEventID;
                cell(`E${tableTwoStart}`).value = value.dangerEvent;
                cell(`F${tableTwoStart}`).value = value.obj;
                cell(`G${tableTwoStart}`).value = value.source;
                cell(`H${tableTwoStart}`).value = value.num;
                cell(`I${tableTwoStart}`).value = value.typeSIZ === null ? '' : `Выдавать: ${value.typeSIZ}`;
                cell(`J${tableTwoStart}`).value = value.issuanceRate;
                cell(`K${tableTwoStart}`).value = value.responsiblePerson;
                cell(`L${tableTwoStart}`).value = value.completionMark;

                cell(`A${tableTwoStart}`).style = border;
                cell(`B${tableTwoStart}`).style = border;
                cell(`C${tableTwoStart}`).style = border;
                cell(`D${tableTwoStart}`).style = border;
                cell(`E${tableTwoStart}`).style = border;
                cell(`F${tableTwoStart}`).style = border;
                cell(`G${tableTwoStart}`).style = border;
                cell(`H${tableTwoStart}`).style = border;
                cell(`I${tableTwoStart}`).style = border;
                cell(`J${tableTwoStart}`).style = border;
                cell(`K${tableTwoStart}`).style = border;
                cell(`L${tableTwoStart}`).style = border;
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
  });
};
