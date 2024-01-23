const Excel = require('exceljs');
const Enterprise = require('../models/enterprise');

module.exports.createBaseTabel = (req, res, next) => {
  Enterprise.findById(req.params.id)
    .then((el) => {
      const workbook = new Excel.Workbook();
      const sheet = workbook.addWorksheet('sheet');
      sheet.columns = [
        { header: '№ п/п', key: 'number', width: 9 },
        { header: 'Код профессии (при наличии)', key: 'proffId', width: 20 },
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
      el.value.forEach((item) => {
        sheet.addRow(item);
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
    .catch((i) => {
      next(i);
    });
};
const workbook = new Excel.Workbook();

module.exports.createNormTabel = (req, res, next) => {
  Enterprise.findById(req.params.id)
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
          el.value.forEach((item) => {
            cell('A', startRow).value = item.proffId;
            cell('B', startRow).value = item.proff || item.job || item.subdivision;
            cell('C', startRow).value = item.typeSIZ;
            cell(
              'D',
              startRow,
            ).value = `${item.typeSIZ} ${item.standart} ${item.OperatingLevel}`;
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
  Enterprise.findById(req.params.id).then((el) => {
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

        sheet.getColumn(1).width = 2.33203125;
        sheet.getColumn(2).width = 5.33203125;
        sheet.getColumn(3).width = 5.33203125;
        sheet.getColumn(4).width = 10.5;
        sheet.getColumn(5).width = 20.5;
        sheet.getColumn(6).width = 9.6640625;
        sheet.getColumn(7).width = 25.33203125;
        sheet.getColumn(8).width = 15;
        sheet.getColumn(9).width = 8;
        sheet.getColumn(10).width = 12.83203125;
        sheet.getColumn(11).width = 10.1640625;
        sheet.getColumn(12).width = 11.1640625;
        sheet.getColumn(13).width = 11.1640625;
        sheet.getColumn(14).width = 14.83203125;
        sheet.getColumn(15).width = 15.1640625;
        sheet.getColumn(16).width = 13.83203125;
        sheet.getColumn(17).width = 14.33203125;
        sheet.getColumn(18).width = 9;
        sheet.getColumn(19).width = 9.83203125;
        sheet.getColumn(20).width = 11.5;
        sheet.getColumn(21).width = 10;

        let i = 30;

        el.value.forEach((elem) => {
          sheet.getCell(`B${i}`).value = i - 29;
          sheet.getCell(`D${i}`).value = elem.danger776Id;
          sheet.getCell(`E${i}`).value = elem.danger776;
          sheet.getCell(`F${i}`).value = elem.dangerEvent776Id;
          sheet.getCell(`G${i}`).value = elem.dangerEvent776;
          sheet.getCell(`H${i}`).value = elem.obj;
          sheet.getCell(`J${i}`).value = elem.source;
          sheet.getCell(`L${i}`).value = elem.existingRiskManagement;
          sheet.getCell(`O${i}`).value = elem.probability;
          sheet.getCell(`P${i}`).value = elem.heaviness;
          sheet.getCell(`Q${i}`).value = elem.ipr;
          sheet.getCell(`R${i}`).value = elem.riskAttitude;
          sheet.getCell(`T${i}`).value = elem.acceptability;
          sheet.getCell(`B${i}`).style = style;
          sheet.getCell(`D${i}`).style = style;
          sheet.getCell(`E${i}`).style = style;
          sheet.getCell(`F${i}`).style = style;
          sheet.getCell(`G${i}`).style = style;
          sheet.getCell(`H${i}`).style = style;
          sheet.getCell(`J${i}`).style = style;
          sheet.getCell(`L${i}`).style = style;
          sheet.getCell(`O${i}`).style = style;
          sheet.getCell(`P${i}`).style = style;
          sheet.getCell(`Q${i}`).style = style;
          sheet.getCell(`R${i}`).style = style;
          sheet.getCell(`T${i}`).style = style;
          sheet.mergeCells(`B${i}`, `C${i}`);
          sheet.mergeCells(`H${i}`, `I${i}`);
          sheet.mergeCells(`J${i}`, `K${i}`);
          sheet.mergeCells(`L${i}`, `N${i}`);
          sheet.mergeCells(`R${i}`, `S${i}`);
          sheet.mergeCells(`T${i}`, `U${i}`);

          i += 1;
          sheet.insertRow(i);
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
