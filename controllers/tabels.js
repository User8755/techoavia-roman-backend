/* eslint-disable no-restricted-syntax */
/* eslint-disable no-mixed-operators */
/* eslint-disable no-underscore-dangle */
/* eslint-disable no-param-reassign */
/* eslint-disable no-return-assign */
const Excel = require('exceljs');
const Value = require('../models/value');
const Enterprise = require('../models/enterprise');
const NotFound = require('../errors/NotFound');
const ConflictError = require('../errors/ConflictError');
const convertValues = require('../forNormTable');
const logs = require('../models/logs');
const Proff767 = require('../models/proff767');
const TypeSiz = require('../models/typeSiz');

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

const borderCell = {
  left: { style: 'thin' },
  right: { style: 'thin' },
  bottom: { style: 'thin' },
  top: { style: 'thin' },
};
// Базовая таблица СИЗ
module.exports.createBaseTabelSIZ = async (req, res, next) => {
  const enterprise = await Enterprise.findById(req.params.id);

  const owner =
    enterprise.owner.toString() != req.user._id ||
    !enterprise.access.includes(req.user._id);

  if (!enterprise) next(new NotFound('Предприятие не найдено'));
  if (!owner) next(new ConflictError('Нет доступа'));

  const value = await Value.find({ enterpriseId: req.params.id });

  const uniqWorkPlace = [...new Set(value.map((i) => i.num))];
  const workbook = new Excel.Workbook();
  const sheet = workbook.addWorksheet('sheet');
  sheet.columns = [
    {
      header: '№ п/п',
      key: 'number',
      width: 9,
      style: {
        fill: {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'E0E0E0' },
        },
      },
      border: {
        left: { style: 'thin' },
        right: { style: 'thin' },
        bottom: { style: 'thin' },
        top: { style: 'thin' },
      },
    },
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
    // { header: 'Тип средства защиты', key: 'type', width: 20 },
    // {
    //   header:
    //     'Наименование специальной одежды, специальной обуви и других средств индивидуальной защиты',
    //   key: 'vid',
    //   width: 20,
    // },
    // {
    //   header: 'Нормы выдачи на год (период) (штуки, пары, комплекты, мл)',
    //   key: 'norm',
    //   width: 20,
    // },
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
    { header: 'Тип синт:', key: 'typeSint', width: 20 },
    { header: 'Вид синт:', key: 'vidSint', width: 20 },
    { header: 'маркер общий:', key: 'marker', width: 20 },
  ];

  for (const obj of uniqWorkPlace) {
    const filtredArr = value.filter((v) => v.num === obj);
    if (filtredArr[0].proffId) {
      const siz = await Proff767.find(
        { proffId: filtredArr[0].proffId },
        { vid: 1, type: 1, norm: 1 }
      );
      const array3 = filtredArr.concat(siz);
      array3.forEach((i) => {
        if (i.type) {
          i.proff = filtredArr[0].proff;
          i.num = filtredArr[0].num;
          i.proffId = filtredArr[0].proffId;
          i.job = filtredArr[0].job;
          i.subdivision = filtredArr[0].subdivision;
          i.typeSIZ = i.type;
          i.speciesSIZ = i.vid;
          i.issuanceRate = i.norm;
        }
      });
      array3
        .sort((a, b) => {
          const nameA = a.typeSIZ;
          const nameB = b.typeSIZ;
          if (nameA > nameB) return 1;
          if (nameA < nameB) return -1;
          return 0;
        })
        .forEach((i) => {
          sheet.addRow(i);
        });
    }
  }
  sheet.autoFilter = 'A1:AZ1';
  res.set(
    'Content-Type',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );
  res.set(
    'Content-Disposition',
    `attachment; filename="${Date.now()}_My_Workbook.xlsx"`
  );
  await workbook.xlsx
    .write(res)
    .then(() => {
      res.end();
    })
    .catch((err) => next(err));
  // Enterprise.findById(req.params.id).then((ent) => {
  //   if (!ent) {
  //     next(new NotFound('Предприятие не найдено'));
  //   }
  //   if (
  //     ent.owner.toString() === req.user._id ||
  //     ent.access.includes(req.user._id)
  //   ) {
  //     Value.find({ enterpriseId: req.params.id })
  //       .then(async (el) => {
  //         const sheet = workbook.addWorksheet('sheet');
  //         sheet.columns = [
  //           {
  //             header: '№ п/п',
  //             key: 'number',
  //             width: 9,
  //             style: {
  //               fill: {
  //                 type: 'pattern',
  //                 pattern: 'solid',
  //                 fgColor: { argb: 'E0E0E0' },
  //               },
  //             },
  //             border: {
  //               left: { style: 'thin' },
  //               right: { style: 'thin' },
  //               bottom: { style: 'thin' },
  //               top: { style: 'thin' },
  //             },
  //           },
  //           {
  //             header: 'Код профессии (при наличии)',
  //             key: 'proffId',
  //             width: 20,
  //           },
  //           { header: 'Номер рабочего места', key: 'num', width: 20 },
  //           {
  //             header: 'Профессия (Приказ 767н приложения 1):',
  //             key: 'proff',
  //             width: 20,
  //           },
  //           { header: 'Профессия', key: 'job', width: 20 },
  //           { header: 'Подразделение', key: 'subdivision', width: 20 },
  //           { header: 'Тип средства защиты', key: 'type', width: 20 },
  //           {
  //             header:
  //               'Наименование специальной одежды, специальной обуви и других средств индивидуальной защиты',
  //             key: 'vid',
  //             width: 20,
  //           },
  //           {
  //             header:
  //               'Нормы выдачи на год (период) (штуки, пары, комплекты, мл)',
  //             key: 'norm',
  //             width: 20,
  //           },
  //           { header: 'ОБЪЕКТ', key: 'obj', width: 20 },
  //           { header: 'Источник', key: 'source', width: 20 },
  //           { header: 'ID группы опасностей', key: 'dangerID', width: 20 },
  //           { header: 'Группа опасности', key: 'danger', width: 25 },
  //           { header: 'Опасность, ID 767', key: 'dangerGroupId', width: 17 },
  //           { header: 'Опасности', key: 'dangerGroup', width: 25 },
  //           {
  //             header: 'Опасное событие, текст 767',
  //             key: 'dangerEventID',
  //             width: 25,
  //           },
  //           { header: 'Опасное событие', key: 'dangerEvent', width: 25 },
  //           { header: 'Тяжесть', key: 'heaviness', width: 8 },
  //           { header: 'Вероятность', key: 'probability', width: 12 },
  //           { header: 'ИПР', key: 'ipr', width: 5 },
  //           { header: 'Уровень риска', key: 'risk', width: 20 },
  //           { header: 'Приемлемость', key: 'acceptability', width: 20 },
  //           { header: 'Отношение к риску', key: 'riskAttitude', width: 20 },
  //           { header: 'Тип СИЗ', key: 'typeSIZ', width: 20 },
  //           { header: 'Вид СИЗ', key: 'speciesSIZ', width: 40 },
  //           {
  //             header:
  //               'Нормы выдачи средств индивидуальной защиты на год (штуки, пары, комплекты, мл)',
  //             key: 'issuanceRate',
  //             width: 20,
  //           },
  //           { header: 'ДОП средства', key: 'additionalMeans', width: 20 },
  //           {
  //             header:
  //               'Нормы выдачи средств индивидуальной защиты, выдаваемых дополнительно, на год (штуки, пары, комплекты, мл)',
  //             key: 'AdditionalIssuanceRate',
  //             width: 20,
  //           },
  //           { header: 'Стандарты (ГОСТ, EN)', key: 'standart', width: 20 },
  //           { header: 'Экспл.уровень', key: 'OperatingLevel', width: 20 },
  //           { header: 'Комментарий', key: 'commit', width: 20 },
  //           { header: 'ID опасности 776н', key: 'danger776Id', width: 20 },
  //           { header: 'Опасности 776н', key: 'danger776', width: 20 },
  //           {
  //             header: 'ID опасного события 776н',
  //             key: 'dangerEvent776Id',
  //             width: 20,
  //           },
  //           {
  //             header: 'Опасное событие 776н',
  //             key: 'dangerEvent776',
  //             width: 20,
  //           },
  //           { header: 'ID мер управления', key: 'riskManagementID', width: 20 },
  //           {
  //             header: 'Меры управления/контроля профессиональных рисков',
  //             key: 'riskManagement',
  //             width: 20,
  //           },
  //           { header: 'Тяжесть', key: 'heaviness1', width: 8 },
  //           { header: 'Вероятность', key: 'probability1', width: 12 },
  //           { header: 'ИПР', key: 'ipr1', width: 5 },
  //           { header: 'Уровень риска1', key: 'risk1', width: 20 },
  //           { header: 'Приемлемость1', key: 'acceptability1', width: 20 },
  //           { header: 'Отношение к риску1', key: 'riskAttitude1', width: 20 },
  //           {
  //             header: 'Существующие меры упр-я рисками',
  //             key: 'existingRiskManagement',
  //             width: 20,
  //           },
  //           { header: 'Периодичность', key: 'periodicity', width: 20 },
  //           {
  //             header: 'Ответственное лицо',
  //             key: 'responsiblePerson',
  //             width: 20,
  //           },
  //           {
  //             header: 'Отметка о выполнении',
  //             key: 'completionMark',
  //             width: 20,
  //           },
  //           { header: 'Кол-во работников', key: 'numWorkers', width: 20 },
  //           { header: 'Оборудование', key: 'equipment', width: 20 },
  //           { header: 'Материалы', key: 'materials', width: 20 },
  //           { header: 'Трудовая функция', key: 'laborFunction', width: 20 },
  //           { header: 'Код ОК-016-94:', key: 'code', width: 20 },
  //           { header: 'Тип синт:', key: 'typeSint', width: 20 },
  //           { header: 'Вид синт:', key: 'vidSint', width: 20 },
  //           { header: 'маркер общий:', key: 'marker', width: 20 },
  //         ];

  //         let strIndex = 1;
  //         const arr = [];
  //         el.forEach((item) => {
  //           item.num = String(item.num);
  //           if (!arr.some((u) => u === item.num)) {
  //             arr.push(item.num);
  //           }
  //         });
  //         arr.sort((a, b) => a.localeCompare(b));
  //         arr.sort((a, b) => a - b);
  //         sheet.autoFilter = 'A1:AZ1';

  //         for (const data of arr) {
  //           const filtredArr = el.filter((f) => f.num === data);
  //           for (let k = 0; k <= filtredArr.length - 1; k += 1) {
  //             // eslint-disable-next-line no-plusplus
  //             filtredArr[k].number = strIndex++;
  //             filtredArr[k].typeSint = filtredArr[k].typeSIZ;
  //             filtredArr[k].vidSint = filtredArr[k].speciesSIZ;
  //             sheet.addRow(filtredArr[k]);
  //             let addSiz = true;
  //             if (filtredArr[k].proffSIZ.length > 0 && addSiz) {
  //               // if (k === 0 && filtredArr[0].proff) {
  //               filtredArr[k].proffSIZ.forEach((siz) => {
  //                 siz.num = filtredArr[k].num;
  //                 siz.proff = filtredArr[k].proff;
  //                 siz.proffId = filtredArr[k].proffId;
  //                 siz.typeSint = siz.type;
  //                 siz.vidSint = siz.vid;
  //                 sheet.addRow(siz);
  //               });
  //               addSiz = false;
  //             }
  //           }
  //         }
  //         res.setHeader(
  //           'Content-Type',
  //           'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  //         );
  //         res.setHeader(
  //           'Content-Disposition',
  //           'attachment; filename="Workbook.xlsx"'
  //         );

  //         workbook.xlsx
  //           .write(res)
  //           .then(() => {
  //             res.end();
  //           })
  //           .catch((err) => {
  //             next(err);
  //           });
  //       })
  //       .catch((i) => {
  //         next(i);
  //       });
  //   }
  //   logs
  //     .create({
  //       action: `Пользователь ${req.user.name} выгрузил(а) таблицу Базовая таблица  ${ent.enterprise}`,
  //       userId: req.user._id,
  //       enterpriseId: ent._id,
  //     })
  //     .catch((e) => next(e));
  // });
};

module.exports.createBaseTabel = (req, res, next) => {
  Enterprise.findById(req.params.id).then((ent) => {
    if (!ent) {
      next(new NotFound('Предприятие не найдено'));
    }
    if (
      ent.owner.toString() === req.user._id ||
      ent.access.includes(req.user._id)
    ) {
      Value.find({ enterpriseId: req.params.id })
        .then(async (el) => {
          const workbook = new Excel.Workbook();
          const sheet = workbook.addWorksheet('sheet');
          sheet.columns = [
            {
              header: '№ п/п',
              key: 'number',
              width: 9,
              style: {
                fill: {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'E0E0E0' },
                },
              },
              border: {
                left: { style: 'thin' },
                right: { style: 'thin' },
                bottom: { style: 'thin' },
                top: { style: 'thin' },
              },
            },
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
            { header: 'Тип синт:', key: 'typeSint', width: 20 },
            { header: 'Вид синт:', key: 'vidSint', width: 20 },
            { header: 'маркер общий:', key: 'marker', width: 20 },
          ];

          let strIndex = 1;
          const arr = [];
          el.forEach((item) => {
            item.num = String(item.num);
            if (!arr.some((u) => u === item.num)) {
              arr.push(item.num);
            }
          });
          arr.sort((a, b) => a.localeCompare(b));
          arr.sort((a, b) => a - b);
          sheet.autoFilter = 'A1:AZ1';

          for (const data of arr) {
            const filtredArr = el.filter((f) => f.num === data);
            for (let k = 0; k <= filtredArr.length - 1; k += 1) {
              // eslint-disable-next-line no-plusplus
              filtredArr[k].number = strIndex++;
              filtredArr[k].typeSint = filtredArr[k].typeSIZ;
              filtredArr[k].vidSint = filtredArr[k].speciesSIZ;
              sheet.addRow(filtredArr[k]);
            }
          }
          res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          );
          res.setHeader(
            'Content-Disposition',
            'attachment; filename="Workbook.xlsx"'
          );

          workbook.xlsx
            .write(res)
            .then(() => {
              res.end();
            })
            .catch((err) => {
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
module.exports.createNormTabel = async (req, res, next) => {
  const enterprise = await Enterprise.findById(req.params.id);

  const entName = `Нормы выдачи средств индивидуальной защиты (далее — СИЗ) в ${enterprise.enterprise} (наименование подразделения, организации)
  в соответствии с требованиями приказов Минтруда от 29 октября 2021 г.
  №767н «Об утверждении единых типовых норм (далее – ЕТН) выдачи СИЗ и смывающих средств»,
  №766н «Об утверждении правил обеспечения работников средствами индивидуальной защиты и смывающими средствами»
  (далее - приказ №766н)`;

  const owner =
    enterprise.owner.toString() != req.user._id ||
    !enterprise.access.includes(req.user._id);

  if (!enterprise) next(new NotFound('Предприятие не найдено'));
  if (!owner) next(new ConflictError('Нет доступа'));

  const value = await Value.find(
    { enterpriseId: req.params.id },
    { proffSIZ: 0 }
  );
  const allTypeSIZ = await TypeSiz.find({});
  const uniqWorkPlace = [...new Set(value.map((i) => i.num))];

  const fileName = 'normSIZ.xlsx';
  await workbook.xlsx.readFile(fileName);
  const sheet = workbook.getWorksheet(1);

  const cell = (c, i) => sheet.getCell(c + i);

  let startRow = 9;
  cell('A', 5).value = entName;

  let array3 = [];
  for (const obj of uniqWorkPlace) {
    const filtredArr = value.filter((v) => v.num === obj);
    if (filtredArr[0].proffId) {
      const siz = await Proff767.find(
        { proffId: filtredArr[0].proffId },
        { proff: 0 }
      );
      array3 = filtredArr.concat(siz);
    }

    for (const i of array3) {
      const strWorlPlace = `${filtredArr[0].num}. ${
        filtredArr[0].proffId ? filtredArr[0].proff : filtredArr[0].job
      }`;
      const typeSIZ = i.typeSIZ || i.type;
      const splitType = i.typeSIZ?.split(' для ')[1] || i.typeSIZ;
      const nameSIZ = `${i.speciesSIZ ? i.speciesSIZ : i.vid} для защиты от ${
        i.typeSIZ ? splitType : i.type
      } ${i.OperatingLevel ? i.OperatingLevel : ''}`;

      const basis = !i.dangerEventID
        ? `Пункт ${i.proffId} Приложения 1 Приказа 767н`
        : `п. опасное событие, текст ${i.dangerEventID} Приложения 2 Приказа 767н`;

      const filtredTypeSIZ = allTypeSIZ.filter(
        (f) => f.dependence === i.dangerEventID && f.speciesSIZ === i.speciesSIZ
      );

      cell('B', startRow).value = filtredArr[0].subdivision;
      cell('C', startRow).value = strWorlPlace;
      cell('D', startRow).value = typeSIZ;
      cell('E', startRow).value = nameSIZ;
      cell('F', startRow).value = i.issuanceRate || i.norm;
      cell('G', startRow).value = basis;

      cell('B', startRow).style = style;
      cell('C', startRow).style = style;
      cell('D', startRow).style = style;
      cell('E', startRow).style = style;
      cell('F', startRow).style = style;
      cell('G', startRow).style = style;

      cell('H', startRow).value = i.markerBase || filtredTypeSIZ[0]?.markerBase;
      cell('I', startRow).value =
        i.markerRubber || filtredTypeSIZ[0]?.markerRubber;
      cell('J', startRow).value = i.markerSlip || filtredTypeSIZ[0]?.markerSlip;
      cell('K', startRow).value =
        i.markerPuncture || filtredTypeSIZ[0]?.markerPuncture;
      cell('L', startRow).value =
        i.markerGlovesAbrasion || filtredTypeSIZ[0]?.markerGlovesAbrasion;
      cell('M', startRow).value =
        i.markerGlovesCut || filtredTypeSIZ[0]?.markerGlovesCut;
      cell('N', startRow).value =
        i.markerGlovesPuncture || filtredTypeSIZ[0]?.markerGlovesPuncture;
      cell('O', startRow).value =
        i.markerWinterShoes || filtredTypeSIZ[0]?.markerWinterShoes;
      cell('P', startRow).value =
        i.markerWinterclothes || filtredTypeSIZ[0]?.markerWinterclothes;
      cell('Q', startRow).value =
        i.markerHierarchyOfClothing ||
        filtredTypeSIZ[0]?.markerHierarchyOfClothing;
      cell('R', startRow).value =
        i.markerHierarchyOfShoes || filtredTypeSIZ[0]?.markerHierarchyOfShoes;
      cell('S', startRow).value =
        i.markerHierarchyOfGloves || filtredTypeSIZ[0]?.markerHierarchyOfGloves;

      startRow++;
      sheet.insertRow(startRow);
    }
  }
  sheet.autoFilter = 'A8:S8';
  res.set(
    'Content-Type',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );
  res.set(
    'Content-Disposition',
    `attachment; filename="${Date.now()}_My_Workbook.xlsx"`
  );
  await workbook.xlsx
    .write(res)
    .then(() => {
      res.end();
    })
    .catch((err) => next(err));
  // const dataSiz = await TypeSiz.find(
  //   {},
  //   {
  //     speciesSIZ: 0,
  //     issuanceRate: 0,
  //     additionalMeans: 0,
  //     AdditionalIssuanceRate: 0,
  //     OperatingLevel: 0,
  //     standart: 0,
  //   }
  // );

  // Enterprise.findById(req.params.id).then((ent) => {
  //   if (!ent) {
  //     next(new NotFound('Предприятие не найдено'));
  //   }
  //   if (
  //     ent.owner.toString() === req.user._id ||
  //     ent.access.includes(req.user._id)
  //   ) {
  //     Value.find({ enterpriseId: req.params.id })
  //       .then((el) => {
  //         for (i of el) {
  //           if (i.proffId) {
  //             Proff767.find(
  //               { proffId: i.proffId },
  //               {
  //                 markerBase: 1,
  //                 markerRubber: 1,
  //                 markerSlip: 1,
  //                 markerPuncture: 1,
  //                 markerGlovesAbrasion: 1,
  //                 markerGlovesCut: 1,
  //                 markerGlovesPuncture: 1,
  //                 markerWinterShoes: 1,
  //                 markerWinterclothes: 1,
  //                 markerHierarchyOfClothing: 1,
  //                 markerHierarchyOfShoes: 1,
  //                 markerHierarchyOfGloves: 1,
  //                 vid: 1,
  //                 norm: 1,
  //                 type: 1,
  //               }
  //             )
  //               .then((item) => {
  //                 i.proffSIZ = item;
  //               })
  //               .catch((e) => next(e));
  //           }
  //         }

  //         const fileName = 'normSIZ.xlsx';
  //         workbook.xlsx
  //           .readFile(fileName)
  //           .then((e) => {
  //             const entName = `Нормы выдачи средств индивидуальной защиты (далее — СИЗ) в ${ent.enterprise} (наименование подразделения, организации)
  //               в соответствии с требованиями приказов Минтруда от 29 октября 2021 г.
  //               №767н «Об утверждении единых типовых норм (далее – ЕТН) выдачи СИЗ и смывающих средств»,
  //               №766н «Об утверждении правил обеспечения работников средствами индивидуальной защиты и смывающими средствами»
  //               (далее - приказ №766н)`;

  //             const sheet = e.getWorksheet('Лист1');
  //             const cell = (c, i) => sheet.getCell(c + i);
  //             let startRow = 9;
  //             sheet.getCell('A5').value = entName;

  //             el.forEach((item) => {
  //               const handletFilterTypeSiz = dataSiz.find(
  //                 (siz) =>
  //                   siz.dependence === item.dangerEventID &&
  //                   siz.label === item.typeSIZ
  //               );

  //               const handleFilterTypeSIZ = convertValues.find(
  //                 (i) => i.typeSIZ === item.typeSIZ
  //               );

  //               const stringProff = item.proffId
  //                 ? `${item.num}. ${item.proff}. ${item.subdivision}`
  //                 : `${item.num}. ${item.job}. ${item.subdivision}.`;
  //               if (item.typeSIZ) {
  //                 //новое
  //                 cell('A', startRow).value = startRow - 8;
  //                 cell('B', startRow).value = item.subdivision;
  //                 cell('C', startRow).value = stringProff;
  //                 cell('D', startRow).value = item.typeSIZ;
  //                 cell('E', startRow).value = !handleFilterTypeSIZ
  //                   ? `${item.speciesSIZ} ${
  //                       item.OperatingLevel ? `${item.OperatingLevel}` : ''
  //                     }  ${item.standart ? `${item.standart}` : ''}`
  //                   : `${item.speciesSIZ} ${handleFilterTypeSIZ.forTable}  ${
  //                       item.OperatingLevel ? `${item.OperatingLevel}` : ''
  //                     }  ${item.standart ? `${item.standart}` : ''}`;
  //                 cell('F', startRow).value = item.issuanceRate;
  //                 cell(
  //                   'G',
  //                   startRow
  //                 ).value = `${item.dangerEventID}, Приложения 2 Приказа 767н`;
  //                 // маркеры
  //                 cell('H', startRow).value = handletFilterTypeSiz.markerBase;
  //                 cell('I', startRow).value = handletFilterTypeSiz.markerRubber;
  //                 cell('J', startRow).value = handletFilterTypeSiz.markerSlip;
  //                 cell('K', startRow).value =
  //                   handletFilterTypeSiz.markerPuncture;
  //                 cell('L', startRow).value =
  //                   handletFilterTypeSiz.markerGlovesAbrasion;
  //                 cell('M', startRow).value =
  //                   handletFilterTypeSiz.markerGlovesCut;
  //                 cell('N', startRow).value =
  //                   handletFilterTypeSiz.markerGlovesPuncture;
  //                 cell('O', startRow).value =
  //                   handletFilterTypeSiz.markerWinterShoes;
  //                 cell('P', startRow).value =
  //                   handletFilterTypeSiz.markerWinterclothes;
  //                 cell('Q', startRow).value =
  //                   handletFilterTypeSiz.markerHierarchyOfClothing;
  //                 cell('R', startRow).value =
  //                   handletFilterTypeSiz.markerHierarchyOfShoes;
  //                 cell('S', startRow).value =
  //                   handletFilterTypeSiz.markerHierarchyOfGloves;
  //                 // стили
  //                 cell('A', startRow).style = style;
  //                 cell('B', startRow).style = style;
  //                 cell('C', startRow).style = style;
  //                 cell('D', startRow).style = style;
  //                 cell('E', startRow).style = style;
  //                 cell('F', startRow).style = style;
  //                 cell('G', startRow).style = style;

  //                 startRow += 1;
  //                 sheet.insertRow(startRow);
  //                 if (item.additionalMeans) {
  //                   cell('B', startRow).value = item.subdivision;
  //                   cell('C', startRow).value = stringProff;
  //                   cell('E', startRow).value = item.additionalMeans;
  //                   cell('F', startRow).value = item.AdditionalIssuanceRate;
  //                   cell(
  //                     'G',
  //                     startRow
  //                   ).value = `${item.dangerEventID}, Приложения 2 Приказа 767н`;

  //                   // стили
  //                   cell('A', startRow).style = style;
  //                   cell('B', startRow).style = style;
  //                   cell('C', startRow).style = style;
  //                   cell('D', startRow).style = style;
  //                   cell('E', startRow).style = style;
  //                   cell('F', startRow).style = style;
  //                   cell('G', startRow).style = style;

  //                   startRow += 1;
  //                   sheet.insertRow(startRow);
  //                 }
  //                 if (item.proffSIZ) {
  //                   item.proffSIZ.forEach((SIZ) => {
  //                     cell('B', startRow).value = item.subdivision;
  //                     cell('C', startRow).value = stringProff;
  //                     cell('D', startRow).value = SIZ.type;
  //                     cell('E', startRow).value = SIZ.vid;
  //                     cell('F', startRow).value = SIZ.norm;
  //                     cell(
  //                       'G',
  //                       startRow
  //                     ).value = `Пункт ${item.proffId} Приложения 1 Приказа 767н`;

  //                     // маркеры
  //                     cell('K', startRow).value = item.markerBase;
  //                     cell('L', startRow).value = item.markerRubber;
  //                     cell('M', startRow).value = item.markerSlip;
  //                     cell('N', startRow).value = item.markerPuncture;
  //                     cell('O', startRow).value = item.markerGlovesAbrasion;
  //                     cell('P', startRow).value = item.markerGlovesCut;
  //                     cell('Q', startRow).value = item.markerGlovesPuncture;
  //                     cell('R', startRow).value = item.markerWinterShoes;
  //                     cell('S', startRow).value = item.markerWinterclothes;
  //                     cell('T', startRow).value =
  //                       item.markerHierarchyOfClothing;
  //                     cell('U', startRow).value = item.markerHierarchyOfShoes;
  //                     cell('V', startRow).value = item.markerHierarchyOfGloves;
  //                     // стили
  //                     cell('A', startRow).style = style;
  //                     cell('B', startRow).style = style;
  //                     cell('C', startRow).style = style;
  //                     cell('D', startRow).style = style;
  //                     cell('E', startRow).style = style;
  //                     cell('F', startRow).style = style;
  //                     cell('G', startRow).style = style;
  //                     startRow += 1;
  //                     sheet.insertRow(startRow);
  //                   });
  //                 }
  //                 //конец
  //                 // cell('A', startRow).value = startRow - 9;
  //                 // cell('B', startRow).value = stringProff;
  //                 // cell('C', startRow).value = `${item.typeSIZ}`;
  //                 // cell('D', startRow).value = !handleFilterTypeSIZ
  //                 //   ? `${item.speciesSIZ} ${
  //                 //       item.OperatingLevel ? `${item.OperatingLevel}` : ''
  //                 //     }  ${item.standart ? `${item.standart}` : ''}`
  //                 //   : `${item.speciesSIZ} ${handleFilterTypeSIZ.forTable}  ${
  //                 //       item.OperatingLevel ? `${item.OperatingLevel}` : ''
  //                 //     }  ${item.standart ? `${item.standart}` : ''}`;
  //                 // cell('E', startRow).value = item.issuanceRate;
  //                 // cell(
  //                 //   'F',
  //                 //   startRow
  //                 // ).value = `${item.dangerEventID}, Приложения 2 Приказа 767н`;
  //                 // cell('G', startRow).value = item.dangerGroupId;
  //                 // cell('H', startRow).value = item.dangerGroup;
  //                 // cell('I', startRow).value = item.dangerEventID;
  //                 // cell('J', startRow).value = item.dangerEvent;
  //                 // // маркеры
  //                 // cell('K', startRow).value = handletFilterTypeSiz.markerBase;
  //                 // cell('L', startRow).value = handletFilterTypeSiz.markerRubber;
  //                 // cell('M', startRow).value = handletFilterTypeSiz.markerSlip;
  //                 // cell('N', startRow).value =
  //                 //   handletFilterTypeSiz.markerPuncture;
  //                 // cell('O', startRow).value =
  //                 //   handletFilterTypeSiz.markerGlovesAbrasion;
  //                 // cell('P', startRow).value =
  //                 //   handletFilterTypeSiz.markerGlovesCut;
  //                 // cell('Q', startRow).value =
  //                 //   handletFilterTypeSiz.markerGlovesPuncture;
  //                 // cell('R', startRow).value =
  //                 //   handletFilterTypeSiz.markerWinterShoes;
  //                 // cell('S', startRow).value =
  //                 //   handletFilterTypeSiz.markerWinterclothes;
  //                 // cell('T', startRow).value =
  //                 //   handletFilterTypeSiz.markerHierarchyOfClothing;
  //                 // cell('U', startRow).value =
  //                 //   handletFilterTypeSiz.markerHierarchyOfShoes;
  //                 // cell('V', startRow).value =
  //                 //   handletFilterTypeSiz.markerHierarchyOfGloves;
  //                 // // стили
  //                 // cell('A', startRow).style = style;
  //                 // cell('B', startRow).style = style;
  //                 // cell('C', startRow).style = style;
  //                 // cell('D', startRow).style = style;
  //                 // cell('E', startRow).style = style;
  //                 // cell('F', startRow).style = style;
  //                 // cell('G', startRow).style = style;
  //                 // cell('H', startRow).style = style;
  //                 // cell('I', startRow).style = style;
  //                 // cell('J', startRow).style = style;
  //                 // startRow += 1;
  //                 // sheet.insertRow(startRow);
  //                 // if (item.additionalMeans) {
  //                 //   cell('B', startRow).value = stringProff;
  //                 //   cell('D', startRow).value = item.additionalMeans;
  //                 //   cell('E', startRow).value = item.AdditionalIssuanceRate;
  //                 //   cell(
  //                 //     'F',
  //                 //     startRow
  //                 //   ).value = `${item.dangerEventID}, Приложения 2 Приказа 767н`;
  //                 //   cell('G', startRow).value = item.dangerGroupId;
  //                 //   cell('H', startRow).value = item.dangerGroup;
  //                 //   cell('I', startRow).value = item.dangerEventID;
  //                 //   cell('J', startRow).value = item.dangerEvent;
  //                 //   // стили
  //                 //   cell('A', startRow).style = style;
  //                 //   cell('B', startRow).style = style;
  //                 //   cell('C', startRow).style = style;
  //                 //   cell('D', startRow).style = style;
  //                 //   cell('E', startRow).style = style;
  //                 //   cell('F', startRow).style = style;
  //                 //   cell('G', startRow).style = style;
  //                 //   cell('H', startRow).style = style;
  //                 //   cell('I', startRow).style = style;
  //                 //   cell('J', startRow).style = style;
  //                 //   startRow += 1;
  //                 //   sheet.insertRow(startRow);
  //                 // }
  //                 // if (item.proffSIZ) {
  //                 //   item.proffSIZ.forEach((SIZ) => {
  //                 //     cell('B', startRow).value = stringProff;
  //                 //     cell('D', startRow).value = SIZ.vid;
  //                 //     cell('E', startRow).value = SIZ.norm;
  //                 //     cell(
  //                 //       'F',
  //                 //       startRow
  //                 //     ).value = `Пункт ${item.proffId} Приложения 1 Приказа 767н`;
  //                 //     cell('G', startRow).value = item.dangerGroupId;
  //                 //     cell('H', startRow).value = item.dangerGroup;
  //                 //     cell('I', startRow).value = item.dangerEventID;
  //                 //     cell('J', startRow).value = item.dangerEvent;

  //                 //     // маркеры
  //                 //     cell('K', startRow).value = item.markerBase;
  //                 //     cell('L', startRow).value = item.markerRubber;
  //                 //     cell('M', startRow).value = item.markerSlip;
  //                 //     cell('N', startRow).value = item.markerPuncture;
  //                 //     cell('O', startRow).value = item.markerGlovesAbrasion;
  //                 //     cell('P', startRow).value = item.markerGlovesCut;
  //                 //     cell('Q', startRow).value = item.markerGlovesPuncture;
  //                 //     cell('R', startRow).value = item.markerWinterShoes;
  //                 //     cell('S', startRow).value = item.markerWinterclothes;
  //                 //     cell('T', startRow).value =
  //                 //       item.markerHierarchyOfClothing;
  //                 //     cell('U', startRow).value = item.markerHierarchyOfShoes;
  //                 //     cell('V', startRow).value = item.markerHierarchyOfGloves;
  //                 //     // стили
  //                 //     cell('A', startRow).style = style;
  //                 //     cell('B', startRow).style = style;
  //                 //     cell('C', startRow).style = style;
  //                 //     cell('D', startRow).style = style;
  //                 //     cell('E', startRow).style = style;
  //                 //     cell('F', startRow).style = style;
  //                 //     cell('G', startRow).style = style;
  //                 //     cell('H', startRow).style = style;
  //                 //     cell('I', startRow).style = style;
  //                 //     cell('J', startRow).style = style;
  //                 //     startRow += 1;
  //                 //     sheet.insertRow(startRow);
  //                 //   });
  //                 // }
  //               }
  //             });
  //             sheet.autoFilter = 'A8:V8';
  //             res.setHeader(
  //               'Content-Type',
  //               'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  //             );
  //             res.setHeader(
  //               'Content-Disposition',
  //               `attachment; filename="${Date.now()}_My_Workbook.xlsx"`
  //             );

  //             workbook.xlsx
  //               .write(res)
  //               .then(() => {
  //                 res.end();
  //               })
  //               .catch((err) => next(err));
  //           })
  //           .catch((e) => next(e));
  //       })
  //       .catch((i) => {
  //         next(i);
  //       });
  //   }
  //   logs
  //     .create({
  //       action: `Пользователь ${req.user.name} выгрузил(а) таблицу норма выдачи СИЗ  ${ent.enterprise}`,
  //       userId: req.user._id,
  //       enterpriseId: ent._id,
  //     })
  //     .catch((e) => next(e));
  // });
};
// Карты опаснстей
module.exports.createMapOPRTabel = (req, res, next) => {
  Enterprise.findById(req.params.id)
    .then((ent) => {
      if (!ent) {
        next(new NotFound('Предприятие не найдено'));
      }
      if (
        ent.owner.toString() === req.user._id ||
        ent.access.includes(req.user._id)
      ) {
        Value.find({ enterpriseId: req.params.id })
          .sort({ ipr: -1 })
          .then((el) => {
            const uniq = el.reduce((accumulator, value) => {
              if (accumulator.every((item) => !(item.num === value.num)))
                accumulator.push(value);
              return accumulator;
            }, []);

            const fileName = 'mapOPR.xlsx';
            workbook.xlsx
              .readFile(fileName)
              .then((e) => {
                const sheet = workbook.getWorksheet('Лист1');

                uniq.forEach((w) => {
                  workbook.addWorksheet(w.num);
                });

                workbook.worksheets.forEach((ss) => {
                  const newSheet = e.getWorksheet(ss.name);
                  for (let a = 1; a <= 43; a += 1) {
                    if (ss.name !== 'Лист1') {
                      const sheetRow2 = sheet.getRow(a);
                      newSheet.getRow(2).height = sheetRow2.height;
                      sheetRow2.eachCell(
                        { includeEmpty: true },
                        (sourceCell) => {
                          const targetCell = newSheet.getCell(
                            sourceCell.address
                          );

                          // style
                          targetCell.style = sourceCell.style;
                          targetCell.height = sheetRow2.height;
                          targetCell._column.width = sourceCell._column.width;
                          // value
                          targetCell.value = sourceCell.value;
                          // merge cell
                          const range = `${
                            sourceCell.model.master || sourceCell.address
                          }:${targetCell.address}`;
                          newSheet.unMergeCells(range);
                          newSheet.mergeCells(range);
                        }
                      );
                    }
                  }
                  const numFilter = el.filter(
                    (filterEl) => filterEl.num === ss.name
                  );

                  let i = 20;
                  const Ncell = (c) => newSheet.getCell(c);

                  Ncell('E25').value = ent.chairmanJob;
                  Ncell('H25').value = ent.chairman;
                  Ncell('E28').value = ent.member1Job;
                  Ncell('H28').value = ent.member1;
                  Ncell('E30').value = ent.member2Job;
                  Ncell('H30').value = ent.member2;

                  numFilter.forEach((elem, index) => {
                    if (index === 0) {
                      Ncell('F8').value = elem.subdivision;
                      Ncell('F6').value = elem.proff || elem.job;
                      Ncell('G11').value = elem.numWorkers;
                      Ncell('G12').value = elem.equipment;
                      Ncell('G13').value = elem.materials;
                      Ncell('G14').value = elem.laborFunction;
                      Ncell('L6').value = elem.code || elem.proffId;
                      Ncell('B1').value = ent.enterprise;
                      Ncell('C7').value = elem.num;
                      Ncell('G4').value = elem.num;
                    }

                    Ncell(`A${i}`).value = index + 1;
                    Ncell(`B${i}`).value =
                      elem.danger776Id || elem.dangerGroupId;
                    Ncell(`C${i}`).value = elem.danger776 || elem.dangerGroup;
                    Ncell(`D${i}`).value =
                      elem.dangerEvent776Id || elem.dangerEventID;
                    Ncell(`E${i}`).value =
                      elem.dangerEvent776 || elem.dangerEvent;
                    Ncell(`F${i}`).value = elem.obj;
                    Ncell(`G${i}`).value = elem.source;
                    Ncell(`H${i}`).value = elem.existingRiskManagement;
                    Ncell(`I${i}`).value = elem.probability;
                    Ncell(`J${i}`).value = elem.heaviness;
                    Ncell(`K${i}`).value = elem.ipr;
                    Ncell(`L${i}`).value = elem.riskAttitude;
                    Ncell(`M${i}`).value = elem.acceptability;
                    Ncell(`A${i}`).style = style;
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

                    if (elem.ipr <= 2) {
                      Ncell(`K${i}`).style = {
                        ...(Ncell(`L${i}`).style || {}),
                        fill: darkGeen,
                      };
                      Ncell(`L${i}`).style = {
                        ...(Ncell(`M${i}`).style || {}),
                        fill: darkGeen,
                      };
                    }
                    if (elem.ipr >= 3 && elem.ipr <= 6) {
                      Ncell(`K${i}`).style = {
                        ...(Ncell(`L${i}`).style || {}),
                        fill: green,
                      };
                      Ncell(`L${i}`).style = {
                        ...(Ncell(`M${i}`).style || {}),
                        fill: green,
                      };
                    }
                    if (elem.ipr >= 8 && elem.ipr <= 12) {
                      Ncell(`K${i}`).style = {
                        ...(Ncell(`L${i}`).style || {}),
                        fill: yellow,
                      };
                      Ncell(`L${i}`).style = {
                        ...(Ncell(`M${i}`).style || {}),
                        fill: yellow,
                      };
                    }
                    if (elem.ipr >= 15 && elem.ipr <= 16) {
                      Ncell(`K${i}`).style = {
                        ...(Ncell(`L${i}`).style || {}),
                        fill: orange,
                      };
                      Ncell(`L${i}`).style = {
                        ...(Ncell(`M${i}`).style || {}),
                        fill: orange,
                      };
                    }
                    if (elem.ipr >= 20) {
                      Ncell(`K${i}`).style = {
                        ...(Ncell(`L${i}`).style || {}),
                        fill: red,
                      };
                      Ncell(`L${i}`).style = {
                        ...(Ncell(`M${i}`).style || {}),
                        fill: red,
                      };
                    }

                    i += 1;
                    newSheet.insertRow(i);
                  });
                });

                workbook.removeWorksheet(1);

                res.setHeader(
                  'Content-Type',
                  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                );
                res.setHeader(
                  'Content-Disposition',
                  `attachment; filename="${Date.now()}_My_Workbook.xlsx"`
                );
                workbook.xlsx
                  .write(res)
                  .then(() => {
                    res.end();
                  })
                  .catch((err) => {
                    next(err);
                  });
              })
              .catch((err) => {
                next(err);
              });
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
        ent.owner.toString() === req.user._id ||
        ent.access.includes(req.user._id)
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
                        (n) =>
                          n.dangerEventID === i.dangerEventID &&
                          n.obj.toLocaleLowerCase().trim() ===
                            i.obj.toLocaleLowerCase().trim() &&
                          n.source.toLocaleLowerCase().trim() ===
                            i.source.toLocaleLowerCase().trim() &&
                          n.ipr === i.ipr &&
                          n.ipr1 === i.ipr1
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
                      obj.existingRiskManagement = i.existingRiskManagement;
                      arr.push(obj);
                    }
                  }
                  if (i.dangerEventID === '') {
                    if (
                      !arr.find(
                        (n) =>
                          n.dangerEvent776Id === i.dangerEvent776Id &&
                          n.obj.toLocaleLowerCase().trim() ===
                            i.obj.toLocaleLowerCase().trim() &&
                          n.source.toLocaleLowerCase().trim() ===
                            i.source.toLocaleLowerCase().trim() &&
                          n.ipr === i.ipr &&
                          n.ipr1 === i.ipr1
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
                      obj.existingRiskManagement = i.existingRiskManagement;
                      arr.push(obj);
                    }
                  }
                });
                arr.forEach((i) => {
                  const numArr = [];
                  if (
                    el.filter((n) => {
                      if (
                        n.dangerEventID === i.dangerEventID &&
                        n.obj.toLocaleLowerCase().trim() ===
                          i.obj.toLocaleLowerCase().trim() &&
                        n.source.toLocaleLowerCase().trim() ===
                          i.source.toLocaleLowerCase().trim() &&
                        n.ipr === i.ipr &&
                        n.ipr1 === i.ipr1
                      ) {
                        if (!numArr.includes(n.num)) numArr.push(n.num);
                        let numResult = '';
                        numArr.forEach((nu) => {
                          numResult += `${nu}; `;
                        });
                        i.num = numResult;
                      }
                      return n;
                    })
                  );
                  if (
                    el.filter((n) => {
                      if (
                        n.dangerEvent776Id === i.dangerEvent776Id &&
                        n.obj.toLocaleLowerCase().trim() ===
                          i.obj.toLocaleLowerCase().trim() &&
                        n.source.toLocaleLowerCase().trim() ===
                          i.source.toLocaleLowerCase().trim() &&
                        n.ipr === i.ipr &&
                        n.ipr1 === i.ipr1
                      ) {
                        if (!numArr.includes(n.num)) numArr.push(n.num);
                        let numResult = '';
                        numArr.forEach((nu) => {
                          numResult += `${nu}; `;
                        });
                        i.num = numResult;
                      }
                      return n;
                    })
                  );
                });

                let line = 13;
                const cell = (c) => sheet.getCell(c + line);
                sheet.getCell('C8').value = ent.enterprise;
                sheet.getCell('P3').value = ent.chairman;
                sheet.getCell('B14').value = ent.member1Job;
                sheet.getCell('G14').value = ent.member1;
                sheet.getCell('B16').value = ent.member2Job;
                sheet.getCell('G16').value = ent.member2;
                sheet.getCell('B18').value = ent.member3Job;
                sheet.getCell('G18').value = ent.member3;

                arr.forEach((i) => {
                  cell('A').value = line - 12;
                  cell('B').value = i.danger776Id || i.dangerGroupId;
                  cell('C').value = i.danger776 || i.dangerGroup;
                  cell('D').value = i.dangerEvent776Id || i.dangerEventID;
                  cell('E').value = i.dangerEvent776 || i.dangerEvent;
                  cell('F').value = i.obj;
                  cell('G').value = i.source;
                  cell('H').value = i.num;
                  cell('I').value = i.existingRiskManagement;
                  cell('K').value = i.periodicity;
                  cell('J').value = i.riskManagement;
                  cell('L').value = i.responsiblePerson;
                  cell('M').value = i.completionMark;
                  cell('N').value = i.probability;
                  cell('O').value = i.probability1;
                  cell('P').value = i.heaviness;
                  cell('Q').value = i.heaviness1;
                  cell('R').value = i.ipr;
                  cell('S').value = i.ipr1;

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
                  cell('S').style = style;

                  if (i.ipr <= 2) {
                    cell('R').style = {
                      ...(cell('R').style || {}),
                      fill: darkGeen,
                    };
                  }
                  if (i.ipr >= 3 && i.ipr <= 6) {
                    cell('R').style = {
                      ...(cell('R').style || {}),
                      fill: green,
                    };
                  }
                  if (i.ipr >= 8 && i.ipr <= 12) {
                    cell('R').style = {
                      ...(cell('R').style || {}),
                      fill: yellow,
                    };
                  }
                  if (i.ipr >= 15 && i.ipr <= 16) {
                    cell('R').style = {
                      ...(cell('R').style || {}),
                      fill: orange,
                    };
                  }
                  if (i.ipr >= 20) {
                    cell('R').style = {
                      ...(cell('R').style || {}),
                      fill: red,
                    };
                  }

                  if (i.ipr1 <= 2) {
                    cell('S').style = {
                      ...(cell('S').style || {}),
                      fill: darkGeen,
                    };
                  }
                  if (i.ipr1 >= 3 && i.ipr1 <= 6) {
                    cell('S').style = {
                      ...(cell('S').style || {}),
                      fill: green,
                    };
                  }
                  if (i.ipr1 >= 8 && i.ipr1 <= 12) {
                    cell('S').style = {
                      ...(cell('S').style || {}),
                      fill: yellow,
                    };
                  }
                  if (i.ipr1 >= 15 && i.ipr1 <= 16) {
                    cell('S').style = {
                      ...(cell('S').style || {}),
                      fill: orange,
                    };
                  }
                  if (i.ipr1 >= 20) {
                    cell('S').style = {
                      ...(cell('S').style || {}),
                      fill: red,
                    };
                  }

                  line += 1;
                  sheet.insertRow(line);
                });

                res.setHeader(
                  'Content-Type',
                  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                );
                res.setHeader(
                  'Content-Disposition',
                  `attachment; filename="${Date.now()}_My_Workbook.xlsx"`
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

// Перечень идентифицированных опасностей
module.exports.createListHazardsTable = (req, res, next) => {
  Enterprise.findById(req.params.id).then((ent) => {
    if (!ent) {
      next(new NotFound('Предприятие не найдено'));
    }
    if (
      ent.owner.toString() === req.user._id ||
      ent.access.includes(req.user._id)
    ) {
      Value.find(
        { enterpriseId: req.params.id },
        {
          dangerEvent776Id: 1,
          dangerEventID: 1,
          num: 1,
          danger776Id: 1,
          dangerGroupId: 1,
          danger776: 1,
          dangerGroup: 1,
          dangerEvent776: 1,
          dangerEvent: 1,
        }
      )
        .then((el) => {
          const fileName = 'Реестр опасностей.xlsx';
          workbook.xlsx
            .readFile(fileName)
            .then((e) => {
              const sheet = e.getWorksheet(1);

              let i = 14;
              let col = 6;

              const table2 = {};

              const uniq = el.reduce((accumulator, currentValue) => {
                if (
                  accumulator.every(
                    (item) =>
                      !(
                        item.dangerEvent776Id ===
                          currentValue.dangerEvent776Id &&
                        item.dangerEventID === currentValue.dangerEventID
                      )
                  )
                )
                  accumulator.push(currentValue);
                return accumulator;
              }, []);

              const resProff = el.filter(
                ({ num }) => !table2[num] && (table2[num] = 1)
              );

              uniq.forEach((e1) => {
                if (e1.dangerEvent776Id || e1.dangerEventID) {
                  sheet.getCell(`A${i}`).value = i - 13;
                  sheet.getCell(`B${i}`).value =
                    e1.danger776Id || e1.dangerGroupId;
                  sheet.getCell(`C${i}`).value = e1.danger776 || e1.dangerGroup;
                  sheet.getCell(`D${i}`).value =
                    e1.dangerEvent776Id || e1.dangerEventID;
                  sheet.getCell(`E${i}`).value =
                    e1.dangerEvent776 || e1.dangerEvent;

                  sheet.getCell(`A${i}`).style = style;
                  sheet.getCell(`B${i}`).style = style;
                  sheet.getCell(`C${i}`).style = style;
                  sheet.getCell(`D${i}`).style = style;
                  sheet.getCell(`E${i}`).style = style;
                  i += 1;
                }
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
                  rowAddress.push(currentCell + 13);

                  sheet.getCell(currentCell + 13).value = element.num;
                  sheet.getCell(currentCell + 13).style = style;

                  col += 1;
                });

              rowAddress.forEach((address) => {
                const filterJobValue = el.filter(
                  (element) => element.num === sheet.getCell(address).value
                );

                filterJobValue.forEach((v) => {
                  let colStr = i - 1;
                  while (colStr >= 13) {
                    sheet.getCell(
                      sheet.getCell(address)._column.letter + colStr
                    ).style = style;
                    if (
                      sheet.getCell(`D${colStr}`).value ===
                        v.dangerEvent776Id &&
                      sheet.getCell(`D${colStr}`).value !== null
                    ) {
                      sheet.getCell(
                        sheet.getCell(address)._column.letter + colStr
                      ).value = '+';
                    } else if (
                      sheet.getCell(`D${colStr}`).value === v.dangerEventID &&
                      sheet.getCell(`D${colStr}`).value !== null
                    ) {
                      sheet.getCell(
                        sheet.getCell(address)._column.letter + colStr
                      ).value = '+';
                    }
                    colStr -= 1;
                  }
                });
              });
              res.setHeader(
                'Content-Type',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
              );
              res.setHeader(
                'Content-Disposition',
                `attachment; filename="${Date.now()}_My_Workbook.xlsx"`
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
      ent.owner.toString() === req.user._id ||
      ent.access.includes(req.user._id)
    ) {
      Value.find({ enterpriseId: req.params.id })
        .then((el) => {
          const fileName = 'План-график.xlsx';
          workbook.xlsx
            .readFile(fileName)
            .then((e) => {
              const sheet = e.getWorksheet(1);
              const sheetTwo = e.getWorksheet(2);
              const sheetThree = e.getWorksheet(3);

              const cell = (c) => sheet.getCell(c);
              const cellSheetTwo = (c) => sheetTwo.getCell(c);
              const cellSheetThree = (c) => sheetThree.getCell(c);
              sheet.autoFilter = 'A15:L15';
              sheetTwo.autoFilter = 'A3:L3';
              sheetThree.autoFilter = 'A14:L14';
              cell('B10').value = ent.enterprise;
              let start = 15;
              const arr = [];
              el.forEach((i) => {
                const obj = {};
                if (i.dangerEventID !== '') {
                  if (
                    !arr.find(
                      (n) =>
                        n.dangerEventID === i.dangerEventID &&
                        n.obj.toLocaleLowerCase().trim() ===
                          i.obj.toLocaleLowerCase().trim() &&
                        n.source.toLocaleLowerCase().trim() ===
                          i.source.toLocaleLowerCase().trim()
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
                      (n) =>
                        n.dangerEvent776Id === i.dangerEvent776Id &&
                        n.obj.toLocaleLowerCase().trim() ===
                          i.obj.toLocaleLowerCase().trim() &&
                        n.source.toLocaleLowerCase().trim() ===
                          i.source.toLocaleLowerCase().trim()
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
                  (n) =>
                    n.dangerEventID === i.dangerEventID &&
                    n.obj.toLocaleLowerCase().trim() ===
                      i.obj.toLocaleLowerCase().trim() &&
                    n.source.toLocaleLowerCase().trim() ===
                      i.source.toLocaleLowerCase().trim()
                ).forEach((nu) => {
                  if (!numArr.includes(nu.num)) numArr.push(nu.num);
                  let numResult = '';
                  numArr.forEach((num) => {
                    numResult += `${num}; `;
                  });
                  i.num = numResult;
                });
                el.filter(
                  (n) =>
                    n.dangerEvent776Id === i.dangerEvent776Id &&
                    n.obj.toLocaleLowerCase().trim() ===
                      i.obj.toLocaleLowerCase().trim() &&
                    n.source.toLocaleLowerCase().trim() ===
                      i.source.toLocaleLowerCase().trim()
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
                cell(`B${start}`).value =
                  value.danger776Id || value.dangerGroupId;
                cell(`C${start}`).value = value.danger776 || value.dangerGroup;
                cell(`D${start}`).value =
                  value.dangerEvent776Id || value.dangerEventID;
                cell(`E${start}`).value =
                  value.dangerEvent776 || value.dangerEvent;
                cell(`F${start}`).value = value.obj;
                cell(`G${start}`).value = value.source;
                cell(`H${start}`).value = value.num;
                cell(`I${start}`).value = value.riskManagement;
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

              let tableTwoStart = 3;

              arr.forEach((value, index) => {
                if (value.typeSIZ.length > 10) {
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
                  cellSheetTwo(`K${tableTwoStart}`).value =
                    value.responsiblePerson;
                  cellSheetTwo(`L${tableTwoStart}`).value =
                    value.completionMark;

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
                }
              });
              let tableThreeStart = 15;
              cellSheetThree(`J4`).value = ent.chairman;
              cellSheetThree(`I2`).value = ent.chairmanJob;
              cellSheetThree(`B10`).value = ent.enterprise;
              el.forEach((i) => {
                if (i.typeSIZ) {
                  const number = tableThreeStart - 14;
                  cellSheetThree(`A${tableThreeStart}`).value = number;
                  cellSheetThree(`B${tableThreeStart}`).value = i.num;
                  cellSheetThree(`C${tableThreeStart}`).value = i.prof || i.job;
                  cellSheetThree(`D${tableThreeStart}`).value = i.dangerGroupId;
                  cellSheetThree(`E${tableThreeStart}`).value = i.dangerGroup;
                  cellSheetThree(`F${tableThreeStart}`).value = i.dangerEventID;
                  cellSheetThree(`G${tableThreeStart}`).value = i.dangerEvent;
                  cellSheetThree(`H${tableThreeStart}`).value = i.obj;
                  cellSheetThree(`I${tableThreeStart}`).value = i.source;
                  cellSheetThree(`J${tableThreeStart}`).value = i.typeSIZ;
                  cellSheetThree(`K${tableThreeStart}`).value =
                    i.speciesSIZ + i.additionalMeans;

                  cellSheetThree(`L${tableThreeStart}`).value =
                    i.issuanceRate + i.AdditionalIssuanceRate;

                  cellSheetThree(`A${tableThreeStart}`).style =
                    cellSheetThree(`A15`).style;
                  cellSheetThree(`B${tableThreeStart}`).style =
                    cellSheetThree(`A15`).style;
                  cellSheetThree(`C${tableThreeStart}`).style =
                    cellSheetThree(`A15`).style;
                  cellSheetThree(`D${tableThreeStart}`).style =
                    cellSheetThree(`A15`).style;
                  cellSheetThree(`E${tableThreeStart}`).style =
                    cellSheetThree(`A15`).style;
                  cellSheetThree(`F${tableThreeStart}`).style =
                    cellSheetThree(`A15`).style;
                  cellSheetThree(`G${tableThreeStart}`).style =
                    cellSheetThree(`A15`).style;
                  cellSheetThree(`H${tableThreeStart}`).style =
                    cellSheetThree(`A15`).style;
                  cellSheetThree(`I${tableThreeStart}`).style =
                    cellSheetThree(`A15`).style;
                  cellSheetThree(`J${tableThreeStart}`).style =
                    cellSheetThree(`A15`).style;
                  cellSheetThree(`K${tableThreeStart}`).style =
                    cellSheetThree(`A15`).style;

                  cellSheetThree(`L${tableThreeStart}`).style =
                    cellSheetThree(`A15`).style;
                  tableThreeStart++;
                }
              });

              res.setHeader(
                'Content-Type',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
              );
              res.setHeader(
                'Content-Disposition',
                `attachment; filename="${Date.now()}_My_Workbook.xlsx"`
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
      ent.owner.toString() === req.user._id ||
      ent.access.includes(req.user._id)
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
        }
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

              sheet.getCell('B8').value = ent.enterprise;
              sheet.getCell('K3').value = ent.chairman;
              sheet.getCell('A14').value = ent.member1Job;
              sheet.getCell('F14').value = ent.member1;
              sheet.getCell('A16').value = ent.member2Job;
              sheet.getCell('F16').value = ent.member2;
              sheet.getCell('A18').value = ent.member3Job;
              sheet.getCell('F18').value = ent.member3;
              sheet.getCell('A20').value = ent.member4Job;
              sheet.getCell('F20').value = ent.member4;

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
                      (l) =>
                        l.source.toLocaleLowerCase() ===
                          o.source.toLocaleLowerCase() &&
                        l.dangerEventID === o.dangerEventID
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
                    if (
                      c.source.toLocaleLowerCase() ===
                        i.source.toLocaleLowerCase() &&
                      c.dangerEventID === i.dangerEventID
                    ) {
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
                    if (
                      c.source.toLocaleLowerCase() ===
                        i.source.toLocaleLowerCase() &&
                      c.dangerEvent776Id === i.dangerEvent776Id
                    ) {
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
                  const startRow = index + 12;
                  // Работники
                  const a =
                    item.veryLowWorker +
                    item.lowWorker +
                    item.midWorker +
                    item.highWorker +
                    item.criticalWorker;
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
                  item.vl = Math.ceil(item.vlp * item.numWorkers);
                  item.l = Math.ceil(item.lp * item.numWorkers);
                  item.m = Math.ceil(item.mp * item.numWorkers);
                  item.h = Math.ceil(item.hp * item.numWorkers);
                  item.c = Math.ceil(item.cp * item.numWorkers);
                  // Рабочие места
                  const b =
                    item.veryLowPlace +
                    item.lowPlace +
                    item.midPlace +
                    item.highPlace +
                    item.criticalPlace;
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
                  item.pm =
                    item.pmp > 0 && item.pmp < 1
                      ? 1
                      : Math.round(item.pmp * item.countWorkPlaces);
                  item.ph = Math.round(item.php * item.countWorkPlaces);
                  item.pc = Math.round(item.pcp * item.countWorkPlaces);
                  sheet.getCell(`A${startRow}`).value = index + 1;
                  sheet.getCell(`B${startRow}`).value = item.source;
                  sheet.getCell(`C${startRow}`).value =
                    item.dangerGroupId || item.danger776Id;
                  sheet.getCell(`D${startRow}`).value =
                    item.dangerGroup || item.danger776;
                  sheet.getCell(`E${startRow}`).value =
                    item.dangerEventID || item.dangerEvent776Id;
                  sheet.getCell(`F${startRow}`).value =
                    item.dangerEvent || item.dangerEvent776;
                  sheet.getCell(
                    `G${startRow}`
                  ).value = `${item.numWorkers}/${item.countWorkPlaces}`;
                  sheet.getCell(
                    `H${startRow}`
                  ).value = `${item.vl}/${item.pvl}`;
                  sheet.getCell(`I${startRow}`).value = `${item.l}/${item.pl}`;
                  sheet.getCell(`J${startRow}`).value = `${item.m}/${item.pm}`;
                  sheet.getCell(`K${startRow}`).value = `${item.h}/${item.ph}`;
                  sheet.getCell(`L${startRow}`).value = `${item.c}/${item.pc}`;
                  sheet.getCell(`M${startRow}`).value = item.IPR;
                  sheet.getCell(`A${startRow}`).style = style;
                  sheet.getCell(`B${startRow}`).style = style;
                  sheet.getCell(`C${startRow}`).style = style;
                  sheet.getCell(`D${startRow}`).style = style;
                  sheet.getCell(`E${startRow}`).style = style;
                  sheet.getCell(`F${startRow}`).style = style;
                  sheet.getCell(`G${startRow}`).style = style;
                  sheet.getCell(`H${startRow}`).style = style;
                  sheet.getCell(`I${startRow}`).style = style;
                  sheet.getCell(`J${startRow}`).style = style;
                  sheet.getCell(`K${startRow}`).style = style;
                  sheet.getCell(`L${startRow}`).style = style;
                  sheet.getCell(`M${startRow}`).style = style;
                  if (sheet.getCell(`H${startRow}`).value !== '0/0') {
                    diagrammaValue.veryLow += 1;
                    sheet.getCell(`H${startRow}`).style = {
                      ...(sheet.getCell(`H${startRow}`).style || {}),
                      fill: darkGeen,
                    };
                  }
                  if (sheet.getCell(`I${startRow}`).value !== '0/0') {
                    diagrammaValue.low += 1;
                    sheet.getCell(`I${startRow}`).style = {
                      ...(sheet.getCell(`I${startRow}`).style || {}),
                      fill: green,
                    };
                  }
                  if (sheet.getCell(`J${startRow}`).value !== '0/0') {
                    diagrammaValue.mid += 1;
                    sheet.getCell(`J${startRow}`).style = {
                      ...(sheet.getCell(`J${startRow}`).style || {}),
                      fill: yellow,
                    };
                  }
                  if (sheet.getCell(`K${startRow}`).value !== '0/0') {
                    diagrammaValue.height += 1;
                    sheet.getCell(`K${startRow}`).style = {
                      ...(sheet.getCell(`K${startRow}`).style || {}),
                      fill: orange,
                    };
                  }
                  if (sheet.getCell(`L${startRow}`).value !== '0/0') {
                    diagrammaValue.critical += 1;
                    sheet.getCell(`L${startRow}`).style = {
                      ...(sheet.getCell(`L${startRow}`).style || {}),
                      fill: red,
                    };
                  }
                  sheet.insertRow(index + 13);
                });
              sheetDiagramma.getCell('B3').value = diagrammaValue.veryLow;
              sheetDiagramma.getCell('B4').value = diagrammaValue.low;
              sheetDiagramma.getCell('B5').value = diagrammaValue.mid;
              sheetDiagramma.getCell('B6').value = diagrammaValue.height;
              sheetDiagramma.getCell('B7').value = diagrammaValue.critical;

              res.setHeader(
                'Content-Type',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
              );
              res.setHeader(
                'Content-Disposition',
                `attachment; filename="${Date.now()}_My_Workbook.xlsx"`
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
//Таблица Перечень СИЗ
module.exports.createListSiz = (req, res, next) => {
  const fileName = 'ПЕРЕЧЕНЬ СИЗ.xlsx';
  Enterprise.findById(req.params.id).then((ent) => {
    if (!ent) {
      next(new NotFound('Предприятие не найдено'));
    }
    if (
      ent.owner.toString() === req.user._id ||
      ent.access.includes(req.user._id)
    ) {
      Value.find({ enterpriseId: req.params.id })
        .then((el) => {
          workbook.xlsx
            .readFile(fileName)
            .then((e) => {
              let startRow = 13;
              const sheet = e.getWorksheet(1);
              const cell = (c, i) => sheet.getCell(c + i);
              cell('B', 10).value = ent.enterprise;
              cell('H', 3).value = ent.chairman;
              cell('A', 15).value = ent.member1Job;
              cell('D', 15).value = ent.member1;
              cell('A', 17).value = ent.member2Job;
              cell('D', 17).value = ent.member2;
              cell('A', 19).value = ent.member3Job;
              cell('D', 19).value = ent.member3;
              cell('A', 21).value = ent.member4Job;
              cell('D', 21).value = ent.member4;

              el.forEach((s) => {
                if (s.typeSIZ) {
                  cell('A', startRow).value = s.num;
                  cell('B', startRow).value = s.proffId;
                  cell('C', startRow).value = s.prof || s.job;
                  cell('D', startRow).value = s.subdivision;
                  cell('E', startRow).value = s.dangerEventID;
                  cell('F', startRow).value = s.typeSIZ;
                  cell('G', startRow).value = s.speciesSIZ;
                  cell('H', startRow).value = s.issuanceRate;
                  cell('A', startRow).border = borderCell;
                  cell('B', startRow).border = borderCell;
                  cell('C', startRow).border = borderCell;
                  cell('D', startRow).border = borderCell;
                  cell('E', startRow).border = borderCell;
                  cell('F', startRow).border = borderCell;
                  cell('G', startRow).border = borderCell;
                  cell('H', startRow).border = borderCell;
                  startRow += 1;
                  sheet.insertRow(startRow);
                }
              });
              res.setHeader(
                'Content-Type',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
              );
              res.setHeader(
                'Content-Disposition',
                `attachment; filename="${Date.now()}_My_Workbook.xlsx"`
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
        action: `Пользователь ${req.user.name} выгрузил(а) таблицу Перечень СИЗ  ${ent.enterprise}`,
        userId: req.user._id,
        enterpriseId: ent._id,
      })
      .catch((e) => next(e));
  });
};

// Таблица Соотнесение опасностей
module.exports.createCorrelationOfHazards = async (req, res, next) => {
  const fileName = 'Книга1.xlsx';
  const enterprise = await Enterprise.findById(req.params.id);
  const owner =
    enterprise.owner.toString() != req.user._id ||
    !enterprise.access.includes(req.user._id);

  if (!enterprise) next(new NotFound('Предприятие не найдено'));
  if (!owner) next(new ConflictError('Нет доступа'));

  const value = await Value.find(
    { enterpriseId: req.params.id },
    {
      dangerGroupId: 1,
      dangerGroup: 1,
      dangerEventID: 1,
      dangerEvent: 1,
      num: 1,
    }
  ).catch((err) => next(err));
  const uniqEventId = [...new Set(value.map((i) => i.dangerEventID))];
  const arr = [];
  uniqEventId.forEach((i) => {
    const obj = { num: '' };
    const filterd = value.filter((f) => f.dangerEventID === i);
    obj.dangerGroupId = filterd[0].dangerGroupId;
    obj.dangerGroup = filterd[0].dangerGroup;
    obj.dangerEventID = filterd[0].dangerEventID;
    obj.dangerEvent = filterd[0].dangerEvent;
    filterd.forEach((c) => {
      obj.num = obj.num.concat(c.num, '; ');
    });
    arr.push(obj);
  });
  await workbook.xlsx
    .readFile(fileName)

    .catch((err) => next(err));

  const sheet = workbook.getWorksheet(1);

  const cell = (literal, index) => sheet.getCell(literal + index);
  let startRow = 16;

  const TEXT_CELL_10 =
    'Приказ Минтруда РФ от 29.10.2021 №767н "Об утверждении Единых типовых норм выдачи средств индивидуальной защиты и смывающих средств.';
  cell('I', 6).value = enterprise.chairman;
  arr.forEach((i) => {
    const number = startRow - 15;
    cell('A', startRow).value = number;
    cell('F', startRow).value = i.dangerGroupId;
    cell('G', startRow).value = i.dangerGroup;
    cell('H', startRow).value = i.dangerEventID;
    cell('I', startRow).value = i.dangerEvent;
    cell('J', startRow).value = TEXT_CELL_10;
    cell('K', startRow).value = i.num;

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
    cell('K', startRow).style = style;
    startRow++;
    sheet.insertRow(startRow);
  });

  res.set(
    'Content-Type',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );
  res.set(
    'Content-Disposition',
    `attachment; filename="${Date.now()}_My_Workbook.xlsx"`
  );
  await workbook.xlsx
    .write(res)
    .then(() => {
      res.end();
    })
    .catch((err) => next(err));
  logs
    .create({
      action: `Пользователь ${req.user.name} выгрузил(а) таблицу Соотнесение  опасностей.xls  ${enterprise.enterprise}`,
      userId: req.user._id,
      enterpriseId: enterprise._id,
    })
    .catch((e) => next(e));
};
