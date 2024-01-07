// eslint-disable-next-line import/no-extraneous-dependencies
const { celebrate, Joi } = require('celebrate');

module.exports.validationCreateUser = celebrate({
  body: Joi.object().keys({
    name: Joi.string().min(2).max(30).required(),
    family: Joi.string().min(2).max(30).required(),
    email: Joi.string().required().email(),
    password: Joi.string(),
    role: Joi.string().default('user'),
    branch: Joi.string().required(),
    login: Joi.string().required(),
  }),
});

module.exports.validationLogin = celebrate({
  body: Joi.object().keys({
    login: Joi.string().required(),
    password: Joi.string().required(),
  }),
});

module.exports.validationDangerGroup = celebrate({
  body: Joi.object().keys({
    label: Joi.string()
      .required()
      .regex(/^[А-Яа-я]/),
    dangerID: Joi.string().required(),
  }),
});

module.exports.validationDanger = celebrate({
  body: Joi.object().keys({
    label: Joi.string().required(),
    dependence: Joi.string().required(),
    groupId: Joi.string().required(),
  }),
});

module.exports.validationDangerEvent = celebrate({
  body: Joi.object().keys({
    label: Joi.string().required(),
    dependence: Joi.string().required(),
    groupId: Joi.string().required(),
  }),
});

module.exports.validationEnterprise = celebrate({
  body: Joi.object().keys({
    enterprise: Joi.string().required().min(2),
  }),
});

module.exports.validationEnterpriseValue = celebrate({
  body: Joi.object().keys({
    proff: Joi.string(),
    proffId: Joi.string(),
    danger: Joi.string(),
    dangerID: Joi.string(),
    dangerGroup: Joi.string(),
    dangerGroupId: Joi.string(),
    dangerEvent: Joi.string(),
    dangerEventID: Joi.string(),
    ipr: Joi.string(),
    riskAttitude: Joi.string(),
    risk: Joi.string(),
    acceptability: Joi.string(),
    probability1: Joi.string(),
    heaviness1: Joi.string(),
    ipr1: Joi.string(),
    riskAttitude1: Joi.string(),
    risk1: Joi.string(),
    acceptability1: Joi.string(),
    typeSIZ: Joi.string(),
    speciesSIZ: Joi.string(),
    issuanceRate: Joi.string(),
    commit: Joi.string(),
    proffSIZ: Joi.string(),
    danger776: Joi.string(),
    danger776Id: Joi.string(),
    dangerEvent776: Joi.string(),
    dangerEvent776Id: Joi.string(),
    riskManagement: Joi.string(),
    riskManagementID: Joi.string(),
    standart: Joi.string(),
    OperatingLevel: Joi.string(),
    periodicity: Joi.string(),
    probability: Joi.string(),
    heaviness: Joi.string(),
    responsiblePerson: Joi.string(),
    completionMark: Joi.string(),
    existingRiskManagement: Joi.string(),
    obj: Joi.string(),
    source: Joi.string(),
    job: Joi.string(),
    subdivision: Joi.string(),
  }),
});
