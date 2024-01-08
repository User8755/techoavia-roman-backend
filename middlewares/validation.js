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
    proff: Joi.string().default(''),
    proffId: Joi.number(),
    danger: Joi.string().default(''),
    dangerID: Joi.string().default(''),
    dangerGroup: Joi.string().default(''),
    dangerGroupId: Joi.string().default(''),
    dangerEvent: Joi.string().default(''),
    dangerEventID: Joi.string().default(''),
    ipr: Joi.number(),
    riskAttitude: Joi.string().default(''),
    risk: Joi.string().default(''),
    acceptability: Joi.string().default(''),
    probability1: Joi.number().default(0),
    heaviness1: Joi.number().default(0),
    ipr1: Joi.number(),
    riskAttitude1: Joi.string().default(''),
    risk1: Joi.string().default(''),
    acceptability1: Joi.string().default(''),
    typeSIZ: Joi.string().default(''),
    speciesSIZ: Joi.string().default(''),
    issuanceRate: Joi.string().default(''),
    commit: Joi.string().default(''),
    proffSIZ: Joi.string().default(''),
    danger776: Joi.string().default(''),
    danger776Id: Joi.string().default(''),
    dangerEvent776: Joi.string().default(''),
    dangerEvent776Id: Joi.string().default(''),
    riskManagement: Joi.string().default(''),
    riskManagementID: Joi.string().default(''),
    standart: Joi.string().default(''),
    OperatingLevel: Joi.string().default(''),
    periodicity: Joi.string().default(''),
    probability: Joi.number().default(0),
    heaviness: Joi.number().default(0),
    responsiblePerson: Joi.string().default(''),
    completionMark: Joi.string().default(''),
    existingRiskManagement: Joi.string().default(''),
    obj: Joi.string().default(''),
    source: Joi.string().default(''),
    job: Joi.string().default(''),
    subdivision: Joi.string().default(''),
  }),
});
