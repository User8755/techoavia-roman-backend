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
    inn: Joi.string().required().min(2),
    kpp: Joi.string().required().min(2),
    order: Joi.string().required().min(2),
    chairman: Joi.string().min(0).max(30),
    chairmanJob: Joi.string().min(0).max(30),
    member1: Joi.string().min(0).max(30),
    member1Job: Joi.string().min(0).max(30),
    member2: Joi.string().min(0).max(30),
    member2Job: Joi.string().min(0).max(30),
    member3: Joi.string().min(0).max(30),
    member3Job: Joi.string().min(0).max(30),
    member4: Joi.string().min(0).max(30),
    member4Job: Joi.string().min(0).max(30),
    member5: Joi.string().min(0).max(30),
    member5Job: Joi.string().min(0).max(30),
    member6: Joi.string().min(0).max(30),
    member6Job: Joi.string().min(0).max(30),
  }),
});

module.exports.validationEnterpriseValue = celebrate({
  body: Joi.object().keys({
    proff: Joi.string().default('').min(0).max(200),
    proffId: Joi.number().min(1).max(1000),
    danger: Joi.string().default('').min(0),
    dangerID: Joi.string().default('').min(0),
    dangerGroup: Joi.string().default('').min(0),
    dangerGroupId: Joi.string().default('').min(0),
    dangerEvent: Joi.string().default('').min(0),
    dangerEventID: Joi.string().default('').min(0),
    ipr: Joi.number(),
    riskAttitude: Joi.string().default('').min(0),
    risk: Joi.string().default('').min(0),
    acceptability: Joi.string().default('').min(0),
    probability1: Joi.number(),
    heaviness1: Joi.number().default(0),
    ipr1: Joi.number(),
    riskAttitude1: Joi.string().default('').min(0),
    risk1: Joi.string().default('').min(0),
    acceptability1: Joi.string().default('').min(0),
    typeSIZ: Joi.string().default('').min(0),
    speciesSIZ: Joi.string().default('').min(0),
    issuanceRate: Joi.string().default('').min(0),
    commit: Joi.string().default('').min(0).min(0),
    proffSIZ: Joi.array(),
    danger776: Joi.string().default('').min(0),
    danger776Id: Joi.string().default('').min(0),
    dangerEvent776: Joi.string().default('').min(0),
    dangerEvent776Id: Joi.string().default('').min(0),
    riskManagement: Joi.string().default('').min(0),
    riskManagementID: Joi.string().default('').min(0),
    standart: Joi.string().default('').min(0),
    OperatingLevel: Joi.string().default('').min(0),
    periodicity: Joi.string().default('').min(0).min(0),
    probability: Joi.number().default(0),
    heaviness: Joi.number().default(0),
    responsiblePerson: Joi.string().default('').min(0),
    completionMark: Joi.string().default('').min(0),
    existingRiskManagement: Joi.string().default('').min(0),
    obj: Joi.string().default('').min(0),
    source: Joi.string().default('').min(0),
    job: Joi.string().default('').min(0),
    subdivision: Joi.string().default('').min(0),
  }),
});
