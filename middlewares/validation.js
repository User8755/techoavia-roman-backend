// eslint-disable-next-line import/no-extraneous-dependencies
const { celebrate, Joi } = require('celebrate');

module.exports.validationCreateUser = celebrate({
  body: Joi.object().keys({
    name: Joi.string().min(2).max(30).required()
      .regex(/[А-Яа-яЁё]+/)
      .message('dsa'),
    family: Joi.string().min(2).max(30).required()
      .regex(/[А-Яа-яЁё]+/),
    email: Joi.string().required().email(),
    password: Joi.string().required(),
    role: Joi.string().default('admin'),
  }),
});

module.exports.validationLogin = celebrate({
  body: Joi.object().keys({
    email: Joi.string().required().email(),
    password: Joi.string().required(),
  }),
});

module.exports.validationDangerGroup = celebrate({
  body: Joi.object().keys({
    label: Joi.string().required().regex(/^[А-Яа-я]/),
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
