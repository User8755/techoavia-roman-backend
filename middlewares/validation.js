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
