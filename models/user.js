const mongoose = require('mongoose');
const validator = require('validator');

const user = new mongoose.Schema(
  {
    name: {
      type: String,
      require: true,
      minlength: 2,
      validate: /[А-Яа-яЁё]+/,
    },
    family: {
      type: String,
      require: true,
      minlength: 2,
    },
    password: {
      type: String,
      minlength: 4,
      select: false,
      default: '111111',
    },
    email: {
      type: String,
      select: false,
      required: true,
      unique: true,
      validate: {
        validator: (v) => validator.isEmail(v),
        message: 'Не верный email',
      },
    },
    role: {
      type: Array,
      default: ['user'],
      require: true,
    },
    branch: {
      type: String,
      require: true,
    },
    login: {
      type: String,
      require: true,
      unique: true,
      select: false,
    },
  },
  { versionKey: false },
);

module.exports = mongoose.model('user', user);
