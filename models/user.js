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
      require: true,
      minlength: 6,
      select: false,
    },
    email: {
      type: String,
      required: true,
      unique: true,
      validate: {
        validator: (v) => validator.isEmail(v),
        message: 'Не верный email',
      },
    },
    role: {
      type: String,
      default: 'admin',
      require: true,
    },

  },
  { versionKey: false },
);

module.exports = mongoose.model('user', user);
