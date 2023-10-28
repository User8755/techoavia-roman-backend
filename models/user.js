const mongoose = require('mongoose');
const validator = require('validator');

const user = new mongoose.Schema({
  name: {
    type: String,
    require: true,
    minlength: 2,
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
});

module.exports = mongoose.model('user', user);
