const mongoose = require('mongoose');

const enterprise = new mongoose.Schema({
  enterprise: {
    type: String,
    require: true,
    minlength: 2,
  },
  owner: {
    type: mongoose.Schema.Types.ObjectId,
    ref: 'user',
    require: true,
  },
  value: {
    type: Array,
    default: [],
  },
  access: {
    type: Array,
    default: [],
  },
  inn: {
    type: String,
    require: true,
    minlength: 2,
  },
  kpp: {
    type: String,
    require: true,
    minlength: 2,
  },
  order: {
    type: String,
    require: true,
    minlength: 2,
    unique: true,
  },
});

module.exports = mongoose.model('enterprise', enterprise);
