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
  isHiden: {
    type: Boolean,
    default: false,
  },
});

module.exports = mongoose.model('enterprise', enterprise);
