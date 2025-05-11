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
  chairman: {
    type: String,
    maxlength: 30,
  },
  chairmanJob: {
    type: String,
    maxlength: 30,
  },
  member1: {
    type: String,
    maxlength: 30,
  },
  member1Job: {
    type: String,
    maxlength: 30,
  },
  member2: {
    type: String,
    maxlength: 30,
  },
  member2Job: {
    type: String,
    maxlength: 30,
  },
  member3: {
    type: String,
    maxlength: 30,
  },
  member3Job: {
    type: String,
    maxlength: 30,
  },
  member4: {
    type: String,
    maxlength: 30,
  },
  member4Job: {
    type: String,
    maxlength: 30,
  },
  member5: {
    type: String,
    maxlength: 30,
  },
  member5Job: {
    type: String,
    maxlength: 30,
  },
  member6: {
    type: String,
    maxlength: 30,
  },
  member6Job: {
    type: String,
    maxlength: 30,
  },
}, { strict: true });

module.exports = mongoose.model('enterprise', enterprise);
