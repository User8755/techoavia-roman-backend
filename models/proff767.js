const mongoose = require('mongoose');

const proff767 = new mongoose.Schema({
  proffId: {
    type: Number,
    require: true,
    minlength: 1,
    maxlength: 30,
  },
  proff: {
    type: String,
    require: true,
    minlength: 1,
    maxlength: 300,
  },
  vid: {
    type: String,
    require: true,
    minlength: 1,
    maxlength: 500,
  },
  type: {
    type: String,
    require: true,
    minlength: 1,
    maxlength: 500,
  },
  norm: {
    type: String,
    require: true,
    minlength: 1,
    maxlength: 500,
  },
}, { versionKey: false });

module.exports = mongoose.model('proff767', proff767);
