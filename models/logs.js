const mongoose = require('mongoose');

const logs = new mongoose.Schema({
  action: {
    type: String,
    require: true,
    minlength: 2,
  },
  userId: {
    type: String,
    require: true,
    minlength: 2,
  },
  enterpriseId: {
    type: String,
    require: true,
    minlength: 2,
  },
  date: {
    type: Date,
    default: Date.now,
  },
}, { versionKey: false });

module.exports = mongoose.model('logs', logs);
