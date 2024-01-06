const mongoose = require('mongoose');

const info = new mongoose.Schema({
  info: {
    type: String,
    require: true,
    minlength: 2,
    default: 'nnnnnn',
  },
});

module.exports = mongoose.model('info', info);
