const mongoose = require('mongoose');

const info = new mongoose.Schema({
  info: {
    type: String,
    require: true,
    minlength: 2,
  },
});

module.exports = mongoose.model('info', info);
