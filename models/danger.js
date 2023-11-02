const mongoose = require('mongoose');
// виды опасности
const danger = new mongoose.Schema({
  label: {
    type: String,
    required: true,
    minlength: 2,
  },
  groupId: {
    type: String,
    required: true,
    minlength: 1,
    unique: true,
  },
  dependence: {
    type: String,
    required: true,
    minlength: 2,
  },
});

module.exports = mongoose.model('danger', danger);
