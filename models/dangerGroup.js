const mongoose = require('mongoose');

// группа опасности
const dangerGroup = new mongoose.Schema({
  label: {
    type: String,
    required: true,
    minlength: 2,
    validate: /[^{}<>=a-zA-Z]/,
  },
  dangerID: {
    type: String,
    required: true,
    minlength: 1,
    unique: true,
    validate: /^\d+\.?\d*$/,
  },
});

module.exports = mongoose.model('dangerGroup', dangerGroup);
