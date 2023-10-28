const mongoose = require('mongoose');

// группа опасности
const dangerGroup = new mongoose.Schema({
  label: {
    type: String,
    required: true,
    minlength: 2,
    unique: true,
  },
  dangerID: {
    type: Number,
    required: true,
    minlength: 1,
    unique: true,
  },
});

module.exports = mongoose.model('dangerGroup', dangerGroup);
