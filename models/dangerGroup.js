const mongoose = require('mongoose');
// группа опасности
const dangerGroup = new mongoose.Schema({
  label: {
    type: String,
    required: true,
    minlength: 2,
  },
  dangerID: {
    type: Number,
    required: true,
    minlength: 1,
  },
});

module.exports = mongoose.model('dangerGroup', dangerGroup);
