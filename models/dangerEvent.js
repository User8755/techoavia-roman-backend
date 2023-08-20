const mongoose = require('mongoose');
// опасные события
const dangerEvent = new mongoose.Schema({
  label: {
    type: String,
    required: true,
    minlength: 2,
  },
  dependence: {
    type: String,
    required: true,
    minlength: 2,
  },
});

module.exports = mongoose.model('dangerEvent', dangerEvent);
