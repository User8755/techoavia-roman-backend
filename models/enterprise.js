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
  value: {
    type: Array,
    default: [],
  },
});

module.exports = mongoose.model('enterprise', enterprise);
