const mongoose = require('mongoose');

const branch = new mongoose.Schema(
  {
    branch: {
      type: String,
      require: true,
      minlength: 2,
      validate: /[А-Яа-яЁё]+/,
      unique: true,
    },
  },
  { versionKey: false },
);

module.exports = mongoose.model('branch', branch);
