const mongoose = require('mongoose');

const workPlace = new mongoose.Schema(
  {
    enterpriseId: {
      type: String,
      require: true,
    },
    proffId: {
      type: Number,
      minlength: 0,
      maxlength: 30,
    },
    num: {
      type: String,
      required: true,
      minlength: 1,
      maxlength: 30,
    },
    proff: {
      type: String,
      minlength: 0,
      maxlength: 300,
    },
    job: {
      type: String,
      minlength: 0,
      maxlength: 300,
    },
    subdivision: {
      type: String,
      minlength: 0,
      maxlength: 300,
    },
    proffSIZ: {
      type: [
        {
          vid: {
            type: String,
            maxlength: 500,
          },
          type: {
            type: String,
            maxlength: 500,
          },
          norm: {
            type: String,
            maxlength: 500,
          },
        },
      ],
      default: [],
      required: true,
    },
    numWorkers: {
      type: String,
    },
    code: {
      type: String,
    },
  },
  { versionKey: false }
);

module.exports = mongoose.model('workPlace', workPlace);
