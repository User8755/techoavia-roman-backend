const mongoose = require('mongoose');

const value = new mongoose.Schema(
  {
    enterpriseId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'enterprise',
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
    obj: {
      type: String,
      minlength: 0,
      maxlength: 300,
      required: true,
    },
    source: {
      type: String,
      minlength: 1,
      maxlength: 300,
      required: true,
    },
    dangerID: {
      type: String,
      minlength: 0,
      maxlength: 10,
    },
    danger: {
      type: String,
      minlength: 0,
      maxlength: 300,
    },
    dangerGroupId: {
      type: String,
      minlength: 0,
      maxlength: 10,
    },
    dangerGroup: {
      type: String,
      minlength: 0,
      maxlength: 300,
    },
    dangerEventID: {
      type: String,
      minlength: 0,
      maxlength: 10,
    },
    dangerEvent: {
      type: String,
      minlength: 0,
      maxlength: 400,
    },
    heaviness: {
      type: Number,
      minlength: 1,
      maxlength: 2,
    },
    probability: {
      type: Number,
      minlength: 1,
      maxlength: 2,
    },
    ipr: {
      type: Number,
      minlength: 1,
      maxlength: 2,
    },
    risk: {
      type: String,
      minlength: 2,
      maxlength: 20,
    },
    acceptability: {
      type: String,
      minlength: 1,
      maxlength: 50,
    },
    riskAttitude: {
      type: String,
    },
    typeSIZ: {
      type: String,
    },
    speciesSIZ: {
      type: String,
    },
    issuanceRate: {
      type: String,
    },
    additionalMeans: {
      type: String,
    },
    AdditionalIssuanceRate: {
      type: String,
    },
    standart: {
      type: String,
    },
    OperatingLevel: {
      type: String,
    },
    commit: {
      type: String,
    },
    danger776Id: {
      type: String,
    },
    danger776: {
      type: String,
    },
    dangerEvent776Id: {
      type: String,
    },
    dangerEvent776: {
      type: String,
    },
    riskManagementID: {
      type: String,
    },
    riskManagement: {
      type: String,
    },
    heaviness1: {
      type: Number,
    },
    probability1: {
      type: Number,
    },
    ipr1: {
      type: Number,
    },
    risk1: {
      type: String,
    },
    acceptability1: {
      type: String,
    },
    riskAttitude1: {
      type: String,
    },
    existingRiskManagement: {
      type: String,
    },
    periodicity: {
      type: String,
    },
    responsiblePerson: {
      type: String,
    },
    completionMark: {
      type: String,
    },
    proffSIZ: {
      type: Array,
    },
    numWorkers: {
      type: String,
    },
    equipment: {
      type: String,
    },
    materials: {
      type: String,
    },
    laborFunction: {
      type: String,
    },
    code: {
      type: String,
    },
  },
  { versionKey: false },
);

module.exports = mongoose.model('value', value);
