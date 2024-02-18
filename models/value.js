const mongoose = require('mongoose');

const value = new mongoose.Schema(
  {
    enterpriseId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'enterprise',
    },
    proffId: {
      type: Number,
    },
    num: {
      type: String,
    },
    proff: {
      type: String,
    },
    job: {
      type: String,
    },
    subdivision: {
      type: String,
    },
    obj: {
      type: String,
    },
    source: {
      type: String,
    },
    dangerID: {
      type: String,
    },
    danger: {
      type: String,
    },
    dangerGroupId: {
      type: String,
    },
    dangerGroup: {
      type: String,
    },
    dangerEventID: {
      type: String,
    },
    dangerEvent: {
      type: String,
    },
    heaviness: {
      type: Number,
    },
    probability: {
      type: Number,
    },
    ipr: {
      type: Number,
    },
    risk: {
      type: String,
    },
    acceptability: {
      type: String,
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
    SIZ: {
      type: Array,
    },
  },
  { versionKey: false },
);

module.exports = mongoose.model('value', value);
