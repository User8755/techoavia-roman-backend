const mongoose = require('mongoose');

const typeSiz = new mongoose.Schema(
  {
    dependence: {
      type: String,
      require: true,
    },
    label: {
      type: String,
      require: true,
    },
    speciesSIZ: {
      type: String,
      require: true,
    },
    issuanceRate: {
      type: String,
      require: true,
    },
    additionalMeans: {
      type: String,
      require: true,
    },
    AdditionalIssuanceRate: {
      type: String,
      require: true,
    },
    standart: {
      type: String,
      require: true,
    },
    OperatingLevel: {
      type: String,
      require: true,
    },
    markerBase: {
      type: String,
      maxlength: 3,
    },
    markerRubber: {
      // Резиновое изделие
      type: String,
      maxlength: 3,
    },
    markerSlip: {
      // Скольжение
      type: String,
      maxlength: 3,
    },
    markerPuncture: {
      // прокол обуви
      type: String,
      maxlength: 3,
    },
    markerGlovesAbrasion: {
      // Перчатки истирание
      type: String,
      maxlength: 3,
    },
    markerGlovesCut: {
      // Перчатки порез
      type: String,
      maxlength: 3,
    },
    markerGlovesPuncture: {
      // Перчатки прокол
      type: String,
      maxlength: 3,
    },
    markerWinterShoes: {
      // Зимняя обувь
      type: String,
      maxlength: 3,
    },
    markerWinterclothes: {
      // Зимняя одежда
      type: String,
      maxlength: 3,
    },
    markerHierarchyOfClothing: {
      // Иерархия одежды
      type: String,
      maxlength: 3,
    },
    markerHierarchyOfShoes: {
      // Иерархия обуви
      type: String,
      maxlength: 3,
    },
    markerHierarchyOfGloves: {
      // Иерархия СИЗ рук
      type: String,
      maxlength: 3,
    },
  },
  { versionKey: false }
);

module.exports = mongoose.model('typeSiz', typeSiz);
