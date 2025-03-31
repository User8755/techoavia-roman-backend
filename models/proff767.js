const mongoose = require('mongoose');

const proff767 = new mongoose.Schema(
  {
    proffId: {
      type: Number,
      require: true,
      minlength: 1,
      maxlength: 30,
    },
    proff: {
      type: String,
      require: true,
      minlength: 1,
      maxlength: 300,
    },
    speciesSIZ: {
      type: String,
      require: true,
      minlength: 1,
      maxlength: 500,
    },
    typeSIZ: {
      type: String,
      require: true,
      minlength: 1,
      maxlength: 500,
    },
    issuanceRate: {
      type: String,
      require: true,
      minlength: 1,
      maxlength: 500,
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

module.exports = mongoose.model('proff767', proff767);
