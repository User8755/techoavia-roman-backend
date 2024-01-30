const Enterprise = require('../models/enterprise');
const ConflictError = require('../errors/ConflictError');

module.exports.createEnterprise = (req, res, next) => {
  const {
    enterprise, inn, kpp, order,
  } = req.body;
  Enterprise.create({
    enterprise, inn, kpp, order, owner: req.user._id,
  })
    .then((i) => res.send(i))
    .catch((e) => next(e));
};

module.exports.getEnterprisesUser = (req, res, next) => {
  Enterprise.find({ owner: req.user._id })
    .then((i) => {
      res.send(i);
    })
    .catch((e) => next(e));
};

module.exports.getEnterprisesAccessUser = (req, res, next) => {
  Enterprise.find({ access: req.user._id })
    .then((i) => {
      res.send(i);
    })
    .catch((e) => next(e));
};
module.exports.getCurrentEnterprise = (req, res, next) => {
  Enterprise.findOne({ _id: req.params.id })
    .then((i) => {
      if (i.owner.toString() === req.user._id || i.access.includes(req.user._id)) {
        res.send(i);
      } else {
        next(new ConflictError('Нет доступа'));
      }
    })
    .catch((e) => next(e));
};

module.exports.updateCurrentEnterpriseValue = (req, res, next) => {
  Enterprise.findByIdAndUpdate(
    req.params.id,
    {
      $addToSet: {
        value: {
          proff: req.body.proff,
          proffId: req.body.proffId,
          danger: req.body.danger,
          dangerID: req.body.dangerID,
          dangerGroup: req.body.dangerGroup,
          dangerGroupId: req.body.dangerGroupId,
          dangerEvent: req.body.dangerEvent,
          dangerEventID: req.body.dangerEventID,
          ipr: req.body.ipr,
          riskAttitude: req.body.riskAttitude,
          risk: req.body.risk,
          acceptability: req.body.acceptability,
          probability1: req.body.probability1,
          heaviness1: req.body.heaviness1,
          ipr1: req.body.ipr1,
          riskAttitude1: req.body.riskAttitude1,
          risk1: req.body.risk1,
          acceptability1: req.body.acceptability1,
          typeSIZ: req.body.typeSIZ,
          speciesSIZ: req.body.speciesSIZ,
          issuanceRate: req.body.issuanceRate,
          commit: req.body.commit,
          proffSIZ: req.body.proffSIZ,
          danger776: req.body.danger776,
          danger776Id: req.body.danger776Id,
          dangerEvent776: req.body.dangerEvent776,
          dangerEvent776Id: req.body.dangerEvent776Id,
          riskManagement: req.body.riskManagement,
          riskManagementID: req.body.riskManagementID,
          standart: req.body.standart,
          OperatingLevel: req.body.OperatingLevel,
          periodicity: req.body.periodicity,
          probability: req.body.probability,
          heaviness: req.body.heaviness,
          responsiblePerson: req.body.responsiblePerson,
          completionMark: req.body.completionMark,
          existingRiskManagement: req.body.existingRiskManagement,
          obj: req.body.obj,
          source: req.body.source,
          job: req.body.job,
          subdivision: req.body.subdivision,
        },
      },
    },
    { new: true },
  )
    .then((value) => res.send(value))
    .catch((e) => next(e));
};

module.exports.updateAccess = (req, res, next) => {
  const { user } = req.body;
  Enterprise.findByIdAndUpdate(
    req.params.id,
    { $push: { access: user } },
    { new: true },
  )
    .then((i) => {
      res.send(i);
    })
    .catch((e) => next(e));
};

module.exports.updateCloseAccess = (req, res, next) => {
  const { user } = req.body;
  Enterprise.findByIdAndUpdate(
    req.params.id,
    { $pull: { access: user } },
    { new: true },
  )
    .then((i) => {
      res.send(i);
    })
    .catch((e) => next(e));
};
