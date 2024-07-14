const router = require('express').Router();
const {
  createBaseTabel,
  createNormTabel,
  createMapOPRTabel,
  createListOfMeasuresTabel,
  createListHazardsTable,
  createPlanTimetable,
  createRegisterHazards,
  createListSiz,
} = require('../controllers/tabels');
const auth = require('../middlewares/auth');

router.get('/base/:id', auth, createBaseTabel);
router.get('/norm/:id', auth, createNormTabel);
router.get('/mapOPR/:id', auth, createMapOPRTabel);
router.get('/listOfMeasures/:id', auth, createListOfMeasuresTabel);
router.get('/listHazards/:id', auth, createListHazardsTable);
router.get('/planTimetable/:id', auth, createPlanTimetable);
router.get('/registerHazards/:id', auth, createRegisterHazards);
router.get('/listSiz/:id', auth, createListSiz);

module.exports = router;
