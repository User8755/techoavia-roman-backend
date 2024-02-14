const router = require('express').Router();
const {
  createBaseTabel,
  createNormTabel,
  createMapOPRTabel,
  createListOfMeasuresTabel,
  updateEnterpriceValue,
} = require('../controllers/tabels');
const auth = require('../middlewares/auth');

router.get('/base/:id', auth, createBaseTabel);
router.get('/norm/:id', auth, createNormTabel);
router.get('/mapOPR/:id', auth, createMapOPRTabel);
router.get('/listOfMeasures/:id', auth, createListOfMeasuresTabel);
router.post('/updateBaseTabel/:id', updateEnterpriceValue);

module.exports = router;
