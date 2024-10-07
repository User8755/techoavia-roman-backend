const router = require('express').Router();
const auth = require('../middlewares/auth');
const checkValue = require('../middlewares/checkValue');

const {
  newValue,
  updateValue,
  getValueEnterprise,
  getUniqWorkPlace,
  getCurentWorkPlace,
  createNewPlace,
  getlastValue,
} = require('../controllers/enterpriceValue');

router.post('/:id', auth, newValue);
router.post('/:id/place/new', auth, createNewPlace);
router.put('/:id', auth, checkValue, updateValue);
router.get('/:id', auth, getValueEnterprise);
router.get('/:id/worker', auth, getUniqWorkPlace);
router.get('/:id/last', auth, getlastValue);
router.post('/:id/worker/curent', auth, getCurentWorkPlace);

module.exports = router;
