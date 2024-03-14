const router = require('express').Router();
const auth = require('../middlewares/auth');

const {
  newValue,
  updateValue,
  getValueEnterprise,
} = require('../controllers/enterpriceValue');

router.post('/:id', auth, newValue);
router.put('/:id', auth, updateValue);
router.get('/:id', auth, getValueEnterprise);

module.exports = router;
