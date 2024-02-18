const router = require('express').Router();

const { newValue, updateValue, getValueEnterprise } = require('../controllers/enterpriceValue');

router.post('/:id', newValue);
router.put('/:id', updateValue);
router.get('/:id', getValueEnterprise);
module.exports = router;
