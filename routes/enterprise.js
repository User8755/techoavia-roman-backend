const router = require('express').Router();
const {
  createEnterprise,
  getEnterprisesUser,
  getCurrentEnterprise,
  updateCurrentEnterpriseValue,
} = require('../controllers/enterprise');
const auth = require('../middlewares/auth');
const { validationEnterprise } = require('../middlewares/validation');

router.post('/', auth, validationEnterprise, createEnterprise);
router.get('/', auth, getEnterprisesUser);
router.get('/:id', auth, getCurrentEnterprise);
router.patch('/:id', auth, updateCurrentEnterpriseValue);

module.exports = router;
