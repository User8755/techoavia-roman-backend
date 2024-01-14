const router = require('express').Router();
const {
  createEnterprise,
  getEnterprisesUser,
  getCurrentEnterprise,
  updateCurrentEnterpriseValue,
  updateAccess,
  getEnterprisesAccessUser,
} = require('../controllers/enterprise');
const auth = require('../middlewares/auth');
const {
  validationEnterprise,
  validationEnterpriseValue,
} = require('../middlewares/validation');

router.post('/', auth, validationEnterprise, createEnterprise);
router.get('/', auth, getEnterprisesUser);
router.get('/access', auth, getEnterprisesAccessUser);
router.get('/:id', auth, getCurrentEnterprise);
router.patch(
  '/:id',
  auth,
  validationEnterpriseValue,
  updateCurrentEnterpriseValue,
);
router.patch('/access/:id', auth, updateAccess);

module.exports = router;
