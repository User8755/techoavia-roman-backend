const router = require('express').Router();
const {
  createEnterprise,
  getEnterprisesUser,
  getCurrentEnterprise,
  updateCurrentEnterpriseValue,
  updateAccess,
  getEnterprisesAccessUser,
  updateCloseAccess,
} = require('../controllers/enterprise');
const auth = require('../middlewares/auth');
const {
  validationEnterprise,
} = require('../middlewares/validation');

router.post('/', auth, validationEnterprise, createEnterprise);
router.get('/', auth, getEnterprisesUser);
router.get('/access', auth, getEnterprisesAccessUser);
router.get('/:id', auth, getCurrentEnterprise);
router.patch(
  '/:id',
  auth,
  updateCurrentEnterpriseValue,
);
router.patch('/access/:id', auth, updateAccess);
router.delete('/access/:id', auth, updateCloseAccess);

module.exports = router;
