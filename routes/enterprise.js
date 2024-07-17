const router = require('express').Router();
const {
  createEnterprise,
  getEnterprisesUser,
  getCurrentEnterprise,
  updateAccess,
  getEnterprisesAccessUser,
  updateCloseAccess,
  statusHiden,
} = require('../controllers/enterprise');
const auth = require('../middlewares/auth');
const {
  validationEnterprise,
} = require('../middlewares/validation');

router.post('/', auth, validationEnterprise, createEnterprise);
router.get('/', auth, getEnterprisesUser);
router.get('/access', auth, getEnterprisesAccessUser);
router.get('/:id', auth, getCurrentEnterprise);
router.patch('/access/:id', auth, updateAccess);
router.delete('/access/:id', auth, updateCloseAccess);
router.post('/status/:id', auth, statusHiden);

module.exports = router;
