const router = require('express').Router();
const auth = require('../middlewares/auth');

const {
  getUniqWorkPlace,
  getCurentWorkPlace,
  createWorkPlace,
} = require('../controllers/workPlace');

router.post('/:id/worker/curent', auth, getCurentWorkPlace);
router.post('/:id/worker', auth, createWorkPlace);
router.get('/:id/worker', auth, getUniqWorkPlace);

module.exports = router;
