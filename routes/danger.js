const router = require('express').Router();
const { getDanger, createDanger } = require('../controllers/danger');
const auth = require('../middlewares/auth');

router.get('/', auth, getDanger);
router.post('/', createDanger);

module.exports = router;
