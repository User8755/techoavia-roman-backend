const router = require('express').Router();
const { getDangerEvent, createDangerEvent } = require('../controllers/dangerEvent');
const auth = require('../middlewares/auth');

router.get('/', auth, getDangerEvent);
router.post('/', auth, createDangerEvent);

module.exports = router;
