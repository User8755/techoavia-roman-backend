const router = require('express').Router();

const { getLogs } = require('../controllers/logs');
const auth = require('../middlewares/auth');

router.get('/', auth, getLogs);

module.exports = router;
