const router = require('express').Router();

const { createInfo, getBranch } = require('../controllers/info');
const auth = require('../middlewares/auth');

router.post('/', auth, createInfo);
router.get('/', auth, getBranch);

module.exports = router;
