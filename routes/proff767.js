const router = require('express').Router();

const { getAllproff767 } = require('../controllers/proff');
const auth = require('../middlewares/auth');

router.get('/proff', auth, getAllproff767);

module.exports = router;
