const router = require('express').Router();
const auth = require('../middlewares/auth');

const { UpdatesProff767, UpdatesTypeSiz } = require('../controllers/update');

router.post('/proff767', auth, UpdatesProff767);
router.post('/type-siz', auth, UpdatesTypeSiz);

module.exports = router;
