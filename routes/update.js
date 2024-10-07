const router = require('express').Router();
const auth = require('../middlewares/auth');

const {
  UpdatesProff767,
} = require('../controllers/update');

router.post('/proff767', auth, UpdatesProff767);

module.exports = router;
