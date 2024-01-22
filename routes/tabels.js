const router = require('express').Router();
const { createBaseTabel, createNormTabel } = require('../controllers/tabels');
const auth = require('../middlewares/auth');

router.get('/base/:id', auth, createBaseTabel);
router.get('/norm/:id', auth, createNormTabel);

module.exports = router;
