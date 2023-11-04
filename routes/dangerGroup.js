const router = require('express').Router();
const { getDangerGroup, createDangerGroup, delDangerGroup } = require('../controllers/dangerGroup');
const { validationDangerGroup } = require('../middlewares/validation');
const auth = require('../middlewares/auth');

router.get('/', auth, getDangerGroup);
router.post('/', auth, validationDangerGroup, createDangerGroup);
router.delete('/:id', auth, delDangerGroup);

module.exports = router;
