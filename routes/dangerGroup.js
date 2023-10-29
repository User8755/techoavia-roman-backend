const router = require('express').Router();
const { getDangerGroup, createDangerGroup, delDangerGroup } = require('../controllers/dangerGroup');
const { validationDangerGroup } = require('../middlewares/validation');

router.get('/dangerGroup', getDangerGroup);
router.post('/dangerGroup', validationDangerGroup, createDangerGroup);
router.delete('/dangerGroup/:id', delDangerGroup);

module.exports = router;
