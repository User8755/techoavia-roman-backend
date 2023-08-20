const router = require('express').Router();
const { getDangerGroup, createDangerGroup } = require('../controllers/dangerGroup');

router.get('/dangerGroup', getDangerGroup);
router.post('/dangerGroup', createDangerGroup);

module.exports = router;
