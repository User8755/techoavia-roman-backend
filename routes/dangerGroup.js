const router = require('express').Router();
const { getDangerGroup, createDangerGroup, delDangerGroup } = require('../controllers/dangerGroup');

router.get('/dangerGroup', getDangerGroup);
router.post('/dangerGroup', createDangerGroup);
router.delete('/dangerGroup/:id', delDangerGroup);

module.exports = router;
