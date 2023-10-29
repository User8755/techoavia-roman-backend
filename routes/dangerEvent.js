const router = require('express').Router();
const { getDangerEvent, createDangerEvent } = require('../controllers/dangerEvent');

router.get('/dangerEvent', getDangerEvent);
router.post('/dangerEvent', createDangerEvent);

module.exports = router;
