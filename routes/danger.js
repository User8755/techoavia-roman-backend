const router = require('express').Router();
const { getDanger, createDanger } = require('../controllers/danger');

router.get('/danger', getDanger);
router.post('/danger', createDanger);

module.exports = router;
