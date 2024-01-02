const router = require('express').Router();
const { createBranch, getBranch } = require('../controllers/update');

router.post('/branch', createBranch);
router.get('/branch', getBranch);

module.exports = router;
