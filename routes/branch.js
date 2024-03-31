const router = require('express').Router();
const { createBranch, getBranch } = require('../controllers/branch');

router.post('/', createBranch);
router.get('/', getBranch);

module.exports = router;
