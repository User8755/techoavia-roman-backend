const router = require('express').Router();
const { createUsers, login } = require('../controllers/user');
const { validationCreateUser, validationLogin } = require('../middlewares/validation');

router.post('/signup', validationCreateUser, createUsers);
router.post('/signin', validationLogin, login);
module.exports = router;
