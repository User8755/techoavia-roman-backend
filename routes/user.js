const router = require('express').Router();
const { createUsers, login, getUsersСurrent } = require('../controllers/user');
const { validationCreateUser, validationLogin } = require('../middlewares/validation');
const auth = require('../middlewares/auth');

router.post('/signup', validationCreateUser, createUsers);
router.post('/signin', validationLogin, login);
router.get('/me', auth, getUsersСurrent);

module.exports = router;
