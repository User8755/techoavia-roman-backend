const router = require('express').Router();
const {
  createUsers, login, getUsersСurrent, getAllUsers, updateProfile,
} = require('../controllers/user');
const {
  validationCreateUser,
  validationLogin,
} = require('../middlewares/validation');
const auth = require('../middlewares/auth');
const role = require('../middlewares/role');

router.post('/signup', validationCreateUser, createUsers);
router.post('/signin', validationLogin, login);
router.get('/me', auth, getUsersСurrent);
router.get('/all', auth, getAllUsers);
router.patch('/me', auth, updateProfile);

module.exports = router;
