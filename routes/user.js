const router = require('express').Router();
const {
  createUsers,
  login,
  getUsersСurrent,
  getAllUsers,
  updateProfile,
  newUserRole,
  getAllUsersBranch,
  delUserRole,
} = require('../controllers/user');
const {
  validationCreateUser,
  validationLogin,
} = require('../middlewares/validation');
const auth = require('../middlewares/auth');
// const role = require('../middlewares/role');

router.post('/signup', auth, validationCreateUser, createUsers);
router.post('/signin', validationLogin, login);
router.get('/me', auth, getUsersСurrent);
router.get('/all', auth, getAllUsers);
router.get('/all/branch', auth, getAllUsersBranch);
router.patch('/me', auth, updateProfile);
router.patch('/role', auth, newUserRole);
router.delete('/role', auth, delUserRole);

module.exports = router;
