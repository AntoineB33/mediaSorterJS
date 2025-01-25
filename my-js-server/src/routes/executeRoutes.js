// src/routes/executeRoutes.js
const express = require('express');
const router = express.Router();
const executeController = require('../controllers/executeController');

router.post('/', executeController.executeProgram);

module.exports = router;