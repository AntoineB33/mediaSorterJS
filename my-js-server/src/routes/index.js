// src/routes/index.js
const express = require('express');
const router = express.Router();

// Import all of your individual route files
const executeRoutes = require('./executeRoutes');
const healthRoutes = require('./healthRoutes');

// Attach route files
router.use('/execute', executeRoutes);
router.use('/health', healthRoutes);

module.exports = router;
