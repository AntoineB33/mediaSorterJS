// server.js
require('dotenv').config(); // if using dotenv for environment variables
const express = require('express');
const app = express();
const port = process.env.PORT || 3000;

// Global middlewares
app.use(express.json());

// Load routes
const routes = require('./src/routes');
app.use('/', routes);

// Start server
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
