const express = require('express');
const path = require('path');
const pptRoutes = require('./routes/routes'); // Import routes

const app = express();

// Serve static files (e.g., images) from the public directory
app.use('/public', express.static(path.join(__dirname, 'public')));

// Use the PPT routes
app.use('/ppt', pptRoutes);

const port = 3000;
app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});
