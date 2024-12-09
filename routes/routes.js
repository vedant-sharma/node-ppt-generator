const express = require('express');
const { generatePpt } = require('../handlers/pptHandler'); // Import the handler

const router = express.Router();

// Route for generating PPT
router.get('/', generatePpt);

module.exports = router;
