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



// let slides = [
//   {
//     "title": "Physics Equation",
//     "content": "Here is the equation for mass-energy equivalence: $$E = mc^2$$"
//   },
//   {
//     "title": "Quadratic Formula",
//     "content": "The quadratic formula is $$x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}$$."
//   },
//   {
//     "title": "Example Text and Math",
//     "content": "A mix of text and math: Distance traveled $$d = vt + \\frac{1}{2}at^2$$."
//   }
// ]

