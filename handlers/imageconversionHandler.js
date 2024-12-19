const fs = require('fs');
const { createCanvas, loadImage } = require('canvas');  // For canvas operations
const mathjax = require('mathjax-node');  // For LaTeX to SVG conversion

// Function to convert LaTeX to SVG
function latexToSvg(latex) {
    return new Promise((resolve, reject) => {
        // Convert LaTeX to SVG using MathJax
        mathjax.typeset({
            math: latex,
            format: "TeX",  // The input format
            svg: true       // Output format SVG
        }, (data) => {
            if (data.errors) {
                reject(`Error in LaTeX conversion: ${data.errors}`);
            } else {
                resolve(data.svg);
            }
        });
    });
}

// Function to convert LaTeX to an image (PNG)
const generatePpt = async (req, res) => {

  const latex = 'x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}'; // Example LaTeX for a fraction

  console.log("asa")

  // Step 1: Convert LaTeX to SVG
  const svg = await latexToSvg(latex);

  // Step 2: Wrap the SVG string as a data URL to convert to Buffer
  const svgDataUrl = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);

  // Step 3: Convert the SVG Data URL to a buffer
  const buffer = Buffer.from(svg, 'utf-8');

  // Step 4: Create a canvas (you can specify any dimensions for your image)
  const canvas = createCanvas(500, 200);  // Adjust the size as needed
  const context = canvas.getContext('2d');

  // Step 5: Load the SVG buffer as an image
  const image = await loadImage(buffer);  // Load buffer directly instead of using string URL
  context.drawImage(image, 0, 0);

  // Step 6: Save the canvas as a PNG image
  const imageBuffer = canvas.toBuffer('image/png');
  fs.writeFileSync('latex_image.png', imageBuffer);

  console.log("Image saved as 'latex_image.png'");

        res.status(200).send('generated PPT.');
        
}

// Example LaTeX expression

// Convert the LaTeX expression to PNG image
module.exports = { generatePpt };


