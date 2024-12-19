const fs = require('fs');
const path = require('path');
const PptxGenJS = require('pptxgenjs');
const { createCanvas, loadImage } = require('canvas');
const mathjax = require('mathjax-node');

// Set MathJax configuration
mathjax.start();

// Convert LaTeX to SVG
function latexToSvg(latex) {
  return new Promise((resolve, reject) => {
    mathjax.typeset(
      {
        math: latex,
        format: 'TeX', // Input LaTeX
        svg: true,     // Output SVG
      },
      (data) => {
        if (data.errors) {
          reject(`Error in LaTeX conversion: ${data.errors}`);
        } else {
          resolve(data.svg);
        }
      }
    );
  });
}

// Convert LaTeX to PNG and return the path
async function latexToPng(latex, outputDir) {
  try {
    const svg = await latexToSvg(latex);

    // Canvas setup
    const canvas = createCanvas(300, 50); // Adjust size as needed
    const context = canvas.getContext('2d');
    const buffer = Buffer.from(svg, 'utf-8');
    const image = await loadImage(buffer);

    context.drawImage(image, 0, 0);

    const fileName = `equation_${Date.now()}.png`;
    const filePath = path.join(outputDir, fileName);
    const imageBuffer = canvas.toBuffer('image/png');
    fs.writeFileSync(filePath, imageBuffer);

    return filePath;
  } catch (err) {
    console.error('Error converting LaTeX to PNG:', err);
    throw err;
  }
}

// Process slide content
async function processSlideContent(content, outputDir) {
  const equationRegex = /\$(.*?)\$/g; // Matches LaTeX equations within '$...$'
  let match;
  const contentParts = [];
  let lastIndex = 0;

  while ((match = equationRegex.exec(content)) !== null) {
    const latex = match[1]; // Extract LaTeX
    if (match.index > lastIndex) {
      contentParts.push({ type: 'text', value: content.slice(lastIndex, match.index) });
    }

    const imagePath = await latexToPng(latex, outputDir);
    contentParts.push({ type: 'image', value: imagePath });

    lastIndex = equationRegex.lastIndex;
  }

  if (lastIndex < content.length) {
    contentParts.push({ type: 'text', value: content.slice(lastIndex) });
  }

  return contentParts;
}

// Generate presentation with properly aligned content
async function generatePpt(req, res) {
  const slidesData = [
    { title: 'Slide 1', content: 'This is a math equation: $E=mc^2$ inline example.' },
    { title: 'Slide 2', content: 'Solve: $x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}$ and see the inline placement.' },
  ];

  const outputDir = path.join(__dirname, 'output');
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

  const pptx = new PptxGenJS();

  for (const slideData of slidesData) {
    const slide = pptx.addSlide();
    slide.addText(slideData.title, { x: 1, y: 0.5, fontSize: 24, bold: true });

    const processedContent = await processSlideContent(slideData.content, outputDir);

    let xPosition = 0.5;
    const yPosition = 1.5; // Fixed y position for inline layout

    for (const part of processedContent) {
      if (part.type === 'text') {
        const textWidth = 1.5 * (part.value.length / 10); // Estimate width based on text length
        slide.addText(part.value, { x: xPosition, y: yPosition, fontSize: 16, color: '363636' });
        xPosition += textWidth; // Increment x position
      } else if (part.type === 'image') {
        const imageWidth = 1.5; // Fixed width for image
        slide.addImage({ path: part.value, x: xPosition, y: yPosition, w: imageWidth, h: 0.5 });
        xPosition += imageWidth + 0.1; // Increment x position, adding small margin
      }
    }
  }

  const outputFilePath = path.join(outputDir, 'Presentation.pptx');
  await pptx.writeFile({ fileName: outputFilePath });

  res.status(200).send(`PowerPoint saved to: ${outputFilePath}`);
}

module.exports = { generatePpt };
