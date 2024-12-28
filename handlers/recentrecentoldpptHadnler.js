const PptxGenJS = require('pptxgenjs');
const katex = require('katex');
const { createCanvas } = require('canvas');
const fs = require('fs').promises;
const path = require('path');

const generateEquationImage = async (equation, index) => {
  const canvas = createCanvas(800, 200); // Create a blank canvas
  const ctx = canvas.getContext('2d');

  try {
    // Render equation into HTML using KaTeX
    const renderedEquation = katex.renderToString(equation, { throwOnError: false });

    // Use canvas to render HTML
    const katexHTML = `<html>
      <body style="margin: 0; padding: 0; display: flex; align-items: center; justify-content: center;">
        ${renderedEquation}
      </body>
    </html>`;

    ctx.fillStyle = 'white';
    ctx.fillRect(0, 0, 800, 200);
    ctx.font = '20px serif';
    ctx.fillStyle = 'black';
    ctx.fillText(equation, 50, 100);

    // Save the image
    const buffer = canvas.toBuffer('image/png');
    const imagePath = path.join(__dirname, `equation_${index}.png`);
    await fs.writeFile(imagePath, buffer);

    return imagePath;
  } catch (error) {
    console.error('Error rendering equation:', error);
    throw new Error('Failed to render math equation.');
  }
};

const generatePpt = async (req, res) => {
  try {

    let slides = req.Body 

    if (!Array.isArray(slides)) {
      return res.status(400).send('Invalid input: slides should be an array.');
    }

    const pptx = new PptxGenJS();

    for (let i = 0; i < slides.length; i++) {
      const { title, content } = slides[i];
      const slide = pptx.addSlide();
    
      // Add the title
      slide.addText(title, { x: 1, y: 0.5, fontSize: 24, bold: true, color: '363636' });
    
      // Ensure content is a string, or default to an empty string
      const safeContent = typeof content === 'string' ? content : '';
    
      // Parse content for math equations (surrounded by $$)
      const equationRegex = /\$\$(.*?)\$\$/g;
      let match;
      let currentY = 1.5; // Positioning y-coordinate on slide
    
      let remainingContent = safeContent;
    
      while ((match = equationRegex.exec(remainingContent)) !== null) {
        const textBeforeEquation = remainingContent.substring(0, match.index).trim();
        const mathEquation = match[1].trim();
        remainingContent = remainingContent.substring(match.index + match[0].length).trim();
    
        // Render text before the equation
        if (textBeforeEquation) {
          slide.addText(textBeforeEquation, {
            x: 1,
            y: currentY,
            fontSize: 18,
            color: '404040',
            align: 'left',
          });
          currentY += 0.8; // Adjust Y position
        }
    
        // Generate and insert the equation image
        const imagePath = await generateEquationImage(mathEquation, `${i}_${match.index}`);
        slide.addImage({ path: imagePath, x: 1, y: currentY, w: 5, h: 1 });
        currentY += 1.2;
      }
    
      // Render any remaining text after the last equation
      if (remainingContent) {
        slide.addText(remainingContent, {
          x: 1,
          y: currentY,
          fontSize: 18,
          color: '404040',
          align: 'left',
        });
      }
    }
    

    const pptxData = await pptx.write('arraybuffer');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', 'attachment; filename="MathPresentation.pptx"');
    res.send(Buffer.from(pptxData));
  } catch (error) {
    console.error('Error generating PPT with math equations:', error);
    res.status(500).send('Failed to generate PPT.');
  }
};

module.exports = { generatePpt };

