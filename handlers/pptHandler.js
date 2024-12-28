const PptxGenJS = require('pptxgenjs');
const { createCanvas, loadImage } = require('canvas');
const mathjax = require('mathjax-node');

mathjax.start();

const MAX_IMAGE_WIDTH = 1.2;
const MAX_IMAGE_HEIGHT = 0.4;
const IMAGE_TEXT_SPACING = 0.1;
const IMAGE_ADJUSTMENT_Y = -0.2;
const LEFT_MARGIN = 1;
const SLIDE_WIDTH = 10;
const FONT_SIZE = 16;

// Helper function to convert LaTeX to SVG
function latexToSvg(latex) {
  return new Promise((resolve, reject) => {
    mathjax.typeset(
      { math: latex, format: 'TeX', svg: true },
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

// Helper function to convert LaTeX SVG to PNG byte stream
async function latexToPngByteStream(latex) {
  const svg = await latexToSvg(latex);

  const canvas = createCanvas(500, 200);
  const context = canvas.getContext('2d');
  const buffer = Buffer.from(svg, 'utf-8');
  const image = await loadImage(buffer);

  context.drawImage(image, 0, 0, canvas.width, canvas.height);
  return canvas.toBuffer('image/png');
}

// Helper function to process slide content and convert equations
async function processSlideContent(content) {
  const formulaRegex = /<span class="ql-custom-formula" data-value="(.*?)"><\/span>/g;
  const parts = [];
  let lastIndex = 0;
  let match;

  while ((match = formulaRegex.exec(content)) !== null) {
    if (match.index > lastIndex) {
      parts.push({ type: 'text', value: content.slice(lastIndex, match.index).trim() });
    }

    const latex = match[1];
    const imageBuffer = await latexToPngByteStream(latex);
    parts.push({
      type: 'image',
      value: imageBuffer,
      width: MAX_IMAGE_WIDTH,
      height: MAX_IMAGE_HEIGHT,
    });

    lastIndex = formulaRegex.lastIndex;
  }

  if (lastIndex < content.length) {
    parts.push({ type: 'text', value: content.slice(lastIndex).trim() });
  }

  return parts;
}

// Main function to generate PowerPoint
async function generatePpt(req, res) {
  const { slides, template } = {
    slides: [
      {
        title: "Math Equation Demo",
        sub_title: "An introduction to Einstein's famous formula",
        content:
          "Einstein said: <span class=\"ql-custom-formula\" data-value=\"E = mc^2\"></span> revolutionized physics.",
      },
      {
        title: "Quadratic Formula",
        sub_title: "Solving quadratic equations in algebra",
        content:
          "The formula is <span class=\"ql-custom-formula\" data-value=\"x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}\"></span>. It is widely used in algebra.",
      },
    ],
    template: {
      image_path:
        "https://prepaze-lms-store-production.s3-us-west-2.amazonaws.com/public/27/27/LYumW3FCjruHQD9mLbYi5e/slide_background2.jpeg",
      font_family: "Verdana",
    },
  };

  const pptx = new PptxGenJS();

  for (const slideData of slides) {
    const slide = pptx.addSlide();

    // Set background image from template
    slide.background = { path: template.image_path };

    // Add title
    slide.addText(slideData.title, { x: 0.5, y: 0.5, fontSize: 24, bold: true, fontFace: template.font_family });

    // Add subtitle
    if (slideData.sub_title) {
      slide.addText(slideData.sub_title, { x: 0.5, y: 1, fontSize: 18, fontFace: template.font_family, italic: true });
    }

    const processedContent = await processSlideContent(slideData.content);
    let xPosition = LEFT_MARGIN;
    let yPosition = slideData.sub_title ? 1.8 : 1.5;

    for (const part of processedContent) {
      if (part.type === 'text') {
        slide.addText(part.value, { x: xPosition, y: yPosition, fontSize: FONT_SIZE, fontFace: template.font_family });
        xPosition += part.value.length * 0.1; // Estimate text width
      } else if (part.type === 'image') {
        slide.addImage({
          data: `data:image/png;base64,${part.value.toString('base64')}`,
          x: xPosition,
          y: yPosition + IMAGE_ADJUSTMENT_Y,
          w: part.width,
          h: part.height,
        });
        xPosition += part.width + IMAGE_TEXT_SPACING;
      }

      if (xPosition > SLIDE_WIDTH - LEFT_MARGIN) {
        xPosition = LEFT_MARGIN;
        yPosition += 0.5;
      }
    }
  }

  const outputFilePath = 'GeneratedPresentation.pptx';
  await pptx.writeFile({ fileName: outputFilePath });

  res.status(200).send(`Presentation saved as ${outputFilePath}`);
}

module.exports = { generatePpt };
