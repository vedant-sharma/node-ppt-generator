const PptxGenJS = require('pptxgenjs');
const { createCanvas, loadImage } = require('canvas');
const mathjax = require('mathjax-node');

mathjax.start();

const MAX_IMAGE_WIDTH = 1.2; // Maximum image width (in inches)
const MAX_IMAGE_HEIGHT = 0.4; // Maximum image height (in inches)
const SMALL_IMAGE_WIDTH = 0.8; // Smaller image width (in inches)
const SMALL_IMAGE_HEIGHT = 0.3; // Smaller image height (in inches)
const IMAGE_TEXT_SPACING = 0.1; // Space between text and image (in inches)
const IMAGE_ADJUSTMENT_Y = -0.2; // Adjust image upwards slightly (in inches)
const LEFT_MARGIN = 1; // Left margin for content (in inches)
const SLIDE_WIDTH = 10; // Total slide width (in inches)
const FONT_SIZE = 16; // Font size for text

// Helper function to calculate the width of text in inches
function calculateTextWidth(text, fontSize = FONT_SIZE) {
  const charWidth = 0.12; // Approx. width of a character for the given font size
  return text.length * charWidth;
}

// Helper function to wrap text to fit within available width
function wrapText(text, availableWidth, fontSize = FONT_SIZE) {
  const words = text.split(' ');
  let wrappedLines = [];
  let currentLine = '';

  for (let word of words) {
    const lineWidth = calculateTextWidth(currentLine + word, fontSize);

    if (lineWidth <= availableWidth) {
      currentLine += (currentLine === '' ? '' : ' ') + word;
    } else {
      wrappedLines.push(currentLine);
      currentLine = word;
    }
  }

  if (currentLine !== '') {
    wrappedLines.push(currentLine);
  }

  return wrappedLines;
}

// Function to convert LaTeX to SVG
function latexToSvg(latex) {
  return new Promise((resolve, reject) => {
    mathjax.typeset(
      {
        math: latex,
        format: 'TeX',
        svg: true,
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

// Function to convert LaTeX SVG directly to a PNG byte stream
async function latexToPngByteStream(latex) {
  try {
    const svg = await latexToSvg(latex);

    const canvas = createCanvas(500, 200); // Larger dimensions for better resolution
    const context = canvas.getContext('2d');
    const buffer = Buffer.from(svg, 'utf-8');
    const image = await loadImage(buffer);

    // Scale the SVG to fit the canvas
    context.drawImage(image, 0, 0, canvas.width, canvas.height);

    // Convert the canvas to a PNG buffer (byte stream)
    const pngBuffer = canvas.toBuffer('image/png');
    return pngBuffer;
  } catch (err) {
    console.error('Error converting LaTeX to PNG:', err);
    throw err;
  }
}

// Process slide content: text and equations with byte stream images
async function processSlideContent(content) {
  const equationRegex = /\$(.*?)\$/g;
  const parts = [];
  let lastIndex = 0;

  let match;
  while ((match = equationRegex.exec(content)) !== null) {
    if (match.index > lastIndex) {
      parts.push({ type: 'text', value: content.slice(lastIndex, match.index).trim() });
    }

    const latex = match[1];
    const imageBuffer = await latexToPngByteStream(latex);

    // Determine image size based on LaTeX length
    const imageWidth = latex.length < 15 ? SMALL_IMAGE_WIDTH : MAX_IMAGE_WIDTH;
    const imageHeight = latex.length < 15 ? SMALL_IMAGE_HEIGHT : MAX_IMAGE_HEIGHT;

    parts.push({ type: 'image', value: imageBuffer, width: imageWidth, height: imageHeight });

    lastIndex = equationRegex.lastIndex;
  }

  if (lastIndex < content.length) {
    parts.push({ type: 'text', value: content.slice(lastIndex).trim() });
  }

  return parts;
}

// Generate PowerPoint presentation
async function generatePpt(req, res) {
  const slidesData = [
    {
        "title": "Introduction to Animal Kingdom",
        "sub_title": "Overview of Biological Classification",
        "content": "Explore the diversity of life in the Animal Kingdom, which includes multicellular organisms that consume organic material, breathe oxygen, and can move. Animals are classified based on characteristics such as body structure, reproduction method, and genetic similarities."
    },
    {
        "title": "Major Animal Phyla",
        "sub_title": "Key Groups Within the Animal Kingdom",
        "content": "Key phyla include: \n1. Chordata (vertebrates and some invertebrates) \n 2. Arthropoda (insects, arachnids, crustaceans) \n 3. Mollusca (snails, bivalves, cephalopods) \n 4. Annelida (segmented worms) \n 5. Cnidaria (jellyfish, corals) \n Each phylum has unique features that adapt members to diverse environments."
    },
    {
        "title": "Reproduction and Lifecycle",
        "sub_title": "From Birth to Adulthood",
        "content": "Animals reproduce primarily through sexual reproduction, though some also exhibit asexual reproduction. Lifecycles can vary widely:\n1. Metamorphosis in insects and amphibians.\n2. Direct development in mammals.\nUnderstanding these cycles is crucial for studying animal growth, behavior, and adaptation."
    }
]

  const pptx = new PptxGenJS();

  for (const slideData of slidesData) {
    const slide = pptx.addSlide();

    // Add title
    slide.addText(slideData.title, { x: 0.5, y: 0.5, fontSize: 24, bold: true });

    // Add subtitle if it exists
    if (slideData.subtitle) {
      slide.addText(slideData.subtitle, { x: 0.5, y: 1, fontSize: 18, color: '757575', italic: true });
    }

    // Process content (text and equations)
    const processedContent = await processSlideContent(slideData.content);

    let yPosition = slideData.subtitle ? 1.8 : 1.5; // Adjust Y position based on subtitle
    let xPosition = LEFT_MARGIN; // Starting X position for content

    for (const part of processedContent) {
      if (part.type === 'text') {
        const wrappedLines = wrapText(part.value, SLIDE_WIDTH - xPosition);

        for (let i = 0; i < wrappedLines.length; i++) {
          const line = wrappedLines[i];

          slide.addText(line, { x: xPosition, y: yPosition, fontSize: FONT_SIZE, color: '363636' });

          if (i < wrappedLines.length - 1) {
            yPosition += 0.4; // Move down for the next line
            xPosition = LEFT_MARGIN; // Reset X position for new line
          } else {
            xPosition += calculateTextWidth(line); // Update X position for inline continuation
          }
        }
      } else if (part.type === 'image') {
        // Add the image inline using the byte stream
        slide.addImage({
          data: `data:image/png;base64,${part.value.toString('base64')}`,
          x: xPosition,
          y: yPosition + IMAGE_ADJUSTMENT_Y, // Adjust Y to move the image slightly up
          w: part.width,
          h: part.height,
        });

        xPosition += part.width + IMAGE_TEXT_SPACING; // Update X position for next element
      }

      // Move to the next line if content exceeds the slide width
      if (xPosition >= SLIDE_WIDTH - LEFT_MARGIN) {
        yPosition += 0.4; // Move down
        xPosition = LEFT_MARGIN; // Reset X position
      }
    }
  }

  const outputFilePath = 'GeneratedPresentationWithDynamicImages.pptx';
  await pptx.writeFile({ fileName: outputFilePath });

  console.log('Presentation saved as:', outputFilePath);
  res.status(200).send(`Presentation saved at: ${outputFilePath}`);
}

module.exports = { generatePpt };