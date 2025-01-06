const PptxGenJS = require("pptxgenjs");
const { createCanvas, loadImage } = require("canvas");
const mathjax = require("mathjax-node");

mathjax.start();

// Constants
const MAX_IMAGE_WIDTH = 1.2;
const MAX_IMAGE_HEIGHT = 0.4;
const IMAGE_TEXT_SPACING = 0.1;
const IMAGE_ADJUSTMENT_Y = -0.2;
const LEFT_MARGIN = 1;
const SLIDE_WIDTH = 10;
const FONT_SIZE = 16;

// Helper function: Convert LaTeX to SVG
async function latexToSvg(latex) {
  return new Promise((resolve, reject) => {
    mathjax.typeset(
      { math: latex, format: "TeX", svg: true },
      (data) => {
        if (data.errors) {
          reject(`Error converting LaTeX to SVG: ${data.errors}`);
        } else {
          resolve(data.svg);
        }
      }
    );
  });
}

// Helper function: Convert SVG to PNG buffer
async function latexToPngBuffer(latex) {
  const svg = await latexToSvg(latex);
  const canvas = createCanvas(500, 200);
  const context = canvas.getContext("2d");
  const buffer = Buffer.from(svg, "utf-8");
  const image = await loadImage(buffer);
  context.drawImage(image, 0, 0, canvas.width, canvas.height);
  return canvas.toBuffer("image/png");
}

// Process slide content and replace equations with images
async function processSlideContent(content) {
  const formulaRegex = /<span class="ql-custom-formula" data-value="(.*?)"><\/span>/g;
  const parts = [];
  let lastIndex = 0;
  let match;

  while ((match = formulaRegex.exec(content)) !== null) {
    if (match.index > lastIndex) {
      const textPart = content.slice(lastIndex, match.index).trim();
      if (textPart) parts.push({ type: "text", value: textPart }); // Add non-empty text
    }
    const latex = match[1];

    // Generate PNG image for the LaTeX equation
    try {
      const imageBuffer = await latexToPngByteStream(latex);
      parts.push({
        type: "image",
        value: imageBuffer,
        width: MAX_IMAGE_WIDTH,
        height: MAX_IMAGE_HEIGHT,
      });
    } catch (err) {
      console.error("Error rendering equation:", latex, err);
    }

    lastIndex = formulaRegex.lastIndex;
  }

  // Add any trailing text content after the last equation
  if (lastIndex < content.length) {
    const trailingText = content.slice(lastIndex).trim();
    if (trailingText) parts.push({ type: "text", value: trailingText });
  }

  return parts;
}


// Generate PowerPoint presentation
async function generatePpt(req, res) {

  const { slides, template } = {
    "slides": [
        {
            "title": "Math Equation Demo",
            "sub_title": "An introduction to Einstein famous formula",
            "content": "<p>Einstein said: <span class=\"ql-custom-formula\" data-value=\"E = mc^2\"><span contenteditable=\"false\"><span class=\"katex\"><span class=\"katex-mathml\"><math><semantics><mrow><mi>E</mi><mo>=</mo><mi>m</mi><msup><mi>c</mi><mn>2</mn></msup></mrow><annotation encoding=\"application/x-tex\">E = mc^2</annotation></semantics></math></span><span class=\"katex-html\" aria-hidden=\"true\"><span class=\"strut\" style=\"height: 0.814108em;\"></span><span class=\"strut bottom\" style=\"height: 0.814108em; vertical-align: 0em;\"></span><span class=\"base\"><span class=\"mord mathit\" style=\"margin-right: 0.05764em;\">E</span><span class=\"mrel\">=</span><span class=\"mord mathit\">m</span><span class=\"mord\"><span class=\"mord mathit\">c</span><span class=\"msupsub\"><span class=\"vlist-t\"><span class=\"vlist-r\"><span class=\"vlist\" style=\"height: 0.814108em;\"><span class=\"\" style=\"top: -3.063em; margin-right: 0.05em;\"><span class=\"pstrut\" style=\"height: 2.7em;\"></span><span class=\"sizing reset-size6 size3 mtight\"><span class=\"mord mathrm mtight\">2</span></span></span></span></span></span></span></span></span></span></span></span></span> is the famous formula that revolutionized physics. It introduced the concept of mass-energy equivalence, which is a fundamental principle of the universe.</p>"
        },
        {
            "title": "Quadratic Formula",
            "sub_title": "Solving quadratic equations in algebra",
            "content": "<p>The formula is <span class=\"ql-custom-formula\" data-value=\"x^2 - 5x + 6 = 0 = 0\"><span contenteditable=\"false\"><span class=\"katex\"><span class=\"katex-mathml\"><math><semantics><mrow><msup><mi>x</mi><mn>2</mn></msup><mo>−</mo><mn>5</mn><mi>x</mi><mo>+</mo><mn>6</mn><mo>=</mo><mn>0</mn><mo>=</mo><mn>0</mn></mrow><annotation encoding=\"application/x-tex\">x^2 - 5x + 6 = 0 = 0</annotation></semantics></math></span><span class=\"katex-html\" aria-hidden=\"true\"><span class=\"strut\" style=\"height: 0.814108em;\"></span><span class=\"strut bottom\" style=\"height: 0.897438em; vertical-align: -0.08333em;\"></span><span class=\"base\"><span class=\"mord\"><span class=\"mord mathit\">x</span><span class=\"msupsub\"><span class=\"vlist-t\"><span class=\"vlist-r\"><span class=\"vlist\" style=\"height: 0.814108em;\"><span class=\"\" style=\"top: -3.063em; margin-right: 0.05em;\"><span class=\"pstrut\" style=\"height: 2.7em;\"></span><span class=\"sizing reset-size6 size3 mtight\"><span class=\"mord mathrm mtight\">2</span></span></span></span></span></span></span></span><span class=\"mbin\">−</span><span class=\"mord mathrm\">5</span><span class=\"mord mathit\">x</span><span class=\"mbin\">+</span><span class=\"mord mathrm\">6</span><span class=\"mrel\">=</span><span class=\"mord mathrm\">0</span><span class=\"mrel\">=</span><span class=\"mord mathrm\">0</span></span></span></span></span></span>. This formula is essential for solving quadratic equations in algebra.</p>"
        },
        {
            "title": "Empty Subtitle Example",
            "sub_title": "",
            "content": "<p>An example slide with an empty subtitle. It should be skipped when rendering.</p>"
        }
    ],
    "template": {
        "image_path": "https://prepaze-lms-store-staging.s3-us-west-2.amazonaws.com/public/1/1/iwgBWwoZkmLzXTpqhvXTCe/1-240RQ02Z3564.jpeg",
        "font_family": "Poppins"
    }
}

  const pptx = new PptxGenJS();

  pptx.title = "Generated Presentation";

  for (const slideData of slides) {
    const slide = pptx.addSlide();

    // Apply background and titles
    slide.background = { path: template.image_path };
    slide.addText(slideData.title, { x: 0.5, y: 0.5, fontSize: 24, bold: true, fontFace: template.font_family });
    if (slideData.sub_title) {
      slide.addText(slideData.sub_title, { x: 0.5, y: 1, fontSize: 18, italic: true, fontFace: template.font_family });
    }

    const processedContent = await processSlideContent(slideData.content);
    let xPosition = LEFT_MARGIN;
    let yPosition = slideData.sub_title ? 1.8 : 1.5;

    for (const part of processedContent) {
      if (part.type === "text") {
        slide.addText(part.value, { x: xPosition, y: yPosition, fontSize: FONT_SIZE, fontFace: template.font_family });
        xPosition += part.value.length * 0.1; // Adjust position based on text length
      } else if (part.type === "image") {
        slide.addImage({
          data: `data:image/png;base64,${part.value.toString("base64")}`,
          x: xPosition + 0.4, // Offset to avoid overlap
          y: yPosition + IMAGE_ADJUSTMENT_Y,
          w: part.width,
          h: part.height,
        });
        xPosition += part.width + IMAGE_TEXT_SPACING; // Adjust position for image
      }

      // Wrap text if content exceeds slide width
      if (xPosition > SLIDE_WIDTH - LEFT_MARGIN) {
        xPosition = LEFT_MARGIN;
        yPosition += 0.5; // Move down for next line
      }
    }
  }

  const buffer = await pptx.write("base64");
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
  res.setHeader("Content-Disposition", "attachment; filename=GeneratedPresentation.pptx");
  res.send(Buffer.from(buffer, "base64"));
}


module.exports = { generatePpt };
