const PptxGenJS = require("pptxgenjs");
const katex = require("katex");
const { createCanvas, loadImage } = require("canvas");
const fs = require("fs").promises;
const path = require("path");
const { JSDOM } = require("jsdom");
const mathjax = require("mathjax-node");

mathjax.start();

// Constants for slide layout
const MAX_IMAGE_WIDTH = 1.2;
const MAX_IMAGE_HEIGHT = 0.4;
const IMAGE_TEXT_SPACING = 0.1;
const IMAGE_ADJUSTMENT_Y = -0.2;
const LEFT_MARGIN = 1;
const SLIDE_WIDTH = 10;
const FONT_SIZE = 16;

// Function to clean up content data
const processContent = (content) => {
  const dom = new JSDOM(content);
  const document = dom.window.document;

  const formulas = document.querySelectorAll(".ql-custom-formula");
  formulas.forEach((formula) => {
    const dataValue = formula.getAttribute("data-value") || "";
    const replacementNode = document.createTextNode(`$${dataValue}$`);
    formula.parentNode.replaceChild(replacementNode, formula);
  });

  return document.body.innerHTML;
};

async function processSlideContent(content) {
  // Regular expression to match LaTeX inside $...$ or $$...$$ for math equations
  const formulaRegex = /\$([^$]+)\$/g;
  const parts = [];
  let lastIndex = 0;
  let match;

  while ((match = formulaRegex.exec(content)) !== null) {
    if (match.index > lastIndex) {
      // Push the text part before the equation
      parts.push({
        type: "text",
        value: content.slice(lastIndex, match.index).trim(),
      });
    }

    const latex = match[1]; // Extract LaTeX code
    const imageBuffer = await latexToPngByteStream(latex); // You will need to implement this function based on your LaTeX-to-image rendering method

    // Determine size based on LaTeX length
    const sizeMultiplier = latex.length > 15 ? 1 : 0.7; // Adjust size for longer or shorter equations
    parts.push({
      type: "image",
      value: imageBuffer,
      width: MAX_IMAGE_WIDTH * sizeMultiplier,
      height: MAX_IMAGE_HEIGHT * sizeMultiplier,
    });

    lastIndex = formulaRegex.lastIndex; // Update the last index to continue processing text after the formula
  }

  if (lastIndex < content.length) {
    // Add any remaining text after the last formula
    parts.push({ type: "text", value: content.slice(lastIndex).trim() });
  }

  return parts;
}

// Helper function to convert LaTeX to SVG
function latexToSvg(latex) {
  return new Promise((resolve, reject) => {
    mathjax.typeset({ math: latex, format: "TeX", svg: true }, (data) => {
      if (data.errors) {
        reject(`Error in LaTeX conversion: ${data.errors}`);
      } else {
        resolve(data.svg);
      }
    });
  });
}

// Helper function to convert LaTeX SVG to PNG byte stream
async function latexToPngByteStream(latex) {
  const svg = await latexToSvg(latex);

  const canvas = createCanvas(500, 200);
  const context = canvas.getContext("2d");
  const buffer = Buffer.from(svg, "utf-8");
  const image = await loadImage(buffer);

  context.drawImage(image, 0, 0, canvas.width, canvas.height);
  return canvas.toBuffer("image/png");
}

// Function to generate equation image and dimensions
const generateEquationImage = async (equation) => {
  const { createCanvas } = require("canvas");
  const jsdom = require("jsdom");
  const { JSDOM } = jsdom;

  try {
    // Render equation with KaTeX as an HTML string
    const katexHtml = katex.renderToString(equation, { throwOnError: false });

    // Use JSDOM to calculate the rendered size
    const dom = new JSDOM("<!DOCTYPE html>");
    const document = dom.window.document;
    const container = document.createElement("div");
    container.style.fontSize = "20px"; // Set base font size
    container.style.display = "inline-block"; // Ensure inline dimensions
    container.innerHTML = katexHtml;

    document.body.appendChild(container);
    const boundingBox = container.getBoundingClientRect();

    const width = Math.ceil(boundingBox.width) + 20; // Add some padding
    const height = Math.ceil(boundingBox.height) + 20;

    // Create a properly sized canvas
    const canvas = createCanvas(width, height);
    const ctx = canvas.getContext("2d");

    // Set background to white and clear canvas
    ctx.fillStyle = "white";
    ctx.fillRect(0, 0, width, height);

    // Draw the rendered HTML onto the canvas
    ctx.fillStyle = "black";
    ctx.font = "20px serif"; // Match the font used for KaTeX rendering
    ctx.textAlign = "center";

    // Draw rendered math on canvas
    ctx.fillText(katexHtml, width / 2, height / 2); // Center the math

    // Convert canvas to image buffer
    const buffer = canvas.toBuffer("image/png");

    return {
      base64: `data:image/png;base64,${buffer.toString("base64")}`,
      width,
      height,
    };
  } catch (error) {
    console.error("Error rendering equation:", error);
    throw new Error("Failed to render math equation.");
  }
};

// Function to generate the PPT
const generatePpt = async (req, res) => {
  try {
    let { slides, template } = {
      slides: [
        {
          title: "Math Equation Demo",
          sub_title: "An introduction to Einstein famous formula",
          content:
            '<p>Einstein said: <span class="ql-custom-formula" data-value="E = mc^2"><span contenteditable="false"><span class="katex"><span class="katex-mathml"><math><semantics><mrow><mi>E</mi><mo>=</mo><mi>m</mi><msup><mi>c</mi><mn>2</mn></msup></mrow><annotation encoding="application/x-tex">E = mc^2</annotation></semantics></math></span><span class="katex-html" aria-hidden="true"><span class="strut" style="height: 0.814108em;"></span><span class="strut bottom" style="height: 0.814108em; vertical-align: 0em;"></span><span class="base"><span class="mord mathit" style="margin-right: 0.05764em;">E</span><span class="mrel">=</span><span class="mord mathit">m</span><span class="mord"><span class="mord mathit">c</span><span class="msupsub"><span class="vlist-t"><span class="vlist-r"><span class="vlist" style="height: 0.814108em;"><span class="" style="top: -3.063em; margin-right: 0.05em;"><span class="pstrut" style="height: 2.7em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mathrm mtight">2</span></span></span></span></span></span></span></span></span></span></span> is the famous formula that revolutionized physics. It introduced the concept of mass-energy equivalence, which is a fundamental principle of the universe.</p>',
        },
        {
          title: "Quadratic Formula",
          sub_title: "Solving quadratic equations in algebra",
          content:
            '<p>Which of the following are solutions to the quadratic equation <span class="ql-custom-formula" data-value="x^2 - 3x + 2 = 0">﻿<span contenteditable="false"><span class="katex"><span class="katex-mathml"><math><semantics><mrow><msup><mi>x</mi><mn>2</mn></msup><mo>−</mo><mn>3</mn><mi>x</mi><mo>+</mo><mn>2</mn><mo>=</mo><mn>0</mn></mrow><annotation encoding="application/x-tex">x^2 - 3x + 2 = 0</annotation></semantics></math></span><span class="katex-html" aria-hidden="true"><span class="strut" style="height: 0.814108em;"></span><span class="strut bottom" style="height: 0.897438em; vertical-align: -0.08333em;"></span><span class="base"><span class="mord"><span class="mord mathit">x</span><span class="msupsub"><span class="vlist-t"><span class="vlist-r"><span class="vlist" style="height: 0.814108em;"><span class="" style="top: -3.063em; margin-right: 0.05em;"><span class="pstrut" style="height: 2.7em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mathrm mtight">2</span></span></span></span></span></span></span></span></span><span class="mbin">−</span><span class="mord mathrm">3</span><span class="mord mathit">x</span><span class="mbin">+</span><span class="mord mathrm">2</span><span class="mrel">=</span><span class="mord mathrm">0</span></span></span></span></span>﻿</span>?</p>',
        },
        {
          title: "fractions",
          content:
            '<p>Consider a cake cut into 16 equal pieces. Emily ate <span class="ql-custom-formula" data-value="\\frac{3}{16}">﻿<span contenteditable="false"><span class="katex"><span class="katex-mathml"><math><semantics><mrow><mfrac><mrow><mn>3</mn></mrow><mrow><mn>1</mn><mn>6</mn></mrow></mfrac></mrow><annotation encoding="application/x-tex">\\frac{3}{16}</annotation></semantics></math></span><span class="katex-html" aria-hidden="true"><span class="strut" style="height: 0.845108em;"></span><span class="strut bottom" style="height: 1.19011em; vertical-align: -0.345em;"></span><span class="base"><span class="mord"><span class="mopen nulldelimiter"></span><span class="mfrac"><span class="vlist-t vlist-t2"><span class="vlist-r"><span class="vlist" style="height: 0.845108em;"><span class="" style="top: -2.655em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">1</span><span class="mord mathrm mtight">6</span></span></span></span><span class="" style="top: -3.23em;"><span class="pstrut" style="height: 3em;"></span><span class="frac-line" style="border-bottom-width: 0.04em;"></span></span><span class="" style="top: -3.394em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">3</span></span></span></span></span><span class="vlist-s">​</span></span><span class="vlist-r"><span class="vlist" style="height: 0.345em;"></span></span></span></span><span class="mclose nulldelimiter"></span></span></span></span></span></span>﻿</span> of the cake, John ate <span class="ql-custom-formula" data-value="\\frac{5}{16}">﻿<span contenteditable="false"><span class="katex"><span class="katex-mathml"><math><semantics><mrow><mfrac><mrow><mn>5</mn></mrow><mrow><mn>1</mn><mn>6</mn></mrow></mfrac></mrow><annotation encoding="application/x-tex">\\frac{5}{16}</annotation></semantics></math></span><span class="katex-html" aria-hidden="true"><span class="strut" style="height: 0.845108em;"></span><span class="strut bottom" style="height: 1.19011em; vertical-align: -0.345em;"></span><span class="base"><span class="mord"><span class="mopen nulldelimiter"></span><span class="mfrac"><span class="vlist-t vlist-t2"><span class="vlist-r"><span class="vlist" style="height: 0.845108em;"><span class="" style="top: -2.655em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">1</span><span class="mord mathrm mtight">6</span></span></span></span><span class="" style="top: -3.23em;"><span class="pstrut" style="height: 3em;"></span><span class="frac-line" style="border-bottom-width: 0.04em;"></span></span><span class="" style="top: -3.394em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">5</span></span></span></span></span><span class="vlist-s">​</span></span><span class="vlist-r"><span class="vlist" style="height: 0.345em;"></span></span></span></span><span class="mclose nulldelimiter"></span></span></span></span></span></span>﻿</span> of the cake. How much of the cake remains?</p>',
        },
      ],
      template: {
        name: "calculation_template",
        theme: "unit_one",
      },
    };

    // Data cleaning: process each slide's content
    slides.forEach((slide) => {
      slide.content = processContent(slide.content);
    });

    const pptx = new PptxGenJS();

    for (const slideData of slides) {
      const slide = pptx.addSlide();

      // Set background image from template
      slide.background = { path: template.image_path };

      // Add title
      slide.addText(slideData.title, {
        x: 0.5,
        y: 0.5,
        fontSize: 24,
        bold: true,
        fontFace: template.font_family,
      });

      // Add subtitle
      if (slideData.sub_title) {
        slide.addText(slideData.sub_title, {
          x: 0.5,
          y: 1,
          fontSize: 18,
          fontFace: template.font_family,
          italic: true,
        });
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
            x: xPosition + 0.4, // Add extra offset to the right
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
  
    const buffer = await pptx.write("base64");
    const responseBuffer = Buffer.from(buffer, "base64");
  
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", "attachment; filename=GeneratedPresentation.pptx");
    res.send(responseBuffer);
  } catch (err) {
    console.error("Error generating PPT:", err);
    res.status(500).send("Error generating PPT.");
  }
};

module.exports = { generatePpt };
