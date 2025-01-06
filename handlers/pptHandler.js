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
const MAX_IMAGE_HEIGHT = 0.6;
const IMAGE_TEXT_SPACING = 0.1;
const IMAGE_ADJUSTMENT_Y = -0.2;
const LEFT_MARGIN = 1;
const SLIDE_WIDTH = 10;
const FONT_SIZE = 16;

// Function to clean up content data
const processContent = (content) => {

  content = content.replace(/(<p[^>]+?>|<p>|<\/p>)/img, "");

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

    console.log(latex)
    console.log(latex.length)

    // Determine size based on LaTeX length
    const sizeMultiplier = latex.length > 25 ? 1 : 0.5; // Adjust size for longer or shorter equations
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

  console.log(parts)

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

// Function to generate the PPT
const generatePpt = async (req, res) => {
  try {
    let { slides, template } = {
      "slides": [
          {
              "title": "Math Equation Demo",
              "sub_title": "An introduction to Einstein famous formula",
              "content": "<p>Einstein said: <span class=\"ql-custom-formula\" data-value=\"E = mc^2\"><span contenteditable=\"false\"><span class=\"katex\"><span class=\"katex-mathml\"><math><semantics><mrow><mi>E</mi><mo>=</mo><mi>m</mi><msup><mi>c</mi><mn>2</mn></msup></mrow><annotation encoding=\"application/x-tex\">E = mc^2</annotation></semantics></math></span><span class=\"katex-html\" aria-hidden=\"true\"><span class=\"strut\" style=\"height: 0.814108em;\"></span><span class=\"strut bottom\" style=\"height: 0.814108em; vertical-align: 0em;\"></span><span class=\"base\"><span class=\"mord mathit\" style=\"margin-right: 0.05764em;\">E</span><span class=\"mrel\">=</span><span class=\"mord mathit\">m</span><span class=\"mord\"><span class=\"mord mathit\">c</span><span class=\"msupsub\"><span class=\"vlist-t\"><span class=\"vlist-r\"><span class=\"vlist\" style=\"height: 0.814108em;\"><span class=\"\" style=\"top: -3.063em; margin-right: 0.05em;\"><span class=\"pstrut\" style=\"height: 2.7em;\"></span><span class=\"sizing reset-size6 size3 mtight\"><span class=\"mord mathrm mtight\">2</span></span></span></span></span></span></span></span></span></span></span></span></span> is the famous formula that revolutionized physics. It introduced the concept of mass-energy equivalence, which is a fundamental principle of the universe.</p>"
          },
          {
              "title": "Quadratic Formula",
              "sub_title": "Solving quadratic equations in algebra",
              "content": "<p>Which of the following are solutions to the quadratic equation <span class=\"ql-custom-formula\" data-value=\"x^2 - 3x + 2 = 0\">﻿<span contenteditable=\"false\"><span class=\"katex\"><span class=\"katex-mathml\"><math><semantics><mrow><msup><mi>x</mi><mn>2</mn></msup><mo>−</mo><mn>3</mn><mi>x</mi><mo>+</mo><mn>2</mn><mo>=</mo><mn>0</mn></mrow><annotation encoding=\"application/x-tex\">x^2 - 3x + 2 = 0</annotation></semantics></math></span><span class=\"katex-html\" aria-hidden=\"true\"><span class=\"strut\" style=\"height: 0.814108em;\"></span><span class=\"strut bottom\" style=\"height: 0.897438em; vertical-align: -0.08333em;\"></span><span class=\"base\"><span class=\"mord\"><span class=\"mord mathit\">x</span><span class=\"msupsub\"><span class=\"vlist-t\"><span class=\"vlist-r\"><span class=\"vlist\" style=\"height: 0.814108em;\"><span class=\"\" style=\"top: -3.063em; margin-right: 0.05em;\"><span class=\"pstrut\" style=\"height: 2.7em;\"></span><span class=\"sizing reset-size6 size3 mtight\"><span class=\"mord mathrm mtight\">2</span></span></span></span></span></span></span></span><span class=\"mbin\">−</span><span class=\"mord mathrm\">3</span><span class=\"mord mathit\">x</span><span class=\"mbin\">+</span><span class=\"mord mathrm\">2</span><span class=\"mrel\">=</span><span class=\"mord mathrm\">0</span></span></span></span></span>﻿</span>?</p>"
          },
          {
              "title": "Empty Subtitle Example",
              "sub_title": "",
              "content": "<p>An example slide with an empty subtitle. It should be skipped when rendering. An example slide with an empty subtitle. It should be skipped when rendering. An example slide with an empty subtitle. It should be skipped when rendering. An example slide with an empty subtitle. It should be skipped when rendering.</p>"
          }
      ],
      "template": {
          "image_path": "https://prepaze-lms-store-staging.s3-us-west-2.amazonaws.com/public/1/1/iwgBWwoZkmLzXTpqhvXTCe/1-240RQ02Z3564.jpeg",
          "font_family": "Poppins"
      }
  }

    // Data cleaning: process each slide's content
    slides.forEach((slide) => {
      slide.content = processContent(slide.content);
    });

    console.log(slides)

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
  
      // for (const part of processedContent) {
      //   if (part.type === 'text') {
      //     slide.addText(part.value, { x: xPosition, y: yPosition, fontSize: FONT_SIZE, fontFace: template.font_family });
      //     xPosition += part.value.length * 0.1; // Estimate text width
      //   } else if (part.type === 'image') {
      //     slide.addImage({
      //       data: `data:image/png;base64,${part.value.toString('base64')}`,
      //       x: xPosition + 0.1, // Add extra offset to the right
      //       y: yPosition + IMAGE_ADJUSTMENT_Y,
      //       w: part.width,
      //       h: part.height,
      //     });
      //     xPosition += part.width + IMAGE_TEXT_SPACING;
      //   }
  
      //   if (xPosition > SLIDE_WIDTH - LEFT_MARGIN) {
      //     xPosition = LEFT_MARGIN;
      //     yPosition += 0.5;
      //   }
      // }

      for (const part of processedContent) {
        if (part.type === "text") {
          slide.addText(part.value, {
            x: xPosition,
            y: yPosition,
            w: SLIDE_WIDTH - LEFT_MARGIN * 2, // Define text box width to prevent overflow
            fontSize: FONT_SIZE,
            fontFace: template.font_family,
            color: "000000", // Optional: Set text color
            align: "left", // Optional: Text alignment
          });
          yPosition += 0.5; // Move to the next line after adding text
        } else if (part.type === "image") {
          slide.addImage({
            data: `data:image/png;base64,${part.value.toString("base64")}`,
            x: xPosition,
            y: yPosition + IMAGE_ADJUSTMENT_Y,
            w: part.width,
            h: part.height,
          });
          xPosition += part.width + IMAGE_TEXT_SPACING;
      
          // Check if image exceeds slide width
          if (xPosition > SLIDE_WIDTH - LEFT_MARGIN) {
            xPosition = LEFT_MARGIN;
            yPosition += part.height + IMAGE_TEXT_SPACING;
          }
        }
      
        // Reset xPosition and increment yPosition if overflowing the slide width
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
