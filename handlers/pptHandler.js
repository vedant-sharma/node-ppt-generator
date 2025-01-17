const PptxGenJS = require("pptxgenjs");
const { JSDOM } = require("jsdom");
const mjAPI = require("mathjax-node");
const sharp = require("sharp");
const { createCanvas } = require("canvas");

mjAPI.config({
  MathJax: {
    svg: {
      fontCache: "local",
    },
  },
});
mjAPI.start();

// Constants for layout
const MAX_IMAGE_WIDTH = 1.2;
const MAX_IMAGE_HEIGHT = 0.6;
const IMAGE_TEXT_SPACING = 0.1;
const IMAGE_ADJUSTMENT_Y = -0.2;
const LEFT_MARGIN = 0.5;
const SLIDE_WIDTH = 10;
const FONT_SIZE = 16;

// Function to clean up content data
const processContent = (content) => {
  // content = content.replace(/(<p[^>]+?>|<p>|<\/p>)/gim, "");

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

/**
 * Determine dynamic scaling factor based on LaTeX complexity.
 */
function getDynamicScalingFactor(latexCode) {
  const baseScale = 4; // Default scale for simple equations
  const complexityMultiplier =
    latexCode.match(/\\(frac|sum|int|sqrt|prod)/g)?.length || 0;
  const lengthMultiplier = Math.min(latexCode.length / 50, 2); // Scale more for longer equations

  const dynamicScale = baseScale + complexityMultiplier + lengthMultiplier;

  // Enforce a minimum scaling factor for short equations
  const minScale = 6; // Minimum scale to ensure small equations are readable
  return Math.max(dynamicScale, minScale);
}

// Function to measure text width
function measureTextWidth(text, fontFamily, fontSize) {
  const canvas = createCanvas(1, 1); // No need to create a visible canvas
  const context = canvas.getContext("2d");
  context.font = `${fontSize}pt ${fontFamily}`;
  return context.measureText(text).width / 96; // Convert pixels to inches (96 DPI)
}

// Function to convert LaTeX to a high-resolution PNG
async function latexToPngByteStream(latex) {
  return new Promise((resolve, reject) => {
    const scale = getDynamicScalingFactor(latex);

    mjAPI.typeset(
      {
        math: latex,
        format: "TeX",
        svg: true,
        scale: scale * 1.5, // High-resolution scale
      },
      (data) => {
        if (data.errors) {
          reject(data.errors);
        } else {
          sharp(Buffer.from(data.svg))
            // .resize(1500) // High resolution (width in px)
            .png({ quality: 100, adaptiveFiltering: true, dpi: 300 })
            .toBuffer()
            .then((buffer) => resolve(buffer))
            .catch((err) => reject(err));
        }
      }
    );
  });
}

async function processSlideContent(content) {
  // Normalize content by replacing '\\n' with '\n', and handle <p></p> tags
  const normalizedContent = content
    .replace(/\\n/g, "\n")
    .replace(/<\/?p[^>]*>/g, "\n"); // Handles all forms of <p>, </p>, <p />, <p class="some-class">, etc.

  // Now split based on the newline characters
  const segments = normalizedContent.split("\n");

  const parts = [];
  let series = 0;
  for (const segment of segments) {
    if (!segment.trim()) continue;

    const formulaRegex = /\$([^$]+)\$/g;
    let lastIndex = 0;
    let match;

    while ((match = formulaRegex.exec(segment)) !== null) {
      if (match.index > lastIndex) {
        parts.push({
          type: "text",
          series: series,
          value: segment.slice(lastIndex, match.index).trim(),
        });
      }

      const latex = match[1];
      const imageBuffer = await latexToPngByteStream(latex);

      // Get image dimensions from PNG buffer
      const metadata = await sharp(imageBuffer).metadata();
      const imgWidth = metadata.width;
      const imgHeight = metadata.height;

      // Calculate PowerPoint dimensions (inches, assuming 96 DPI)
      const pptxWidth = 10; // Default slide width
      const pptxHeight = 5.63; // Default slide height

      // Convert image dimensions to inches
      const dpi = 96;
      const imgAspectRatio = imgWidth / imgHeight;
      const maxImgWidthInches = pptxWidth * 0.8; // Allow 80% of slide width
      const imgWidthInches = Math.min(maxImgWidthInches, imgWidth / dpi);
      const imgHeightInches = imgWidthInches / imgAspectRatio;

      // const sizeMultiplier = latex.length > 25 ? 1 : 0.5;
      parts.push({
        type: "image",
        value: imageBuffer,
        series,
        series,
        width: imgWidthInches,
        height: imgHeightInches,
      });

      lastIndex = formulaRegex.lastIndex;
    }

    if (lastIndex < segment.length) {
      parts.push({
        type: "text",
        series: series,
        value: segment.slice(lastIndex).trim(),
      });
    }

    series++;
  }

  return parts;
}

// Generate PowerPoint presentation
const generatePpt = async (req, res) => {
  try {
    let { slides, template } = {
      slides: [
        {
          title: "Introduction to Fractions",
          sub_title: "Understanding Basics",
          content:
            '<p>Definition: A fraction represents a part of a whole. It consists of a numerator and a denominator. Example of a fraction: <span class="ql-custom-formula" data-value="\\frac{3}{4}">﻿<span contenteditable="false"><span class="katex"><span class="katex-mathml"><math><semantics><mrow><mfrac><mrow><mn>3</mn></mrow><mrow><mn>4</mn></mrow></mfrac></mrow><annotation encoding="application/x-tex">\\frac{3}{4}</annotation></semantics></math></span><span class="katex-html" aria-hidden="true"><span class="strut" style="height: 0.845108em;"></span><span class="strut bottom" style="height: 1.19011em; vertical-align: -0.345em;"></span><span class="base"><span class="mord"><span class="mopen nulldelimiter"></span><span class="mfrac"><span class="vlist-t vlist-t2"><span class="vlist-r"><span class="vlist" style="height: 0.845108em;"><span class="" style="top: -2.655em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">4</span></span></span></span><span class="" style="top: -3.23em;"><span class="pstrut" style="height: 3em;"></span><span class="frac-line" style="border-bottom-width: 0.04em;"></span></span><span class="" style="top: -3.394em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">3</span></span></span></span></span><span class="vlist-s">​</span></span><span class="vlist-r"><span class="vlist" style="height: 0.345em;"></span></span></span></span><span class="mclose nulldelimiter"></span></span></span></span></span></span>﻿</span> represents 3 parts out of 4 equal parts.</p>',
        },
        {
          title: "Operations with Fractions",
          sub_title: "Adding, Subtracting, Multiplying, and Dividing",
          content:
            '<p>Adding Fractions: To add fractions, make the denominators the same and add the numerators. Example:  <span class="ql-custom-formula" data-value="\\frac{1}{4} + \\frac{2}{4} = \\frac{3}{4}">﻿<span contenteditable="false"><span class="katex"><span class="katex-mathml"><math><semantics><mrow><mfrac><mrow><mn>1</mn></mrow><mrow><mn>4</mn></mrow></mfrac><mo>+</mo><mfrac><mrow><mn>2</mn></mrow><mrow><mn>4</mn></mrow></mfrac><mo>=</mo><mfrac><mrow><mn>3</mn></mrow><mrow><mn>4</mn></mrow></mfrac></mrow><annotation encoding="application/x-tex">\\frac{1}{4} + \\frac{2}{4} = \\frac{3}{4}</annotation></semantics></math></span><span class="katex-html" aria-hidden="true"><span class="strut" style="height: 0.845108em;"></span><span class="strut bottom" style="height: 1.19011em; vertical-align: -0.345em;"></span><span class="base"><span class="mord"><span class="mopen nulldelimiter"></span><span class="mfrac"><span class="vlist-t vlist-t2"><span class="vlist-r"><span class="vlist" style="height: 0.845108em;"><span class="" style="top: -2.655em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">4</span></span></span></span><span class="" style="top: -3.23em;"><span class="pstrut" style="height: 3em;"></span><span class="frac-line" style="border-bottom-width: 0.04em;"></span></span><span class="" style="top: -3.394em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">1</span></span></span></span></span><span class="vlist-s">​</span></span><span class="vlist-r"><span class="vlist" style="height: 0.345em;"></span></span></span></span><span class="mclose nulldelimiter"></span></span><span class="mbin">+</span><span class="mord"><span class="mopen nulldelimiter"></span><span class="mfrac"><span class="vlist-t vlist-t2"><span class="vlist-r"><span class="vlist" style="height: 0.845108em;"><span class="" style="top: -2.655em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">4</span></span></span></span><span class="" style="top: -3.23em;"><span class="pstrut" style="height: 3em;"></span><span class="frac-line" style="border-bottom-width: 0.04em;"></span></span><span class="" style="top: -3.394em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">2</span></span></span></span></span><span class="vlist-s">​</span></span><span class="vlist-r"><span class="vlist" style="height: 0.345em;"></span></span></span></span><span class="mclose nulldelimiter"></span></span><span class="mrel">=</span><span class="mord"><span class="mopen nulldelimiter"></span><span class="mfrac"><span class="vlist-t vlist-t2"><span class="vlist-r"><span class="vlist" style="height: 0.845108em;"><span class="" style="top: -2.655em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">4</span></span></span></span><span class="" style="top: -3.23em;"><span class="pstrut" style="height: 3em;"></span><span class="frac-line" style="border-bottom-width: 0.04em;"></span></span><span class="" style="top: -3.394em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">3</span></span></span></span></span><span class="vlist-s">​</span></span><span class="vlist-r"><span class="vlist" style="height: 0.345em;"></span></span></span></span><span class="mclose nulldelimiter"></span></span></span></span></span></span>﻿</span> Multiplying Fractions: Multiply the numerators and the denominators. Example: <span class="ql-custom-formula" data-value="\\frac{2}{3} \\times \\frac{3}{4} = \\frac{6}{12}">﻿<span contenteditable="false"><span class="katex"><span class="katex-mathml"><math><semantics><mrow><mfrac><mrow><mn>2</mn></mrow><mrow><mn>3</mn></mrow></mfrac><mo>×</mo><mfrac><mrow><mn>3</mn></mrow><mrow><mn>4</mn></mrow></mfrac><mo>=</mo><mfrac><mrow><mn>6</mn></mrow><mrow><mn>1</mn><mn>2</mn></mrow></mfrac></mrow><annotation encoding="application/x-tex">\\frac{2}{3} \\times \\frac{3}{4} = \\frac{6}{12}</annotation></semantics></math></span><span class="katex-html" aria-hidden="true"><span class="strut" style="height: 0.845108em;"></span><span class="strut bottom" style="height: 1.19011em; vertical-align: -0.345em;"></span><span class="base"><span class="mord"><span class="mopen nulldelimiter"></span><span class="mfrac"><span class="vlist-t vlist-t2"><span class="vlist-r"><span class="vlist" style="height: 0.845108em;"><span class="" style="top: -2.655em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">3</span></span></span></span><span class="" style="top: -3.23em;"><span class="pstrut" style="height: 3em;"></span><span class="frac-line" style="border-bottom-width: 0.04em;"></span></span><span class="" style="top: -3.394em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">2</span></span></span></span></span><span class="vlist-s">​</span></span><span class="vlist-r"><span class="vlist" style="height: 0.345em;"></span></span></span></span><span class="mclose nulldelimiter"></span></span><span class="mbin">×</span><span class="mord"><span class="mopen nulldelimiter"></span><span class="mfrac"><span class="vlist-t vlist-t2"><span class="vlist-r"><span class="vlist" style="height: 0.845108em;"><span class="" style="top: -2.655em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">4</span></span></span></span><span class="" style="top: -3.23em;"><span class="pstrut" style="height: 3em;"></span><span class="frac-line" style="border-bottom-width: 0.04em;"></span></span><span class="" style="top: -3.394em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">3</span></span></span></span></span><span class="vlist-s">​</span></span><span class="vlist-r"><span class="vlist" style="height: 0.345em;"></span></span></span></span><span class="mclose nulldelimiter"></span></span><span class="mrel">=</span><span class="mord"><span class="mopen nulldelimiter"></span><span class="mfrac"><span class="vlist-t vlist-t2"><span class="vlist-r"><span class="vlist" style="height: 0.845108em;"><span class="" style="top: -2.655em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">1</span><span class="mord mathrm mtight">2</span></span></span></span><span class="" style="top: -3.23em;"><span class="pstrut" style="height: 3em;"></span><span class="frac-line" style="border-bottom-width: 0.04em;"></span></span><span class="" style="top: -3.394em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">6</span></span></span></span></span><span class="vlist-s">​</span></span><span class="vlist-r"><span class="vlist" style="height: 0.345em;"></span></span></span></span><span class="mclose nulldelimiter"></span></span></span></span></span></span>﻿</span></p>',
        },
        {
          title: "Introduction to Trigonometry",
          sub_title: "Basics and Functions",
          content:
            "<p>Trigonometry deals with relationships between the angles and sides of triangles. Key Functions: - Sine (sin) - Cosine (cos) - Tangent (tan) These functions relate the angles of a triangle to the lengths of its sides.</p>",
        },
        {
          title: "Trigonometric Ratios and Identities",
          sub_title: "Practical Applications",
          content:
            '<p>The sine, cosine, and tangent functions are used to find unknown sides and angles in right triangles. Example Identity: <span class="ql-custom-formula" data-value="\\sin^2 \\theta + \\cos^2 \\theta = 1">﻿<span contenteditable="false"><span class="katex"><span class="katex-mathml"><math><semantics><mrow><msup><mi>sin</mi><mn>2</mn></msup><mi>θ</mi><mo>+</mo><msup><mi>cos</mi><mn>2</mn></msup><mi>θ</mi><mo>=</mo><mn>1</mn></mrow><annotation encoding="application/x-tex">\\sin^2 \\theta + \\cos^2 \\theta = 1</annotation></semantics></math></span><span class="katex-html" aria-hidden="true"><span class="strut" style="height: 0.871868em;"></span><span class="strut bottom" style="height: 0.955198em; vertical-align: -0.08333em;"></span><span class="base"><span class="mop"><span class="mop">sin</span><span class="msupsub"><span class="vlist-t"><span class="vlist-r"><span class="vlist" style="height: 0.871868em;"><span class="" style="top: -3.12076em; margin-right: 0.05em;"><span class="pstrut" style="height: 2.7em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mathrm mtight">2</span></span></span></span></span></span></span></span><span class="mord mathit" style="margin-right: 0.02778em;">θ</span><span class="mbin">+</span><span class="mop"><span class="mop">cos</span><span class="msupsub"><span class="vlist-t"><span class="vlist-r"><span class="vlist" style="height: 0.814108em;"><span class="" style="top: -3.063em; margin-right: 0.05em;"><span class="pstrut" style="height: 2.7em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mathrm mtight">2</span></span></span></span></span></span></span></span><span class="mord mathit" style="margin-right: 0.02778em;">θ</span><span class="mrel">=</span><span class="mord mathrm">1</span></span></span></span></span>﻿</span> This identity is fundamental in trigonometry and helps simplify complex equations.</p>',
        },
      ],
      template: {
        image_path:
          "https://prepaze-lms-store-staging.s3-us-west-2.amazonaws.com/public/1/1/SnnkCdjHBNH2bni3ZTPKp7/1-240RQ02Z3564.jpeg",
        font_family: "Times New Roman",
      },
    };

    slides.forEach((slide) => {
      slide.content = processContent(slide.content);
    });

    const pptx = new PptxGenJS();

    for (const slideData of slides) {
      const slide = pptx.addSlide();

      slide.background = { path: template.image_path };

      slide.addText(slideData.title, {
        x: 0.5,
        y: 0.5,
        fontSize: 24,
        bold: true,
        fontFace: template.font_family,
      });

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

      // Group the content by series
      const groupedContent = processedContent.reduce((acc, item) => {
        if (!acc[item.series]) acc[item.series] = [];
        acc[item.series].push(item);
        return acc;
      }, {});

      const LINE_HEIGHT = 0.35; // Standard line height
      const SERIES_BREAK_SPACING = 0.45; // Reduced spacing for series break
      const IMAGE_BASELINE_OFFSET = -0.15;
      const MAX_INLINE_SPACE = SLIDE_WIDTH - LEFT_MARGIN * 2;

      for (const series in groupedContent) {
        const seriesContent = groupedContent[series];
        // let xPosition = LEFT_MARGIN;

        let xPosition = LEFT_MARGIN; // Start at the left margin
        let yPosition = slideData.sub_title ? 1.8 : 1.5; // Start position, accounting for title/subtitle
        let currentLineHeight = 0; // Track the height of the current line (max of elements in the line)

        const hasImage = seriesContent.some((part) => part.type === "image");

        for (let i = 0; i < seriesContent.length; i++) {

          const part = seriesContent[i];
          const nextPart = seriesContent[i + 1]; // Look ahead to the next part

          if (hasImage) {
            if (part.type === "text") {
              // Append two whitespaces if the next part is an image
              if (nextPart && nextPart.type === "image") {
                part.value += "  ";
              }

              const words = part.value.split(" ");
              for (const word of words) {
                // Measure the actual width of the word in the specified font
                const wordWidth = measureTextWidth(
                  word,
                  template.font_family,
                  FONT_SIZE
                );
                const remainingSpace = MAX_INLINE_SPACE - xPosition;

                if (wordWidth > remainingSpace) {
                  // Move to the next line if the word doesn't fit
                  xPosition = LEFT_MARGIN;
                  yPosition += LINE_HEIGHT;
                }

                slide.addText(word, {
                  x: xPosition,
                  y: yPosition,
                  fontSize: FONT_SIZE,
                  fontFace: template.font_family,
                  color: "000000",
                  align: "left",
                });

                // Update xPosition with consistent spacing
                xPosition += wordWidth + 0.1; // Add minimal spacing after each word
              }
            } else if (part.type === "image") {
              const imageWidth = part.width;
              const imageHeight = part.height;

              if (xPosition + imageWidth > MAX_INLINE_SPACE) {
                // Move to the next line if the image doesn't fit
                xPosition = LEFT_MARGIN;
                yPosition += LINE_HEIGHT + IMAGE_TEXT_SPACING;
              }

              slide.addImage({
                data: `data:image/png;base64,${part.value.toString("base64")}`,
                x: xPosition,
                y: yPosition + IMAGE_BASELINE_OFFSET,
                w: imageWidth,
                h: imageHeight,
              });

              xPosition += imageWidth + IMAGE_TEXT_SPACING;
            }
          } else {
            if (part.type === "image") {
              const imageWidth = part.width;
              const imageHeight = part.height;

              // If the image doesn't fit horizontally, wrap to a new line
              if (xPosition + imageWidth > MAX_INLINE_SPACE) {
                xPosition = LEFT_MARGIN; // Reset xPosition for new line
                yPosition += currentLineHeight; // Move yPosition down for the new line
                currentLineHeight = 0; // Reset current line height
              }

              // Add the image to the slide
              slide.addImage({
                data: `data:image/png;base64,${part.value.toString("base64")}`,
                x: xPosition,
                y: yPosition,
                w: imageWidth,
                h: imageHeight,
              });

              // Update cursor position and current line height
              xPosition += imageWidth + IMAGE_TEXT_SPACING; // Move xPosition for the next content
              currentLineHeight = Math.max(currentLineHeight, imageHeight); // Adjust current line height
            } else if (part.type === "text") {
              const textWidth = part.value.length * 0.1; // Estimate text width (adjust as needed)

              // If the text doesn't fit horizontally, wrap to a new line
              if (xPosition + textWidth > MAX_INLINE_SPACE) {
                xPosition = LEFT_MARGIN; // Reset xPosition for new line
                yPosition += LINE_HEIGHT; // Move yPosition down for the new line
                currentLineHeight = 0; // Reset current line height
              }

              // Add the text to the slide
              slide.addText(part.value, {
                x: xPosition,
                y: yPosition,
                fontSize: FONT_SIZE,
                fontFace: template.font_family,
                color: "000000",
                align: "left",
              });

              // Update cursor position and current line height
              xPosition += textWidth + 0.15; // Add small spacing between text and next content
              currentLineHeight = Math.max(currentLineHeight, LINE_HEIGHT); // Adjust current line height
            }
          }
        }

        // Adjust spacing after each series
        yPosition += SERIES_BREAK_SPACING; // Smaller spacing for series change
        xPosition = LEFT_MARGIN;
      }
    }

    const buffer = await pptx.write("base64");
    const responseBuffer = Buffer.from(buffer, "base64");

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=GeneratedPresentation.pptx"
    );
    res.send(responseBuffer);
  } catch (err) {
    console.error("Error generating PPT:", err);
    res.status(500).send("Error generating PPT.");
  }
};

module.exports = { generatePpt };
