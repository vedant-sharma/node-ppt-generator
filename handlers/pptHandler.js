const PptxGenJS = require("pptxgenjs");
const { JSDOM } = require("jsdom");
const mjAPI = require("mathjax-node");
const sharp = require("sharp");
const { createCanvas } = require("canvas");
const axios = require("axios");

mjAPI.config({
  MathJax: {
    svg: {
      fontCache: "local",
    },
  },
});
mjAPI.start();

// Constants for layout
const IMAGE_TEXT_SPACING = -0.1;
const LEFT_MARGIN = 0.7;
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

// Function to fetch an image from a URL and convert it to Base64
const getBase64FromUrl = async (url) => {
  const response = await axios.get(url, { responseType: "arraybuffer" });
  const buffer = Buffer.from(response.data, "binary");
  return `data:image/jpeg;base64,${buffer.toString("base64")}`;
};

/**
 * Determine dynamic scaling factor based on LaTeX complexity.
 */
function getDynamicScalingFactor(latexCode) {
  const baseScale = 5; // Default scale for simple equations
  const complexityMultiplier =
    (latexCode.match(/\\(frac|sum|int|sqrt|prod)/g) || []).length || 0;
  const lengthMultiplier = Math.min(latexCode.length / 50, 2); // Scale more for longer equations

  const dynamicScale = baseScale + complexityMultiplier + lengthMultiplier;

  // Enforce a minimum scaling factor for short equations
  const minScale = 100; // Minimum scale to ensure small equations are readable
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
        scale: scale * 4, // High-resolution scale
      },
      (data) => {
        if (data.errors) {
          reject(data.errors);
        } else {
          sharp(Buffer.from(data.svg))
            // .resize(resizeValue) // High resolution (width in px)
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
  .replace(/<p>\s*<br\s*\/?>\s*<\/p>/gi, "\n") // Handle <p><br></p> as a single newline
    .replace(/\\n/g, "\n")
    .replace(/<br\s*\/?>/gi, "\n") // Replace br tag with \n
    .replace(/<\/p>/gi, "\n") // Replace only </p> with newlines
    .replace(/<p>/gi, ""); // Replace only <p> with an empty string

    console.log("nrma -- ", normalizedContent)
  // Now split based on the newline characters
  const segments = normalizedContent.split("\n");

  console.log("segments -- ", segments)
  const parts = [];
  let series = 0;
  for (const segment of segments) {
    if (!segment.trim()) {
      parts.push({
        type: "break",
        series,
      });
    };

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
      const imgHeight = Math.max(metadata.height, 24);

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
    slides: [{"title":"Introduction to Slope","sub_title":"Understanding the Basics","content":"<p>Definition of Slope: Measure of steepness of a line. </p><p>Formula: Slope (m) = <span class=\"ql-custom-formula\" data-value=\"\\frac{\\text{rise}}{\\text{run}}\">﻿<span contenteditable=\"false\"><span class=\"katex\"><span class=\"katex-mathml\"><math><semantics><mrow><mfrac><mrow><mtext>rise</mtext></mrow><mrow><mtext>run</mtext></mrow></mfrac></mrow><annotation encoding=\"application/x-tex\">\\frac{\\text{rise}}{\\text{run}}</annotation></semantics></math></span><span class=\"katex-html\" aria-hidden=\"true\"><span class=\"strut\" style=\"height: 0.861502em;\"></span><span class=\"strut bottom\" style=\"height: 1.2065em; vertical-align: -0.345em;\"></span><span class=\"base\"><span class=\"mord\"><span class=\"mopen nulldelimiter\"></span><span class=\"mfrac\"><span class=\"vlist-t vlist-t2\"><span class=\"vlist-r\"><span class=\"vlist\" style=\"height: 0.861502em;\"><span class=\"\" style=\"top: -2.655em;\"><span class=\"pstrut\" style=\"height: 3em;\"></span><span class=\"sizing reset-size6 size3 mtight\"><span class=\"mord mtight\"><span class=\"mord text mtight\"><span class=\"mord mathrm mtight\">run</span></span></span></span></span><span class=\"\" style=\"top: -3.23em;\"><span class=\"pstrut\" style=\"height: 3em;\"></span><span class=\"frac-line\" style=\"border-bottom-width: 0.04em;\"></span></span><span class=\"\" style=\"top: -3.394em;\"><span class=\"pstrut\" style=\"height: 3em;\"></span><span class=\"sizing reset-size6 size3 mtight\"><span class=\"mord mtight\"><span class=\"mord text mtight\"><span class=\"mord mathrm mtight\">rise</span></span></span></span></span></span><span class=\"vlist-s\">​</span></span><span class=\"vlist-r\"><span class=\"vlist\" style=\"height: 0.345em;\"></span></span></span></span><span class=\"mclose nulldelimiter\"></span></span></span></span></span></span>﻿</span></p><p>Importance: Helps in understanding the direction and steepness of the line.</p>"},{"title":"Method 1: Graphical Method","sub_title":"Finding Slope from a Graph","content":"<p>Steps to find the slope from a graph:</p><p>1. Identify two points on the line.</p><p>2. Calculate the rise (vertical change) between the points.</p><p><br></p><p>3. Calculate the run (horizontal change) between the points.</p><p>4. Use the slope formula: <span class=\"ql-custom-formula\" data-value=\"m = \\frac{\\text{rise}}{\\text{run}}\">﻿<span contenteditable=\"false\"><span class=\"katex\"><span class=\"katex-mathml\"><math><semantics><mrow><mi>m</mi><mo>=</mo><mfrac><mrow><mtext>rise</mtext></mrow><mrow><mtext>run</mtext></mrow></mfrac></mrow><annotation encoding=\"application/x-tex\">m = \\frac{\\text{rise}}{\\text{run}}</annotation></semantics></math></span><span class=\"katex-html\" aria-hidden=\"true\"><span class=\"strut\" style=\"height: 0.861502em;\"></span><span class=\"strut bottom\" style=\"height: 1.2065em; vertical-align: -0.345em;\"></span><span class=\"base\"><span class=\"mord mathit\">m</span><span class=\"mrel\">=</span><span class=\"mord\"><span class=\"mopen nulldelimiter\"></span><span class=\"mfrac\"><span class=\"vlist-t vlist-t2\"><span class=\"vlist-r\"><span class=\"vlist\" style=\"height: 0.861502em;\"><span class=\"\" style=\"top: -2.655em;\"><span class=\"pstrut\" style=\"height: 3em;\"></span><span class=\"sizing reset-size6 size3 mtight\"><span class=\"mord mtight\"><span class=\"mord text mtight\"><span class=\"mord mathrm mtight\">run</span></span></span></span></span><span class=\"\" style=\"top: -3.23em;\"><span class=\"pstrut\" style=\"height: 3em;\"></span><span class=\"frac-line\" style=\"border-bottom-width: 0.04em;\"></span></span><span class=\"\" style=\"top: -3.394em;\"><span class=\"pstrut\" style=\"height: 3em;\"></span><span class=\"sizing reset-size6 size3 mtight\"><span class=\"mord mtight\"><span class=\"mord text mtight\"><span class=\"mord mathrm mtight\">rise</span></span></span></span></span></span><span class=\"vlist-s\">​</span></span><span class=\"vlist-r\"><span class=\"vlist\" style=\"height: 0.345em;\"></span></span></span></span><span class=\"mclose nulldelimiter\"></span></span></span></span></span></span>﻿</span>.</p>"},{"title":"Method 2: Algebraic Method","sub_title":"Using Coordinates","content":"<p>Steps to find the slope using coordinates: </p><p>1. Take two points (x1, y1) and (x2, y2). </p><p>2. Substitute into the formula: <span class=\"ql-custom-formula\" data-value=\"m = \\frac{y2 - y1}{x2 - x1}\">﻿<span contenteditable=\"false\"><span class=\"katex\"><span class=\"katex-mathml\"><math><semantics><mrow><mi>m</mi><mo>=</mo><mfrac><mrow><mi>y</mi><mn>2</mn><mo>−</mo><mi>y</mi><mn>1</mn></mrow><mrow><mi>x</mi><mn>2</mn><mo>−</mo><mi>x</mi><mn>1</mn></mrow></mfrac></mrow><annotation encoding=\"application/x-tex\">m = \\frac{y2 - y1}{x2 - x1}</annotation></semantics></math></span><span class=\"katex-html\" aria-hidden=\"true\"><span class=\"strut\" style=\"height: 0.897216em;\"></span><span class=\"strut bottom\" style=\"height: 1.30055em; vertical-align: -0.403331em;\"></span><span class=\"base\"><span class=\"mord mathit\">m</span><span class=\"mrel\">=</span><span class=\"mord\"><span class=\"mopen nulldelimiter\"></span><span class=\"mfrac\"><span class=\"vlist-t vlist-t2\"><span class=\"vlist-r\"><span class=\"vlist\" style=\"height: 0.897216em;\"><span class=\"\" style=\"top: -2.655em;\"><span class=\"pstrut\" style=\"height: 3em;\"></span><span class=\"sizing reset-size6 size3 mtight\"><span class=\"mord mtight\"><span class=\"mord mathit mtight\">x</span><span class=\"mord mathrm mtight\">2</span><span class=\"mbin mtight\">−</span><span class=\"mord mathit mtight\">x</span><span class=\"mord mathrm mtight\">1</span></span></span></span><span class=\"\" style=\"top: -3.23em;\"><span class=\"pstrut\" style=\"height: 3em;\"></span><span class=\"frac-line\" style=\"border-bottom-width: 0.04em;\"></span></span><span class=\"\" style=\"top: -3.44611em;\"><span class=\"pstrut\" style=\"height: 3em;\"></span><span class=\"sizing reset-size6 size3 mtight\"><span class=\"mord mtight\"><span class=\"mord mathit mtight\" style=\"margin-right: 0.03588em;\">y</span><span class=\"mord mathrm mtight\">2</span><span class=\"mbin mtight\">−</span><span class=\"mord mathit mtight\" style=\"margin-right: 0.03588em;\">y</span><span class=\"mord mathrm mtight\">1</span></span></span></span></span><span class=\"vlist-s\">​</span></span><span class=\"vlist-r\"><span class=\"vlist\" style=\"height: 0.403331em;\"></span></span></span></span><span class=\"mclose nulldelimiter\"></span></span></span></span></span></span>﻿</span>.</p><p>Example: For points (1,2) and (3,4), Slope = <span class=\"ql-custom-formula\" data-value=\"m = \\frac{4 - 2}{3 - 1}\">﻿<span contenteditable=\"false\"><span class=\"katex\"><span class=\"katex-mathml\"><math><semantics><mrow><mi>m</mi><mo>=</mo><mfrac><mrow><mn>4</mn><mo>−</mo><mn>2</mn></mrow><mrow><mn>3</mn><mo>−</mo><mn>1</mn></mrow></mfrac></mrow><annotation encoding=\"application/x-tex\">m = \\frac{4 - 2}{3 - 1}</annotation></semantics></math></span><span class=\"katex-html\" aria-hidden=\"true\"><span class=\"strut\" style=\"height: 0.845108em;\"></span><span class=\"strut bottom\" style=\"height: 1.24844em; vertical-align: -0.403331em;\"></span><span class=\"base\"><span class=\"mord mathit\">m</span><span class=\"mrel\">=</span><span class=\"mord\"><span class=\"mopen nulldelimiter\"></span><span class=\"mfrac\"><span class=\"vlist-t vlist-t2\"><span class=\"vlist-r\"><span class=\"vlist\" style=\"height: 0.845108em;\"><span class=\"\" style=\"top: -2.655em;\"><span class=\"pstrut\" style=\"height: 3em;\"></span><span class=\"sizing reset-size6 size3 mtight\"><span class=\"mord mtight\"><span class=\"mord mathrm mtight\">3</span><span class=\"mbin mtight\">−</span><span class=\"mord mathrm mtight\">1</span></span></span></span><span class=\"\" style=\"top: -3.23em;\"><span class=\"pstrut\" style=\"height: 3em;\"></span><span class=\"frac-line\" style=\"border-bottom-width: 0.04em;\"></span></span><span class=\"\" style=\"top: -3.394em;\"><span class=\"pstrut\" style=\"height: 3em;\"></span><span class=\"sizing reset-size6 size3 mtight\"><span class=\"mord mtight\"><span class=\"mord mathrm mtight\">4</span><span class=\"mbin mtight\">−</span><span class=\"mord mathrm mtight\">2</span></span></span></span></span><span class=\"vlist-s\">​</span></span><span class=\"vlist-r\"><span class=\"vlist\" style=\"height: 0.403331em;\"></span></span></span></span><span class=\"mclose nulldelimiter\"></span></span></span></span></span></span>﻿</span> = 1.</p>"}]       ,
    template: {
        image_path:
          "https://prepaze-lms-store-staging.s3-us-west-2.amazonaws.com/public/1/1/SnnkCdjHBNH2bni3ZTPKp7/1-240RQ02Z3564.jpeg",
        theme: "Times New Roman",
      },
    };
    
    slides.forEach((slide) => {
      slide.content = processContent(slide.content);
    });

    const pptx = new PptxGenJS();

    for (const slideData of slides) {
      const slide = pptx.addSlide();

      // Convert the URL to a Base64-encoded string
      const base64Image = await getBase64FromUrl(template.image_path);

      // Set the slide background using the base64 image
      slide.background = { data: base64Image };

      // slide.background = { path: template.image_path };

      slide.addText(slideData.title, {
        x: LEFT_MARGIN,
        y: 1.2,
        fontSize: 24,
        bold: true,
        fontFace: template.font_family,
      });

      if (slideData.sub_title) {
        slide.addText(slideData.sub_title, {
          x: LEFT_MARGIN,
          y: 1.7,
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
      var IMAGE_BASELINE_OFFSET = 0.04;
      const MAX_INLINE_SPACE = SLIDE_WIDTH - LEFT_MARGIN * 2;
      let yPosition = slideData.sub_title ? 2.2 : 1.9; // Start position, accounting for title/subtitle

      for (const series in groupedContent) {
        const seriesContent = groupedContent[series];
        // let xPosition = LEFT_MARGIN;

        let xPosition = LEFT_MARGIN; // Start at the left margin
        let currentLineHeight = 0; // Track the height of the current line (max of elements in the line)

        const hasImage = seriesContent.some((part) => part.type === "image");

        for (let i = 0; i < seriesContent.length; i++) {
          const part = seriesContent[i];
          const nextPart = seriesContent[i + 1]; // Look ahead to the next part

          if (hasImage) {
            if (part.type === "text") {
              // Append two whitespaces if the next part is an image
              if (nextPart && nextPart.type === "image") {
                part.value += " ";
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
                  y: yPosition + 0.2,
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

              // Define minimum dimensions
              const minImgWidthInches = 0.22; // Minimum image width in inches
              const imgAspectRatio = imageWidth / imageHeight;

              // Apply minimum size while preserving aspect ratio
              const finalWidthInches = Math.max(imageWidth, minImgWidthInches);
              const finalHeightInches = finalWidthInches / imgAspectRatio;

              const finalImgWidth = Math.min(
                finalWidthInches,
                MAX_INLINE_SPACE - xPosition
              );
              const finalImgHeight =
                finalWidthInches < imageWidth ? finalHeightInches : imageHeight;

              if (xPosition + finalImgWidth > MAX_INLINE_SPACE) {
                // Move to the next line if the image doesn't fit
                xPosition = LEFT_MARGIN;
                yPosition += LINE_HEIGHT + IMAGE_TEXT_SPACING;
              }

              if (part.height < 0.2) {
                IMAGE_BASELINE_OFFSET = 0.12;
              }

              slide.addImage({
                data: `data:image/png;base64,${part.value.toString("base64")}`,
                x: xPosition,
                y: yPosition + IMAGE_BASELINE_OFFSET,
                w: finalImgWidth,
                h: finalImgHeight,
              });

              // Check the next part for a-z or A-Z
              const nextPart = seriesContent[i + 1];
              let spacing = IMAGE_TEXT_SPACING;

              if (nextPart && nextPart.type === "text") {
                const nextChar = nextPart.value.trim().charAt(0); // Get the first character of the next text
                if (/[.)]/.test(nextChar)) {
                  spacing = -0.05; // Use tighter spacing if next character is '.' or ')'
                } else {
                  spacing = 0.1; // Use normal spacing otherwise
                }
              }
            
              // Update xPosition with the calculated spacing
              xPosition += imageWidth + spacing;

              // xPosition += finalImgWidth + IMAGE_TEXT_SPACING;
            } else if (part.type === "break") {
              // Add spacing for line breaks
              yPosition += SERIES_BREAK_SPACING;
              xPosition = LEFT_MARGIN; // Reset x position
            }
          } else {
            if (part.type == "text" ){

              const textWidth = part.value.length * 0.1; // Estimate text width (adjust as needed)
  
              // If the text doesn't fit horizontally, wrap to a new line
              if (xPosition + textWidth > MAX_INLINE_SPACE) {
                xPosition = LEFT_MARGIN; // Reset xPosition for new line
                yPosition += currentLineHeight; // Move yPosition down for the new line
                currentLineHeight = 0; // Reset current line height
              }
  
              // Add the text to the slide
              slide.addText(part.value, {
                x: xPosition,
                y: yPosition + 0.2,
                fontSize: FONT_SIZE,
                fontFace: template.font_family,
                color: "000000",
                align: "left",
              });
  
              // Update cursor position and current line height
              xPosition += textWidth + 0.15; // Add small spacing between text and next content
              currentLineHeight = Math.max(currentLineHeight, LINE_HEIGHT); // Adjust current line height
            } else if (part.type === "break") {
              // Add spacing for line breaks
              yPosition += SERIES_BREAK_SPACING;
              xPosition = LEFT_MARGIN; // Reset x position
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
