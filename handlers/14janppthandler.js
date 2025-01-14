const PptxGenJS = require("pptxgenjs");
const { JSDOM } = require("jsdom");
const mjAPI = require("mathjax-node");
const sharp = require("sharp");

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
  return baseScale + complexityMultiplier + lengthMultiplier;
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

// Process content by splitting on `\n` and `\\n`
// async function processSlideContent(content) {
//   const normalizedContent = content.replace(/\\n/g, "\n");
//   const segments = normalizedContent.split("\n");

//   const parts = [];
//   for (const segment of segments) {
//     if (!segment.trim()) continue;

//     const formulaRegex = /\$([^$]+)\$/g;
//     let lastIndex = 0;
//     let match;

//     while ((match = formulaRegex.exec(segment)) !== null) {
//       if (match.index > lastIndex) {
//         parts.push({ type: "text", value: segment.slice(lastIndex, match.index).trim() });
//       }

//       const latex = match[1];
//       const imageBuffer = await latexToPngByteStream(latex);

//       const sizeMultiplier = latex.length > 25 ? 1 : 0.5;
//       parts.push({
//         type: "image",
//         value: imageBuffer,
//         width: MAX_IMAGE_WIDTH * sizeMultiplier,
//         height: MAX_IMAGE_HEIGHT * sizeMultiplier,
//       });

//       lastIndex = formulaRegex.lastIndex;
//     }

//     if (lastIndex < segment.length) {
//       parts.push({ type: "text", value: segment.slice(lastIndex).trim() });
//     }
//   }

//   return parts;
// }

async function processSlideContent(content) {
  // Normalize content by replacing '\\n' with '\n', and handle <p></p> tags
  const normalizedContent = content
    .replace(/\\n/g, "\n")
    .replace(/<\/?p[^>]*>/g, "\n"); // Handles all forms of <p>, </p>, <p />, <p class="some-class">, etc.

  // console.log(normalizedContent);

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

      console.log(latex);

      // Get image dimensions from PNG buffer
      const metadata = await sharp(imageBuffer).metadata();
      const imgWidth = metadata.width;
      const imgHeight = metadata.height;

      console.log(`Image dimensions (px): ${imgWidth}x${imgHeight}`);

      // Calculate PowerPoint dimensions (inches, assuming 96 DPI)
      const pptxWidth = 10; // Default slide width
      const pptxHeight = 5.63; // Default slide height

      // Convert image dimensions to inches
      const dpi = 96;
      const imgAspectRatio = imgWidth / imgHeight;
      const maxImgWidthInches = pptxWidth * 0.8; // Allow 80% of slide width
      const imgWidthInches = Math.min(maxImgWidthInches, imgWidth / dpi);
      const imgHeightInches = imgWidthInches / imgAspectRatio;

      console.log(
        `Image dimensions (inches): ${imgWidthInches}x${imgHeightInches}`
      );

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
          title: "Introduction to Fractions in Quadratic Equations",
          sub_title: "Understanding the Basics",
          content:
            "Explore how fractions can appear in quadratic equations and their impact on solutions. Key concepts include:\n- Definition of Quadratic Equations.\n- Introduction to fractions within these equations.\n- Simplifying equations involving fractions.\n",
        },
        {
          title: "Solving Quadratic Equations with Fractions",
          sub_title: "Step-by-Step Approach",
          content:
            'Learn to solve equations where coefficients are fractions. Steps include:\n1. Clearing fractions by multiplying through by the LCD.\n2. Applying the quadratic formula: <span class="ql-custom-formula" data-value="x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}">﻿<span contenteditable="false"><span class="katex"><span class="katex-mathml"><math><semantics><mrow><mi>x</mi><mo>=</mo><mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo><msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt></mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></mrow><annotation encoding="application/x-tex">x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}</annotation></semantics></math></span><span class="katex-html" aria-hidden="true"><span class="strut" style="height: 1.03948em;"></span><span class="strut bottom" style="height: 1.38448em; vertical-align: -0.345em;"></span><span class="base"><span class="mord mathit">x</span><span class="mrel">=</span><span class="mord"><span class="mopen nulldelimiter"></span><span class="mfrac"><span class="vlist-t vlist-t2"><span class="vlist-r"><span class="vlist" style="height: 1.03948em;"><span class="" style="top: -2.655em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">2</span><span class="mord mathit mtight">a</span></span></span></span><span class="" style="top: -3.23em;"><span class="pstrut" style="height: 3em;"></span><span class="frac-line" style="border-bottom-width: 0.04em;"></span></span><span class="" style="top: -3.394em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mtight">−</span><span class="mord mathit mtight">b</span><span class="mbin mtight">±</span><span class="mord sqrt mtight"><span class="vlist-t vlist-t2"><span class="vlist-r"><span class="vlist svg-align" style="height: 0.922117em;"><span class="" style="top: -3em;"><span class="pstrut" style="height: 3em;"></span><span class="mord mtight" style="padding-left: 0.833em;"><span class="mord mtight"><span class="mord mathit mtight">b</span><span class="msupsub"><span class="vlist-t"><span class="vlist-r"><span class="vlist" style="height: 0.746314em;"><span class="" style="top: -2.786em; margin-right: 0.0714286em;"><span class="pstrut" style="height: 2.5em;"></span><span class="sizing reset-size3 size1 mtight"><span class="mord mathrm mtight">2</span></span></span></span></span></span></span></span><span class="mbin mtight">−</span><span class="mord mathrm mtight">4</span><span class="mord mathit mtight">a</span><span class="mord mathit mtight">c</span></span></span><span class="" style="top: -2.88212em;"><span class="pstrut" style="height: 3em;"></span><span class="hide-tail mtight" style="min-width: 0.853em; height: 1em;"><svg width="400em" height="1em" viewBox="0 0 400000 1000" preserveAspectRatio="xMinYMin slice"><path d="M95 622c-2.667 0-7.167-2.667-13.5\n-8S72 604 72 600c0-2 .333-3.333 1-4 1.333-2.667 23.833-20.667 67.5-54s\n65.833-50.333 66.5-51c1.333-1.333 3-2 5-2 4.667 0 8.667 3.333 12 10l173\n378c.667 0 35.333-71 104-213s137.5-285 206.5-429S812 17.333 812 14c5.333\n-9.333 12-14 20-14h399166v40H845.272L620 507 385 993c-2.667 4.667-9 7-19\n7-6 0-10-1-12-3L160 575l-65 47zM834 0h399166v40H845z"></path></svg></span></span></span><span class="vlist-s">​</span></span><span class="vlist-r"><span class="vlist" style="height: 0.117883em;"></span></span></span></span></span></span></span></span><span class="vlist-s">​</span></span><span class="vlist-r"><span class="vlist" style="height: 0.345em;"></span></span></span></span><span class="mclose nulldelimiter"></span></span></span></span></span></span>﻿</span>\n3. Example: Solve <span class="ql-custom-formula" data-value="\\frac{1}{2}x^2 - \\frac{5}{3}x + \\frac{1}{4} = 0">﻿<span contenteditable="false"><span class="katex"><span class="katex-mathml"><math><semantics><mrow><mfrac><mrow><mn>1</mn></mrow><mrow><mn>2</mn></mrow></mfrac><msup><mi>x</mi><mn>2</mn></msup><mo>−</mo><mfrac><mrow><mn>5</mn></mrow><mrow><mn>3</mn></mrow></mfrac><mi>x</mi><mo>+</mo><mfrac><mrow><mn>1</mn></mrow><mrow><mn>4</mn></mrow></mfrac><mo>=</mo><mn>0</mn></mrow><annotation encoding="application/x-tex">\\frac{1}{2}x^2 - \\frac{5}{3}x + \\frac{1}{4} = 0</annotation></semantics></math></span><span class="katex-html" aria-hidden="true"><span class="strut" style="height: 0.845108em;"></span><span class="strut bottom" style="height: 1.19011em; vertical-align: -0.345em;"></span><span class="base"><span class="mord"><span class="mopen nulldelimiter"></span><span class="mfrac"><span class="vlist-t vlist-t2"><span class="vlist-r"><span class="vlist" style="height: 0.845108em;"><span class="" style="top: -2.655em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">2</span></span></span></span><span class="" style="top: -3.23em;"><span class="pstrut" style="height: 3em;"></span><span class="frac-line" style="border-bottom-width: 0.04em;"></span></span><span class="" style="top: -3.394em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">1</span></span></span></span></span><span class="vlist-s">​</span></span><span class="vlist-r"><span class="vlist" style="height: 0.345em;"></span></span></span></span><span class="mclose nulldelimiter"></span></span><span class="mord"><span class="mord mathit">x</span><span class="msupsub"><span class="vlist-t"><span class="vlist-r"><span class="vlist" style="height: 0.814108em;"><span class="" style="top: -3.063em; margin-right: 0.05em;"><span class="pstrut" style="height: 2.7em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mathrm mtight">2</span></span></span></span></span></span></span></span><span class="mbin">−</span><span class="mord"><span class="mopen nulldelimiter"></span><span class="mfrac"><span class="vlist-t vlist-t2"><span class="vlist-r"><span class="vlist" style="height: 0.845108em;"><span class="" style="top: -2.655em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">3</span></span></span></span><span class="" style="top: -3.23em;"><span class="pstrut" style="height: 3em;"></span><span class="frac-line" style="border-bottom-width: 0.04em;"></span></span><span class="" style="top: -3.394em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">5</span></span></span></span></span><span class="vlist-s">​</span></span><span class="vlist-r"><span class="vlist" style="height: 0.345em;"></span></span></span></span><span class="mclose nulldelimiter"></span></span><span class="mord mathit">x</span><span class="mbin">+</span><span class="mord"><span class="mopen nulldelimiter"></span><span class="mfrac"><span class="vlist-t vlist-t2"><span class="vlist-r"><span class="vlist" style="height: 0.845108em;"><span class="" style="top: -2.655em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">4</span></span></span></span><span class="" style="top: -3.23em;"><span class="pstrut" style="height: 3em;"></span><span class="frac-line" style="border-bottom-width: 0.04em;"></span></span><span class="" style="top: -3.394em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">1</span></span></span></span></span><span class="vlist-s">​</span></span><span class="vlist-r"><span class="vlist" style="height: 0.345em;"></span></span></span></span><span class="mclose nulldelimiter"></span></span><span class="mrel">=</span><span class="mord mathrm">0</span></span></span></span></span>﻿</span> vedant sharmapush\n',
        },
        {
          title: "Examples and Practice Problems",
          sub_title: "Applying What You've Learned",
          content:
            'Work through detailed examples and practice solving quadratic equations that include fractions.\n- Example 1: <span class="ql-custom-formula" data-value="\\frac{3}{4}x^2 - x + \\frac{2}{5} = 0">﻿<span contenteditable="false"><span class="katex"><span class="katex-mathml"><math><semantics><mrow><mfrac><mrow><mn>3</mn></mrow><mrow><mn>4</mn></mrow></mfrac><msup><mi>x</mi><mn>2</mn></msup><mo>−</mo><mi>x</mi><mo>+</mo><mfrac><mrow><mn>2</mn></mrow><mrow><mn>5</mn></mrow></mfrac><mo>=</mo><mn>0</mn></mrow><annotation encoding="application/x-tex">\\frac{3}{4}x^2 - x + \\frac{2}{5} = 0</annotation></semantics></math></span><span class="katex-html" aria-hidden="true"><span class="strut" style="height: 0.845108em;"></span><span class="strut bottom" style="height: 1.19011em; vertical-align: -0.345em;"></span><span class="base"><span class="mord"><span class="mopen nulldelimiter"></span><span class="mfrac"><span class="vlist-t vlist-t2"><span class="vlist-r"><span class="vlist" style="height: 0.845108em;"><span class="" style="top: -2.655em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">4</span></span></span></span><span class="" style="top: -3.23em;"><span class="pstrut" style="height: 3em;"></span><span class="frac-line" style="border-bottom-width: 0.04em;"></span></span><span class="" style="top: -3.394em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">3</span></span></span></span></span><span class="vlist-s">​</span></span><span class="vlist-r"><span class="vlist" style="height: 0.345em;"></span></span></span></span><span class="mclose nulldelimiter"></span></span><span class="mord"><span class="mord mathit">x</span><span class="msupsub"><span class="vlist-t"><span class="vlist-r"><span class="vlist" style="height: 0.814108em;"><span class="" style="top: -3.063em; margin-right: 0.05em;"><span class="pstrut" style="height: 2.7em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mathrm mtight">2</span></span></span></span></span></span></span></span><span class="mbin">−</span><span class="mord mathit">x</span><span class="mbin">+</span><span class="mord"><span class="mopen nulldelimiter"></span><span class="mfrac"><span class="vlist-t vlist-t2"><span class="vlist-r"><span class="vlist" style="height: 0.845108em;"><span class="" style="top: -2.655em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">5</span></span></span></span><span class="" style="top: -3.23em;"><span class="pstrut" style="height: 3em;"></span><span class="frac-line" style="border-bottom-width: 0.04em;"></span></span><span class="" style="top: -3.394em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">2</span></span></span></span></span><span class="vlist-s">​</span></span><span class="vlist-r"><span class="vlist" style="height: 0.345em;"></span></span></span></span><span class="mclose nulldelimiter"></span></span><span class="mrel">=</span><span class="mord mathrm">0</span></span></span></span></span>﻿</span>\n- Practice Problem: Solve <span class="ql-custom-formula" data-value="\\frac{2}{3}x^2 - \\frac{3}{4}x + \\frac{1}{6} = 0">﻿<span contenteditable="false"><span class="katex"><span class="katex-mathml"><math><semantics><mrow><mfrac><mrow><mn>2</mn></mrow><mrow><mn>3</mn></mrow></mfrac><msup><mi>x</mi><mn>2</mn></msup><mo>−</mo><mfrac><mrow><mn>3</mn></mrow><mrow><mn>4</mn></mrow></mfrac><mi>x</mi><mo>+</mo><mfrac><mrow><mn>1</mn></mrow><mrow><mn>6</mn></mrow></mfrac><mo>=</mo><mn>0</mn></mrow><annotation encoding="application/x-tex">\\frac{2}{3}x^2 - \\frac{3}{4}x + \\frac{1}{6} = 0</annotation></semantics></math></span><span class="katex-html" aria-hidden="true"><span class="strut" style="height: 0.845108em;"></span><span class="strut bottom" style="height: 1.19011em; vertical-align: -0.345em;"></span><span class="base"><span class="mord"><span class="mopen nulldelimiter"></span><span class="mfrac"><span class="vlist-t vlist-t2"><span class="vlist-r"><span class="vlist" style="height: 0.845108em;"><span class="" style="top: -2.655em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">3</span></span></span></span><span class="" style="top: -3.23em;"><span class="pstrut" style="height: 3em;"></span><span class="frac-line" style="border-bottom-width: 0.04em;"></span></span><span class="" style="top: -3.394em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">2</span></span></span></span></span><span class="vlist-s">​</span></span><span class="vlist-r"><span class="vlist" style="height: 0.345em;"></span></span></span></span><span class="mclose nulldelimiter"></span></span><span class="mord"><span class="mord mathit">x</span><span class="msupsub"><span class="vlist-t"><span class="vlist-r"><span class="vlist" style="height: 0.814108em;"><span class="" style="top: -3.063em; margin-right: 0.05em;"><span class="pstrut" style="height: 2.7em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mathrm mtight">2</span></span></span></span></span></span></span></span><span class="mbin">−</span><span class="mord"><span class="mopen nulldelimiter"></span><span class="mfrac"><span class="vlist-t vlist-t2"><span class="vlist-r"><span class="vlist" style="height: 0.845108em;"><span class="" style="top: -2.655em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">4</span></span></span></span><span class="" style="top: -3.23em;"><span class="pstrut" style="height: 3em;"></span><span class="frac-line" style="border-bottom-width: 0.04em;"></span></span><span class="" style="top: -3.394em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">3</span></span></span></span></span><span class="vlist-s">​</span></span><span class="vlist-r"><span class="vlist" style="height: 0.345em;"></span></span></span></span><span class="mclose nulldelimiter"></span></span><span class="mord mathit">x</span><span class="mbin">+</span><span class="mord"><span class="mopen nulldelimiter"></span><span class="mfrac"><span class="vlist-t vlist-t2"><span class="vlist-r"><span class="vlist" style="height: 0.845108em;"><span class="" style="top: -2.655em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">6</span></span></span></span><span class="" style="top: -3.23em;"><span class="pstrut" style="height: 3em;"></span><span class="frac-line" style="border-bottom-width: 0.04em;"></span></span><span class="" style="top: -3.394em;"><span class="pstrut" style="height: 3em;"></span><span class="sizing reset-size6 size3 mtight"><span class="mord mtight"><span class="mord mathrm mtight">1</span></span></span></span></span><span class="vlist-s">​</span></span><span class="vlist-r"><span class="vlist" style="height: 0.345em;"></span></span></span></span><span class="mclose nulldelimiter"></span></span><span class="mrel">=</span><span class="mord mathrm">0</span></span></span></span></span>﻿</span>\nThese exercises help solidify understanding and technique.\n',
        },
      ],
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
      let xPosition = LEFT_MARGIN;
      let yPosition = slideData.sub_title ? 1.8 : 1.5;

      console.log("processedContent ", processedContent);

      // for (const part of processedContent) {

      //   if (part.type === "text") {
      //     const lines = part.value.split("\n");
      //     console.log("lines ", lines);
      //     for (const line of lines) {
      //       console.log("line ", line);

      //       const textWidth = line.length * 0.1; // Approximate width per character

      //       if (xPosition + textWidth > SLIDE_WIDTH - LEFT_MARGIN) {
      //         // Move to the next line if text exceeds available space
      //         xPosition = LEFT_MARGIN;
      //         yPosition += 0.5;
      //       }

      //       slide.addText(line, {
      //         x: xPosition,
      //         y: yPosition,
      //         w: SLIDE_WIDTH - LEFT_MARGIN * 2,
      //         fontSize: FONT_SIZE,
      //         fontFace: template.font_family,
      //         color: "000000",
      //         align: "left",
      //       });

      //       yPosition += 0.5; // Advance the Y position for the next line

      //       // xPosition += textWidth + 0.2; // Move xPosition after text
      //     }
      //   } else if (part.type === "image") {
      //     if (xPosition + part.width > SLIDE_WIDTH - LEFT_MARGIN) {
      //       xPosition = LEFT_MARGIN;
      //       yPosition += MAX_IMAGE_HEIGHT + IMAGE_TEXT_SPACING;
      //     }

      //     slide.addImage({
      //       data: `data:image/png;base64,${part.value.toString("base64")}`,
      //       x: xPosition,
      //       y: yPosition + IMAGE_ADJUSTMENT_Y,
      //       w: part.width,
      //       h: part.height
      //     });
      //     // console.log("WIDTH -- ", part.width)
      //     // console.log(("HERIGHT -- ", part.width / 16) * 9);
      //     // yPosition += 0.5; // Advance the Y position for the next line
      //     xPosition += part.width + IMAGE_TEXT_SPACING;
      //   }

      //   // Reset to the next line if xPosition exceeds the slide width
      //   if (xPosition > SLIDE_WIDTH - LEFT_MARGIN) {
      //     xPosition = LEFT_MARGIN;
      //     yPosition += 0.5;
      //   }
      // }

      console.log("processedContent ", processedContent);

      // Group the content by series
      const groupedContent = processedContent.reduce((acc, item) => {
        if (!acc[item.series]) acc[item.series] = [];
        acc[item.series].push(item);
        return acc;
      }, {});

      for (const series in groupedContent) {
        let seriesContent = groupedContent[series];
        let xPosition = LEFT_MARGIN;
        // let yPosition = CURRENT_Y_POSITION; // Initialize Y position
        let yPosition = slideData.sub_title ? 1.8 : 1.5;

        for (const part of seriesContent) {
          if (part.type === "text") {
            const lines = part.value.split("\n");
            for (const line of lines) {
              const textWidth = line.length * 0.1; // Approximate width per character

              if (xPosition + textWidth > SLIDE_WIDTH - LEFT_MARGIN) {
                // Move to the next line if text exceeds available space
                xPosition = LEFT_MARGIN;
                yPosition += 0.5;
              }

              slide.addText(line, {
                x: xPosition,
                y: yPosition,
                w: SLIDE_WIDTH - LEFT_MARGIN * 2,
                fontSize: FONT_SIZE,
                fontFace: template.font_family,
                color: "000000",
                align: "left",
              });

              xPosition += textWidth + 0.2; // Advance X position for inline continuation
            }
          } else if (part.type === "image") {
            const imageWidth = part.width;
            const imageHeight = part.height 

            if (xPosition + imageWidth > SLIDE_WIDTH - LEFT_MARGIN) {
              // Move to the next line if image exceeds available space
              xPosition = LEFT_MARGIN;
              yPosition += MAX_IMAGE_HEIGHT + IMAGE_TEXT_SPACING;
            }

            slide.addImage({
              data: `data:image/png;base64,${part.value.toString("base64")}`,
              x: xPosition,
              y: yPosition,
              w: imageWidth,
              h: imageHeight,
            });

            xPosition += imageWidth + IMAGE_TEXT_SPACING; // Advance X position for inline continuation
          }

          // Handle overflow to move to the next line within the series
          if (xPosition > SLIDE_WIDTH - LEFT_MARGIN) {
            xPosition = LEFT_MARGIN;
            yPosition += 0.5;
          }
        }

        // Reset for the next series to start on a new line
        xPosition = LEFT_MARGIN;
        yPosition += 1; // Additional spacing between series
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
