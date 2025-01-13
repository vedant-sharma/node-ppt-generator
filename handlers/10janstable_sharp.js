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
const LEFT_MARGIN = 1;
const SLIDE_WIDTH = 10;
const FONT_SIZE = 16;

// Function to clean up content data
const processContent = (content) => {
  content = content.replace(/(<p[^>]+?>|<p>|<\/p>)/gim, "");

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

// Function to convert LaTeX to a high-resolution PNG
async function latexToPngByteStream(latex) {
  return new Promise((resolve, reject) => {
    mjAPI.typeset(
      {
        math: latex,
        format: "TeX",
        svg: true,
      },
      (data) => {
        if (data.errors) {
          reject(data.errors);
        } else {
          sharp(Buffer.from(data.svg))
            .resize(1500) // High resolution (width in px)
            .png()
            .toBuffer()
            .then((buffer) => resolve(buffer))
            .catch((err) => reject(err));
        }
      }
    );
  });
}

// Process content by splitting on `\n` and `\\n`
async function processSlideContent(content) {
  const normalizedContent = content.replace(/\\n/g, "\n");
  const segments = normalizedContent.split("\n");

  const parts = [];
  for (const segment of segments) {
    if (!segment.trim()) continue;

    const formulaRegex = /\$([^$]+)\$/g;
    let lastIndex = 0;
    let match;

    while ((match = formulaRegex.exec(segment)) !== null) {
      if (match.index > lastIndex) {
        parts.push({ type: "text", value: segment.slice(lastIndex, match.index).trim() });
      }

      const latex = match[1];
      const imageBuffer = await latexToPngByteStream(latex);

      const sizeMultiplier = latex.length > 25 ? 1 : 0.5;
      parts.push({
        type: "image",
        value: imageBuffer,
        width: MAX_IMAGE_WIDTH * sizeMultiplier,
        height: MAX_IMAGE_HEIGHT * sizeMultiplier,
      });

      lastIndex = formulaRegex.lastIndex;
    }

    if (lastIndex < segment.length) {
      parts.push({ type: "text", value: segment.slice(lastIndex).trim() });
    }
  }

  return parts;
}

// Generate PowerPoint presentation
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

      for (const part of processedContent) {
        if (part.type === "text") {
          const lines = part.value.split("\n");
          for (const line of lines) {
            slide.addText(line, {
              x: xPosition,
              y: yPosition,
              w: SLIDE_WIDTH - LEFT_MARGIN * 2,
              fontSize: FONT_SIZE,
              fontFace: template.font_family,
              color: "000000",
              align: "left",
            });

            yPosition += 0.5;
          }
        } else if (part.type === "image") {
          if (xPosition + part.width > SLIDE_WIDTH - LEFT_MARGIN) {
            xPosition = LEFT_MARGIN;
            yPosition += MAX_IMAGE_HEIGHT + IMAGE_TEXT_SPACING;
          }

          slide.addImage({
            data: `data:image/png;base64,${part.value.toString("base64")}`,
            x: xPosition,
            y: yPosition + IMAGE_ADJUSTMENT_Y,
            w: part.width,
            h: part.height,
          });

          xPosition += part.width + IMAGE_TEXT_SPACING;
        }
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
