const PptxGenJS = require('pptxgenjs');
const path = require('path');

// Define the slide background path relative to the handler
const slideBackgroundPath = path.join(__dirname, '../public/images/slide_background.jpeg');

const generatePpt = async (req, res) => {
  try {
    // Create a new presentation
    const pptx = new PptxGenJS();

    // Add a slide
    const slide1 = pptx.addSlide();
    slide1.addText('Slide 1', {
      x: 1,
      y: 1,
      w: 8,
      h: 2,
      fontSize: 24,
      bold: true,
      color: '363636',
      align: 'center',
    });
    slide1.background = { path: slideBackgroundPath };

    // Add a second slide
    const slide2 = pptx.addSlide();
    slide2.addText('Slide 2', {
      x: 1,
      y: 1,
      w: 8,
      h: 2,
      fontSize: 24,
      bold: true,
      color: '363636',
      align: 'center',
    });
    slide2.background = { path: slideBackgroundPath };

    // Add a math equation
    const mathEquation = 'E = mc²';
    slide2.addText(mathEquation, { x: 1, y: 1, fontSize: 24 });

    // Save the presentation
    await pptx.writeFile('SamplePresentation.pptx');
    console.log('Presentation created successfully!');

    res.send('SUCCESS :: PPT created successfully!');
  } catch (error) {
    console.error('Error creating presentation:', error);
    res.status(500).send('Failed to create PPT.');
  }
};

module.exports = { generatePpt };
