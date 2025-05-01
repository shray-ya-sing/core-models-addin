const sharp = require('sharp');
const fs = require('fs');
const path = require('path');

// Path to the SVG file
const svgPath = path.join(__dirname, '..', 'assets', 'cori-logo.svg');
const outputDir = path.join(__dirname, '..', 'assets');
const outputFile = path.join(outputDir, 'logo-filled.png');

// Read the SVG file
const svgBuffer = fs.readFileSync(svgPath);

// Create a 300x300 version
async function createLargeLogo() {
  try {
    await sharp(svgBuffer)
      .resize(300, 300)
      .png()
      .toFile(outputFile);
    
    console.log(`Successfully created logo-filled.png (300x300)`);
  } catch (error) {
    console.error(`Error creating logo-filled.png:`, error);
  }
}

createLargeLogo().then(() => {
  console.log('Large logo created successfully!');
}).catch(err => {
  console.error('Error in conversion process:', err);
});
