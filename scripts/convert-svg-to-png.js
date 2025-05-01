const sharp = require('sharp');
const fs = require('fs');
const path = require('path');

// Path to the SVG file
const svgPath = path.join(__dirname, '..', 'assets', 'cori-logo.svg');
const outputDir = path.join(__dirname, '..', 'assets');

// Read the SVG file
const svgBuffer = fs.readFileSync(svgPath);

// Define the sizes needed for Office Add-in
const sizes = [16, 32, 64, 80, 128];

// Process each size
async function convertToAllSizes() {
  for (const size of sizes) {
    try {
      await sharp(svgBuffer)
        .resize(size, size)
        .png()
        .toFile(path.join(outputDir, `icon-${size}.png`));
      
      console.log(`Successfully created icon-${size}.png`);
    } catch (error) {
      console.error(`Error creating icon-${size}.png:`, error);
    }
  }
}

convertToAllSizes().then(() => {
  console.log('All icons created successfully!');
}).catch(err => {
  console.error('Error in conversion process:', err);
});
