const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// Check if sharp is installed
try {
  require.resolve('sharp');
} catch (e) {
  console.log('Installing sharp package...');
  execSync('npm install sharp', { stdio: 'inherit' });
}

const sharp = require('sharp');

// Convert SVG to PNG
async function convertSvgToPng() {
  const svgPath = path.join(__dirname, 'assets', 'cori-favicon.svg');
  const pngPath = path.join(__dirname, 'assets', 'cori-favicon.png');
  
  try {
    await sharp(svgPath)
      .resize(32, 32)
      .png()
      .toFile(pngPath);
    
    console.log('SVG converted to PNG successfully');
    return pngPath;
  } catch (error) {
    console.error('Error converting SVG to PNG:', error);
    throw error;
  }
}

// Check if png-to-ico is installed
try {
  require.resolve('png-to-ico');
} catch (e) {
  console.log('Installing png-to-ico package...');
  execSync('npm install png-to-ico', { stdio: 'inherit' });
}

const pngToIco = require('png-to-ico');

// Convert PNG to ICO
async function convertPngToIco(pngPath) {
  const icoPath = path.join(__dirname, 'assets', 'favicon.ico');
  
  try {
    const buffer = await pngToIco([pngPath]);
    fs.writeFileSync(icoPath, buffer);
    
    // Also copy to the public folder
    const publicIcoPath = path.join(__dirname, 'public', 'favicon.ico');
    fs.copyFileSync(icoPath, publicIcoPath);
    
    console.log('PNG converted to ICO successfully');
    console.log(`Favicon saved to: ${icoPath}`);
    console.log(`Favicon copied to: ${publicIcoPath}`);
    
    // Clean up the temporary PNG
    fs.unlinkSync(pngPath);
    console.log('Temporary PNG file removed');
  } catch (error) {
    console.error('Error converting PNG to ICO:', error);
    throw error;
  }
}

// Run the conversion
async function main() {
  try {
    const pngPath = await convertSvgToPng();
    await convertPngToIco(pngPath);
    console.log('Favicon conversion completed successfully');
  } catch (error) {
    console.error('Favicon conversion failed:', error);
  }
}

main();
