const fs = require('fs');
const path = require('path');

const sourcePath = path.join(__dirname, 'node_modules', 'pdfjs-dist', 'build', 'pdf.worker.min.mjs');
const destPath = path.join(__dirname, 'src', 'common', 'assets', 'pdf.worker.min.js');

console.log(`Copying worker from ${sourcePath} to ${destPath}...`);

if (fs.existsSync(sourcePath)) {
  try {
    const workerContent = fs.readFileSync(sourcePath);
    fs.writeFileSync(destPath, workerContent);
    console.log('Successfully updated pdf.worker.min.js');
  } catch (error) {
    console.error('Error copying file:', error);
    process.exit(1);
  }
} else {
  console.error('Source file not found. Ensure npm install has been run and pdfjs-dist is available.');
  process.exit(1);
}
