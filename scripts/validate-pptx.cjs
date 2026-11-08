const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const pptxPath = path.resolve(__dirname, '../public/colon-cancer-presentation.pptx');
console.log('Validating PowerPoint file:', pptxPath);

if (!fs.existsSync(pptxPath)) {
  console.error('FAIL: File does not exist!');
  process.exit(1);
}

const stats = fs.statSync(pptxPath);
console.log('File size:', (stats.size / 1024 / 1024).toFixed(2), 'MB (' + stats.size + ' bytes)');

// Check zip archive integrity using PowerShell Expand-Archive or tar
const tempDir = path.resolve(__dirname, '../temp_pptx_validate');
if (fs.existsSync(tempDir)) {
  fs.rmSync(tempDir, { recursive: true, force: true });
}
fs.mkdirSync(tempDir);

try {
  // Use .NET ZipFile to unzip
  execSync(`powershell -Command "Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::ExtractToDirectory('${pptxPath}', '${tempDir}')"`);
  console.log('ZIP integrity check: PASSED (Archive uncompressed cleanly)');

  // Check [Content_Types].xml
  const contentTypesPath = path.join(tempDir, '[Content_Types].xml');
  if (!fs.existsSync(contentTypesPath)) {
    throw new Error('Missing [Content_Types].xml');
  }
  console.log('[Content_Types].xml: PASSED');

  // Check presentation.xml
  const presPath = path.join(tempDir, 'ppt', 'presentation.xml');
  if (!fs.existsSync(presPath)) {
    throw new Error('Missing ppt/presentation.xml');
  }
  const presXml = fs.readFileSync(presPath, 'utf8');
  console.log('presentation.xml: PASSED');

  // Count slides in ppt/slides/
  const slidesDir = path.join(tempDir, 'ppt', 'slides');
  const slideFiles = fs.readdirSync(slidesDir).filter(f => f.match(/^slide\d+\.xml$/));
  console.log('Total Slides Found:', slideFiles.length);

  if (slideFiles.length !== 46) {
    throw new Error(`Expected exactly 46 slides, but found ${slideFiles.length}`);
  }
  console.log('Slide Count Verification: PASSED (Exactly 46 slides)');

  // Verify slide order & sample titles from slide XMLs
  console.log('\n--- Sampling Key Slides ---');
  const samples = [1, 2, 3, 4, 5, 10, 20, 30, 40, 46];
  for (const sNum of samples) {
    const sPath = path.join(slidesDir, `slide${sNum}.xml`);
    if (!fs.existsSync(sPath)) {
      throw new Error(`Slide ${sNum} is missing!`);
    }
    const sXml = fs.readFileSync(sPath, 'utf8');
    const textMatches = sXml.match(/<a:t>([^<]+)<\/a:t>/g) || [];
    const extracted = textMatches.map(m => m.replace(/<\/?a:t>/g, '')).slice(0, 4).join(' | ');
    console.log(`Slide ${sNum}: ${extracted.substring(0, 80)}...`);
  }

  // Check media files
  const mediaDir = path.join(tempDir, 'ppt', 'media');
  if (fs.existsSync(mediaDir)) {
    const mediaFiles = fs.readdirSync(mediaDir);
    console.log('\nEmbedded Media Files Found:', mediaFiles.length);
  }

  console.log('\n=======================================');
  console.log('ALL PPTX VALIDATION CHECKS PASSED (100%)');
  console.log('=======================================');
} finally {
  if (fs.existsSync(tempDir)) {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
}
