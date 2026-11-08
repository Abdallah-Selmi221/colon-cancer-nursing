const fs = require('fs');
const path = require('path');
const { createPresentation } = require('./pptxGenerator.cjs');

console.log('Generating PowerPoint presentation (.pptx)...');
const pres = createPresentation();

const outPath = path.resolve(__dirname, '../public/colon-cancer-presentation.pptx');

pres.writeFile({ fileName: outPath })
  .then(() => {
    const stats = fs.statSync(outPath);
    console.log('SUCCESS: PowerPoint file generated successfully!');
    console.log('File path:', outPath);
    console.log('File size:', (stats.size / 1024 / 1024).toFixed(2), 'MB (' + stats.size + ' bytes)');
  })
  .catch((err) => {
    console.error('ERROR generating PowerPoint:', err);
    process.exit(1);
  });
