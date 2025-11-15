const fs = require('fs');
const { spawn } = require('child_process');

// Since we can't easily convert PPTX to PNG in this environment,
// let's create a simpler solution using the HTML slides

const slides = [];
for (let i = 1; i <= 17; i++) {
  slides.push(`slide${i}.html`);
}

console.log('Creating slide previews...');
console.log(`Total slides to process: ${slides.length}`);

// Create a manifest file listing all slides
const manifest = {
  title: "Mallacoota Beauty Pivot Pitch Deck",
  resolution: "1920x1080",
  slides: slides.map((file, index) => ({
    number: index + 1,
    filename: file,
    outputPng: `slide-${String(index + 1).padStart(2, '0')}.png`
  }))
};

fs.writeFileSync('slides-png/manifest.json', JSON.stringify(manifest, null, 2));

console.log('Created manifest.json in slides-png/');
console.log('\nTo convert HTML slides to PNG, you can use tools like:');
console.log('1. wkhtmltoimage: wkhtmltoimage --width 1920 --height 1080 slide1.html slide-01.png');
console.log('2. Chrome headless: chrome --headless --screenshot --window-size=1920,1080 slide1.html');
console.log('3. Online converters or desktop tools like Adobe Acrobat, PowerPoint export');

// Create a simple HTML index file to view all slides
const indexHtml = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Mallacoota Beauty Pivot - All Slides</title>
  <style>
    body {
      font-family: Arial, sans-serif;
      margin: 0;
      padding: 20px;
      background: #f0f0f0;
    }
    h1 {
      color: #003366;
      text-align: center;
    }
    .slides-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(400px, 1fr));
      gap: 20px;
      margin-top: 20px;
    }
    .slide-preview {
      background: white;
      padding: 15px;
      border-radius: 8px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    .slide-preview h3 {
      margin-top: 0;
      color: #003366;
    }
    .slide-preview iframe {
      width: 100%;
      height: 225px;
      border: 1px solid #ddd;
      border-radius: 4px;
    }
    .slide-link {
      display: inline-block;
      margin-top: 10px;
      color: #D4AF37;
      text-decoration: none;
      font-weight: bold;
    }
    .slide-link:hover {
      text-decoration: underline;
    }
  </style>
</head>
<body>
  <h1>Mallacoota Beauty Pivot - Pitch Deck Preview</h1>
  <p style="text-align: center; color: #666;">Full HD 1920x1080px | 17 Slides</p>

  <div class="slides-grid">
${slides.map((file, i) => `
    <div class="slide-preview">
      <h3>Slide ${i + 1}</h3>
      <iframe src="${file}" scrolling="no"></iframe>
      <a href="${file}" target="_blank" class="slide-link">Open Full Size →</a>
    </div>
`).join('')}
  </div>
</body>
</html>`;

fs.writeFileSync('slides-png/index.html', indexHtml);

console.log('\nCreated index.html for slide preview');
console.log('\nManifest file with slide information saved to slides-png/manifest.json');
