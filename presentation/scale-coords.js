const fs = require('fs');

// Read the file
let content = fs.readFileSync('create-presentation.js', 'utf8');

// Patterns to match and replace
// Match patterns like "x: 0.5," or "y: 1.0," or "w: 4.5," or "h: 0.8,"
// and wrap the number with s()

// Replace coordinate values (x, y, w, h) with scaled versions
// But skip values that are already wrapped in s()
content = content.replace(/\b([xywh]):\s*(\d+(?:\.\d+)?)/g, (match, coord, value) => {
  // Check if this value is already inside s()
  return `${coord}: s(${value})`;
});

// Replace line width values
content = content.replace(/width:\s*(\d+(?:\.\d+)?)/g, (match, value) => {
  return `width: s(${value})`;
});

// Write the modified content
fs.writeFileSync('create-presentation-scaled.js', content, 'utf8');

console.log('Scaling complete! New file: create-presentation-scaled.js');
