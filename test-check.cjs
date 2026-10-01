const fs = require('fs');
const svg = fs.readFileSync('test-2.svg', 'utf8');
const lines = svg.split('\n');
let pathCount = 0;
for (const l of lines) {
  if (l.includes('<path class="typst-shape"')) pathCount++;
}
console.log('Total paths on page 2:', pathCount);
