const fs = require('fs');
const svg = fs.readFileSync('test-2.svg', 'utf8');
const lines = svg.split('\n');
for (const l of lines) {
  if (l.includes('<path class="typst-shape"')) console.log(l.trim());
}
