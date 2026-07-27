const fs = require('fs');
const file = 'c:/Users/Homes247/Desktop/office-suite/frontend/src/app/pages/sheet-editor/sheet-editor.component.ts';
let lines = fs.readFileSync(file, 'utf8').split(/\r?\n/);

let startIndex = -1;
let endIndex = -1;

for (let i = 0; i < lines.length; i++) {
  if (lines[i].includes('toggleMenu(') && lines[i].includes('allSheets') && lines[i].includes('title="All sheets"')) {
    startIndex = i;
  }
  if (startIndex !== -1 && lines[i].includes('</button>')) {
    endIndex = i;
    break;
  }
}

if (startIndex !== -1 && endIndex !== -1) {
  lines.splice(startIndex, endIndex - startIndex + 1);
  fs.writeFileSync(file, lines.join('\n'));
  console.log('Removed allSheets button lines ' + startIndex + ' to ' + endIndex);
} else {
  console.log('Could not find allSheets button');
}
