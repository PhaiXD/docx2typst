import { readDocxFile } from './src/reader.ts';
import { extractDocument } from './src/extractor/documentExtractor.ts';
async function test() {
  const archive = await readDocxFile('lorem_test_docx2typst.docx');
  const doc = extractDocument(archive);
  let count = 0;
  doc.items.forEach((item, i) => {
    if (item.type === 'paragraph' && item.paragraph.text.includes('Praesent auctor')) {
      count = 5;
    }
    if (count > 0 && item.type === 'paragraph') {
      console.log('PARA text:', item.paragraph.text.substring(0, 30));
      console.log('PARA style:', item.paragraph.style);
      count--;
    }
  });
}
test();
