const xml = '<w:p><w:r><w:t>A</w:t></w:r><w:hyperlink r:id="rId1"><w:r><w:t>B</w:t></w:r><w:r><w:t>C</w:t></w:r></w:hyperlink><w:r><w:t>D</w:t></w:r></w:p>';

const modified = xml.replace(/<w:hyperlink[^>]*r:id="([^"]+)"[^>]*>(.*?)<\/w:hyperlink>/gs, (match, rId, content) => {
  return content.replace(/<w:r(?: [^>]+)?>/g, (rTag) => {
    return rTag + '<w:rPr><w:linkTarget w:val="' + rId + '"/></w:rPr>';
  });
});
console.log(modified);

const { XMLParser } = require('fast-xml-parser');
const parser = new XMLParser({ ignoreAttributes: false });
console.log(JSON.stringify(parser.parse(modified), null, 2));
