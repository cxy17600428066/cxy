const fs = require('fs');
const path = require('path');

const workspace = process.cwd();
const outDir = path.join(workspace, '.docx-template-build');

fs.rmSync(outDir, { recursive: true, force: true });
fs.mkdirSync(path.join(outDir, '_rels'), { recursive: true });
fs.mkdirSync(path.join(outDir, 'word', '_rels'), { recursive: true });

const columnWidths = [900, 1100, 1000, 900, 1000, 1100, 2000];

function esc(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function paragraph(text, options = {}) {
  const align = options.align || 'center';
  const size = options.size || 24;
  const bold = options.bold ? '<w:b/>' : '';
  const spacingAfter = options.spacingAfter
    ? `<w:spacing w:after="${options.spacingAfter}"/>`
    : '';

  if (!text) {
    return '<w:p/>';
  }

  return `
    <w:p>
      <w:pPr>
        <w:jc w:val="${align}"/>
        ${spacingAfter}
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="SimSun" w:hAnsi="SimSun" w:eastAsia="SimSun"/>
          ${bold}
          <w:sz w:val="${size}"/>
          <w:szCs w:val="${size}"/>
        </w:rPr>
        <w:t xml:space="preserve">${esc(text)}</w:t>
      </w:r>
    </w:p>
  `;
}

function cell(width, text, options = {}) {
  const gridSpan = options.gridSpan
    ? `<w:gridSpan w:val="${options.gridSpan}"/>`
    : '';
  const vMerge =
    options.vMerge === 'restart'
      ? '<w:vMerge w:val="restart"/>'
      : options.vMerge === 'continue'
        ? '<w:vMerge/>'
        : '';
  const vAlign = `<w:vAlign w:val="${options.vAlign || 'center'}"/>`;

  return `
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="${width}" w:type="dxa"/>
        ${gridSpan}
        ${vMerge}
        ${vAlign}
      </w:tcPr>
      ${paragraph(text)}
    </w:tc>
  `;
}

function row(height, cells) {
  return `
    <w:tr>
      <w:trPr>
        <w:trHeight w:val="${height}" w:hRule="atLeast"/>
      </w:trPr>
      ${cells.join('')}
    </w:tr>
  `;
}

const totalWidth = columnWidths.reduce((sum, n) => sum + n, 0);

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${paragraph('从业人员基本情况登记表', { size: 36, bold: true, spacingAfter: 220 })}
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="${totalWidth}" w:type="dxa"/>
        <w:jc w:val="center"/>
        <w:tblLayout w:type="fixed"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="8" w:space="0" w:color="auto"/>
          <w:left w:val="single" w:sz="8" w:space="0" w:color="auto"/>
          <w:bottom w:val="single" w:sz="8" w:space="0" w:color="auto"/>
          <w:right w:val="single" w:sz="8" w:space="0" w:color="auto"/>
          <w:insideH w:val="single" w:sz="8" w:space="0" w:color="auto"/>
          <w:insideV w:val="single" w:sz="8" w:space="0" w:color="auto"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="${columnWidths[0]}"/>
        <w:gridCol w:w="${columnWidths[1]}"/>
        <w:gridCol w:w="${columnWidths[2]}"/>
        <w:gridCol w:w="${columnWidths[3]}"/>
        <w:gridCol w:w="${columnWidths[4]}"/>
        <w:gridCol w:w="${columnWidths[5]}"/>
        <w:gridCol w:w="${columnWidths[6]}"/>
      </w:tblGrid>
      ${row(650, [
        cell(columnWidths[0], '姓  名'),
        cell(columnWidths[1], ''),
        cell(columnWidths[2], '性  别'),
        cell(columnWidths[3], ''),
        cell(columnWidths[4], '年  龄'),
        cell(columnWidths[5], ''),
        cell(columnWidths[6], '')
      ])}
      ${row(760, [
        cell(columnWidths[0], '职务'),
        cell(columnWidths[1], ''),
        cell(columnWidths[2] + columnWidths[3], '工作岗位', { gridSpan: 2 }),
        cell(columnWidths[4] + columnWidths[5], '', { gridSpan: 2 }),
        cell(columnWidths[6], '照片', { vMerge: 'restart' })
      ])}
      ${row(700, [
        cell(columnWidths[0], '籍贯'),
        cell(
          columnWidths[1] + columnWidths[2] + columnWidths[3] + columnWidths[4] + columnWidths[5],
          '',
          { gridSpan: 5 }
        ),
        cell(columnWidths[6], '', { vMerge: 'continue' })
      ])}
      ${row(700, [
        cell(columnWidths[0], '住址'),
        cell(
          columnWidths[1] + columnWidths[2] + columnWidths[3] + columnWidths[4] + columnWidths[5],
          '',
          { gridSpan: 5 }
        ),
        cell(columnWidths[6], '', { vMerge: 'continue' })
      ])}
      ${row(650, [
        cell(columnWidths[0] + columnWidths[1], '联系电话', { gridSpan: 2 }),
        cell(columnWidths[2] + columnWidths[3], '', { gridSpan: 2 }),
        cell(columnWidths[4], '身份证号'),
        cell(columnWidths[5] + columnWidths[6], '', { gridSpan: 2 })
      ])}
      ${row(3800, [
        cell(totalWidth, '', { gridSpan: 7 })
      ])}
    </w:tbl>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
`;

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="SimSun" w:hAnsi="SimSun" w:eastAsia="SimSun"/>
        <w:sz w:val="24"/>
        <w:szCs w:val="24"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
</w:styles>
`;

const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
`;

const packageRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
`;

const documentRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
`;

fs.writeFileSync(path.join(outDir, '[Content_Types].xml'), contentTypesXml, 'utf8');
fs.writeFileSync(path.join(outDir, '_rels', '.rels'), packageRelsXml, 'utf8');
fs.writeFileSync(path.join(outDir, 'word', 'document.xml'), documentXml, 'utf8');
fs.writeFileSync(path.join(outDir, 'word', 'styles.xml'), stylesXml, 'utf8');
fs.writeFileSync(path.join(outDir, 'word', '_rels', 'document.xml.rels'), documentRelsXml, 'utf8');

console.log(outDir);
