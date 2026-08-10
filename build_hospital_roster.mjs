import fs from 'node:fs/promises';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const workbook = Workbook.create();
const sheet = workbook.worksheets.add('8月值班表');
console.log('checkpoint: workbook');
sheet.showGridLines = false;

sheet.mergeCells('A1:I1');
sheet.getRange('A1').values = [['乐平镇卫生院8月24小时值班表']];
sheet.mergeCells('A2:I2');
sheet.getRange('A2').values = [['值班时段：00:00—24:00    填表说明：浅黄色单元格为可填写区域']];
sheet.getRange('A3:I3').values = [[
  '日期', '星期', '值班医生', '医生联系电话', '值班药房', '药房联系电话', '值班护士', '护士联系电话', '备注'
]];

const weekdays = ['星期六','星期日','星期一','星期二','星期三','星期四','星期五'];
const rows = Array.from({ length: 31 }, (_, i) => [
  `8月${i + 1}日`, weekdays[i % 7], '', '', '', '', '', '', ''
]);
sheet.getRange('A4:I34').values = rows;
console.log('checkpoint: values');

const font = { name: 'Microsoft YaHei', size: 10, color: '#1F2937' };
sheet.getRange('A1:I34').format.font = font;
sheet.getRange('A1:I34').format.verticalAlignment = 'center';
sheet.getRange('A1:I34').format.borders = { preset: 'all', style: 'thin', color: '#9CA3AF' };

sheet.getRange('A1:I1').format = {
  fill: '#0F6B56',
  font: { name: 'Microsoft YaHei', size: 20, bold: true, color: '#FFFFFF' },
  horizontalAlignment: 'center',
  verticalAlignment: 'center',
  borders: { preset: 'outside', style: 'medium', color: '#0F6B56' }
};
sheet.getRange('A2:I2').format = {
  fill: '#E7F3EF',
  font: { name: 'Microsoft YaHei', size: 10, color: '#245B4D' },
  horizontalAlignment: 'left',
  verticalAlignment: 'center',
  borders: { preset: 'outside', style: 'thin', color: '#9CA3AF' }
};
sheet.getRange('A3:I3').format = {
  fill: '#2D8C73',
  font: { name: 'Microsoft YaHei', size: 10, bold: true, color: '#FFFFFF' },
  horizontalAlignment: 'center',
  verticalAlignment: 'center',
  wrapText: true,
  borders: { preset: 'all', style: 'thin', color: '#D1D5DB' }
};
sheet.getRange('A4:B34').format = {
  fill: '#F3F7F5',
  font,
  horizontalAlignment: 'center',
  verticalAlignment: 'center',
  borders: { preset: 'all', style: 'thin', color: '#D1D5DB' }
};
sheet.getRange('C4:I34').format = {
  fill: '#FFF7D6',
  font,
  horizontalAlignment: 'center',
  verticalAlignment: 'center',
  borders: { preset: 'all', style: 'thin', color: '#D1D5DB' }
};

// Mark weekends for quick visual scanning (2026 August begins on Saturday).
for (let r = 4; r <= 34; r++) {
  const weekday = rows[r - 4][1];
  if (weekday === '星期六' || weekday === '星期日') {
    sheet.getRange(`A${r}:B${r}`).format.fill = '#FCE8E6';
    sheet.getRange(`A${r}:B${r}`).format.font = { name: 'Microsoft YaHei', size: 10, bold: true, color: '#B42318' };
  }
}

sheet.getRange('D4:D34').format.numberFormat = '@';
sheet.getRange('F4:F34').format.numberFormat = '@';
sheet.getRange('H4:H34').format.numberFormat = '@';

const widths = { A: 11, B: 10, C: 13, D: 17, E: 14, F: 17, G: 13, H: 17, I: 14 };
for (const [col, width] of Object.entries(widths)) sheet.getRange(`${col}1:${col}34`).format.columnWidth = width;
sheet.getRange('A1:I1').format.rowHeight = 38;
sheet.getRange('A2:I2').format.rowHeight = 24;
sheet.getRange('A3:I3').format.rowHeight = 34;
sheet.getRange('4:34').format.rowHeight = 24;
sheet.freezePanes.freezeRows(3);
console.log('checkpoint: formats');

const outputDir = 'outputs/hospital_aug_roster';
await fs.mkdir(outputDir, { recursive: true });
console.log((await workbook.inspect({
  kind: 'table', range: '8月值班表!A1:I34', include: 'values,formulas', tableMaxRows: 8, tableMaxCols: 9, maxChars: 4000
})).ndjson);
console.log((await workbook.inspect({
  kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A', options: { useRegex: true, maxResults: 100 }, summary: 'final formula error scan'
})).ndjson);

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
console.log('checkpoint: export');
await xlsx.save(`${outputDir}/乐平镇卫生院8月24小时值班表.xlsx`);
console.log('checkpoint: saved');
