const SHEET_NAMES = ["Sales by Region", "Sales by Rep"];

function main(wb: ExcelScript.Workbook): string {
  const sheets = SHEET_NAMES.map((name) => wb.getWorksheet(name));
  const htmlBody = createHtmlBody(sheets);
  return htmlBody;
}

function createHtmlBody(sheets: ExcelScript.Worksheet[]) {
  let html = `<table>`;
  const charts = sheets.map((sheet) => getChartFromSheet(sheet));
  charts.forEach((chart, index) => {
    const sheetName = SHEET_NAMES[index];
    if (chart.length) {
      let tr = `<tr><td><h3>${sheetName}</h3></td></tr><tr>`;
      chart.forEach((data) => {
        tr += `<td><img src="${data}"></td>`;
      });
      tr += `</tr>`;
      html += tr;
    }
  });
  html += `</table>`;
  return html;
}

function getChartFromSheet(sheet: ExcelScript.Worksheet) {
  const charts = sheet.getCharts();
  return charts.map((chart) => {
    return `data:image/png;base64,${chart.getImage()}`;
  });
}
