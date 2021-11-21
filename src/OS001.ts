type ExcelValues = string | number | boolean;

const SHEET_NAME = {
  SALES: "Sales",
  SALES_BY_REP: "Sales by Rep",
  SALES_BY_REGION: "Sales by Region",
};

function main(wb: ExcelScript.Workbook){
  const sheetSales = wb.getWorksheet(SHEET_NAME.SALES)
  const sheetSalesByRep = wb.getWorksheet(SHEET_NAME.SALES_BY_REP)
  const sheetSalesByRegion = wb.getWorksheet(SHEET_NAME.SALES_BY_REGION)

  const htmlTableSalesByRep = createHtmlTable(sheetSalesByRep)
  const htmlTableSalesByRegion = createHtmlTable(sheetSalesByRegion)

  const response = {
    success: true,
    message: "Success",
    htmlBody: `
      <div>Here is the sales report:</div>
      <div>${htmlTableSalesByRep}</div>
      <div>${htmlTableSalesByRegion}</div>
      <div>Thanks,<br>Ashton Fei</div>
    `
  }
  return response
}


function getDataFromSheet(sheet:ExcelScript.Worksheet): string[][]{
  return sheet.getUsedRange().getTexts()
}

function createHtmlTable(sheet: ExcelScript.Worksheet): string{
  const texts:string[][] = getDataFromSheet(sheet)
  let htmlTable = `<table>`
  
  texts.forEach((rowData, index) => {
    let trow = `<tr>`
    if (index === 0) {
      rowData.forEach(value => trow += `<th>${value}</th>`)
    }else{
      rowData.forEach(value => trow += `<td>${value}</td>`)
    }
    trow += `</tr>`
    htmlTable += trow
  })

  return `${htmlTable}</table>`
}