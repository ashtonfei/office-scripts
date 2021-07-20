type ExcelValues = string | number | boolean;

class Utils{
  constructor(){}
  excelDateValueToDate(dateValue: number): Date{
    return new Date(Math.round((dateValue - 25569) * 86400 * 1000));
  }
}

function main(workbook: ExcelScript.Workbook): SalesInterface[]
{
  const app = new App(workbook)
  return app.getSalesYesterday()
}

class App{
  wb: ExcelScript.Workbook;
  utils: Utils;
  constructor(wb: ExcelScript.Workbook){
    this.wb = wb
    this.utils = new Utils()
  }

  getValues(name: string): ExcelValues[][]{
    const ws:ExcelScript.Worksheet = this.wb.getWorksheet(name)
    return ws.getUsedRange().getValues()
  }

  getHeaderIndexes(values: ExcelValues[][], headerRowIndex: number = 0) : {}{
    const indexes: {} = {}
    const headers = values[headerRowIndex]
    headers.forEach((header, index) => indexes[header.toString().trim()] = index)
    return indexes
  }

  getSalesYesterday(): SalesInterface[]{
    const today: Date = new Date()
    const sales: SalesInterface[] = []

    const values: ExcelValues[][] = this.getValues("Sales")
    const indexes: {} = this.getHeaderIndexes(values)

    const salesIndexes: SalesIndexInteface = {
      productIndex: indexes['Product'],
      locationIndex: indexes['Store Location'],
      ownerIndex: indexes['Owner'],
      amountIndex: indexes['Amount'],
      priceIndex: indexes['Price'],
      dateIndex: indexes['Date'],
    }

    values.slice(1).forEach(v => {
      const salesDate: Date = this.utils.excelDateValueToDate(v[salesIndexes.dateIndex] as number)
      if (salesDate < today){
        const salesItem: SalesInterface = {
          product: v[salesIndexes.productIndex] as string,
          location: v[salesIndexes.locationIndex] as string,
          owner: v[salesIndexes.ownerIndex] as string,
          amount: v[salesIndexes.amountIndex] as number,
          price: v[salesIndexes.priceIndex] as number,
          total: (v[salesIndexes.priceIndex] as number) * (v[salesIndexes.amountIndex] as number),
          date: salesDate.toLocaleDateString()
        }
        sales.push(salesItem)
      }
    })
    console.log(sales)
    return sales
  }
}

interface SalesInterface {
  product: string;
  location: string;
  owner: string;
  amount: number;
  price: number;
  total: number;
  date: string;
}

interface SalesIndexInteface {
  productIndex: number;
  locationIndex: number;
  ownerIndex: number;
  amountIndex: number;
  priceIndex: number;
  dateIndex: number;
}
