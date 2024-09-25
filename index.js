import ExcelJs from "exceljs";

//the file path
const filePath = "TEST.xlsx";

//initiate the workbook or the excel package
const workbook = new ExcelJs.Workbook();

// read the file
workbook.xlsx
  .readFile(filePath)
  .then(() => {
    const worksheet = workbook.getWorksheet("59-1B (a)");
    const row = worksheet.getRow("17");
    for (let i = 6; i <= 13; i++) {
      console.log(row.getCell(i).address, row.getCell(i).value);
    }
  })
  .then(() => {
    console.log("Data added successfully!");
  })
  .catch((err) => {
    console.error("Error modifying the Excel file:", err);
  });
