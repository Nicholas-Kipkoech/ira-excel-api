import ExcelJs from "exceljs";

//the file path
const filePath = "TEST.xlsx";

//initiate the workbook or the excel package
const workbook = new ExcelJs.Workbook();

// read the file
workbook.xlsx
  .readFile(filePath)
  .then(() => {
    const worksheet = workbook.getWorksheet("Details");
    const cell = worksheet.getCell("F10");
    console.log(cell.value);
  })
  .then(() => {
    console.log("Data added successfully!");
  })
  .catch((err) => {
    console.error("Error modifying the Excel file:", err);
  });
