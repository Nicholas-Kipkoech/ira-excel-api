import lockfile from "proper-lockfile";
import ExcelJS from "exceljs";
export async function writeFileSafely(filePath, updateFunction) {
  try {
    // Lock the Excel file before writing
    const release = await lockfile.lock(filePath);

    // Load the existing workbook
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    // Call the update function to modify the workbook
    await updateFunction(workbook);

    // Save the updated workbook
    await workbook.xlsx.writeFile(filePath);

    // Release the lock after the file is written
    await release();
  } catch (error) {
    console.error("Error writing to Excel file:", error);
    throw new Error("Failed to update Excel file");
  }
}
