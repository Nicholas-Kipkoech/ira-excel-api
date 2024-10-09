import lockfile from "proper-lockfile";
import ExcelJS from "exceljs";
import path from "path";
import os from "os";

const folderName = "IRA"; // Folder on desktop
const fileName = "IRA_file.xlsx"; // Excel file inside the folder

export async function writeFileSafely(updateFunction) {
  let release;
  try {
    // Lock the Excel file before writing

    const homeDir = os.homedir();

    const desktopDir = path.join(homeDir, "Desktop");
    const folderPath = path.join(desktopDir, folderName);
    const filePath = path.join(folderPath, fileName);
    // Load the existing workbook
    const workbook = new ExcelJS.Workbook();

    release = await lockfile.lock(filePath);

    await workbook.xlsx.readFile(filePath);

    // Call the update function to modify the workbook
    await updateFunction(workbook);

    const newFilePath = path.join(desktopDir, "updated_file_copy.xlsx");
    // Save the updated workbook
    await workbook.xlsx.writeFile(newFilePath);
    console.log(`File written successfully to: ${newFilePath}`);
    // Release the lock after the file is written
  } catch (error) {
    console.error("Error writing to Excel file:", error);
    throw new Error("Failed to update Excel file");
  } finally {
    if (release) {
      await release();
    }
  }
}
