import pool from "../config/database.js";
import ExcelJs from "exceljs";
import {
  cellMapper6,
  classSubclassRowMapper5,
} from "./IRA-class-prem-mapper.js";
import formatOracleData from "../utils/helpers.js";
import { writeFileSafely } from "./excel-service/excel-helper.js";

//the file path
const filePath = "IRA_excel.xlsx";

export class IRAPremiumsByCounty {
  constructor() {}

  static async getPremiumsByCounty(req, res) {
    let connection;
    try {
      const { fromDate, toDate } = req.query;
      connection = (await pool).getConnection();
      if (connection) {
        console.log("Database connected...");
      }
      let query = `
  SELECT NVL (PKG_SYSTEM_ADMIN.GET_COLUMN_VALUE_TWO ('AD_SYSTEM_CODES',
                                                     'SYS_GROUPING',
                                                     'SYS_TYPE',
                                                     'SYS_NAME',
                                                     'UW_PHYS_LOC',
                                                     pl_phys_loc),
              'Un-Attached')                                     phys_loc_grp,
         NVL (SUM (NVL (a.pl_fc_prem, 0) * b.pl_cur_rate), 0)    premium
    FROM uh_policy_risks a, uh_policy b
   WHERE     a.pl_org_code = b.pl_org_code
         AND a.pl_pl_index = b.pl_index
         AND a.pl_end_index = b.pl_end_index
         AND b.pl_org_code(+) = :p_org_code
         AND TRUNC (b.pl_gl_date) BETWEEN TRUNC ( :p_fm_dt)
                                      AND TRUNC ( :p_to_dt)
GROUP BY PKG_SYSTEM_ADMIN.GET_COLUMN_VALUE_TWO ('AD_SYSTEM_CODES',
                                                'SYS_GROUPING',
                                                'SYS_TYPE',
                                                'SYS_NAME',
                                                'UW_PHYS_LOC',
                                                pl_phys_loc)
ORDER BY 2`;
      const results = (await connection).execute(query, {
        p_org_code: "50",
        p_fm_dt: new Date(fromDate),
        p_to_dt: new Date(toDate),
      });
      const finalResults = formatOracleData(await results);
      const updateWorkbook = (workbook) => {
        const worksheet = workbook.getWorksheet("18-1F");
        const startRow = 10;
        let currentRow = startRow;
        finalResults.forEach((value, index) => {
          worksheet.getCell(`D${currentRow}`).value = value.PHYS_LOC_GRP;
          console.log(`D${currentRow}, E${currentRow}`);
          worksheet.getCell(`E${currentRow}`).value = value.PREMIUM;
          currentRow++;
        });
      };
      await writeFileSafely(updateWorkbook);

      // Send a success response
      return res.status(200).json({
        message: "Data written successfully",
        results: finalResults,
      });
    } catch (error) {
      console.error("error getting the commissions", error);
      return res.status(500).json(error);
    } finally {
      try {
        if (connection) {
          (await connection).close();
          console.info("Connection closed successfully");
        }
      } catch (error) {
        console.error(error);
      }
    }
  }
}
