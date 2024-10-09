import pool from "../config/database.js";
import formatOracleData from "../utils/helpers.js";
import ExcelJs from "exceljs";
import {
  cellMapper,
  cellMapper2,
  classSubclassRowMapper,
} from "./IRA-class-prem-mapper.js";
import { writeFileSafely } from "./excel-service/excel-helper.js";

//the file path
const filePath = "IRA_excel.xlsx";

export class IRAReinsuranceBalancesService {
  constructor() {}

  static async getBalanceReport(req, res) {
    let connection;
    try {
      const { toDate } = req.query;
      connection = (await pool).getConnection();
      if (connection) {
        console.log("Database connected...");
      }
      let query = `/* Formatted on 10/1/2024 10:13:40 AM (QP5 v5.336) */
  SELECT trn_org_code,
         trn_mgl_code,
         trn_sgl_code,
         INITCAP (
             pkg_gl.get_sub_acnt_name (trn_org_code,
                                       trn_mgl_code,
                                       trn_sgl_code))    customer,
         SUM (
             CASE
                 WHEN trn_drcr_flag = 'D'
                 THEN
                     NVL (trn_doc_fc_amt, 0) * trn_cur_rate
                 WHEN trn_drcr_flag = 'C'
                 THEN
                     -NVL (trn_doc_fc_amt, 0) * trn_cur_rate
                 ELSE
                     0
             END)                                        AS balance
    FROM gl_transactions
   WHERE     trn_org_code = :p_org_code
         --and trn_mgl_code='LA048'
         AND trn_doc_gl_dt <= :p_asatdate
         AND trn_sgl_code IS NOT NULL
         AND trn_mgl_code = 'LA049'
GROUP BY trn_org_code,
         trn_mgl_code,
         trn_sgl_code,
         pkg_gl.get_sub_acnt_name (trn_org_code, trn_mgl_code, trn_sgl_code)
--pkg_gl.get_sub_ledger_open_bal(trn_org_code,trn_mgl_code,trn_sgl_code,:p_asatdate)
ORDER BY trn_sgl_code ASC`;
      const results = (await connection).execute(query, {
        p_org_code: "50",
        p_asatdate: new Date(toDate),
      });
      const finalResults = formatOracleData(await results);

      const updateWorkbook = (workbook) => {
        const worksheet = workbook.getWorksheet("41-1D (c)");

        const dropdownRangeStart = 10; // Start row of the dropdown list
        const dropdownRangeEnd = 65; // End row of the dropdown list
        const dropdownColumn = "J"; // The column where the dropdown values are stored

        // Iterate through your data
        finalResults.forEach((dataItem) => {
          let matchedRow = null;
          // Loop through the dropdown list to find a matching customer
          for (let row = dropdownRangeStart; row <= dropdownRangeEnd; row++) {
            const dropdownValue = worksheet.getCell(
              `${dropdownColumn}${row}`
            ).value;
            console.log(
              "drop down value",
              dropdownValue,
              "dataItem",
              dataItem.CUSTOMER.toUpperCase()
            );

            // If the dropdown value matches the CUSTOMER in the data
            if (dropdownValue === dataItem.CUSTOMER.toUpperCase()) {
              matchedRow = row; // Save the matching row number
              break; // Exit the loop once a match is found
            }
          }

          // If a matching dropdown value (customer) is found, fill the corresponding row
          if (matchedRow) {
            worksheet.getCell(`D${matchedRow}`).value = dataItem.BALANCE; // Assuming column D is for BALANCE

            console.log(
              `Filled row ${matchedRow} for customer: ${dataItem.CUSTOMER}`
            );
          } else {
            console.log(`Customer not found in dropdown: ${dataItem.CUSTOMER}`);
          }
        });
      };
      await writeFileSafely(updateWorkbook);

      return res.status(200).json({
        message: "Data written successfully",
        results: finalResults,
      });
    } catch (error) {
      console.error("error getting the premiums", error);
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
