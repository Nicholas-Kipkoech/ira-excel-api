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

export class IRABusinessForce {
  constructor() {}

  static async getBusinessForcePrems(req, res) {
    let connection;
    try {
      const { fromDate, toDate } = req.query;

      connection = (await pool).getConnection();
      if (connection) {
        console.log("Database connected...");
      }
      let query = `/* Formatted on 10/4/2024 11:40:16 AM (QP5 v5.336) */
  SELECT a.pr_org_code,
         a.pr_mc_code,
         INITCAP (
             CASE WHEN a.pr_sc_code IN ('0804') THEN 'PSV' ELSE pr_mc_name END)
             class,
         CASE
             WHEN a.pr_mc_code IN ('03',
                                   '04',
                                   '09',
                                   '11')
             THEN
                 1
             WHEN a.pr_mc_code IN ('070', '080')
             THEN
                 CASE WHEN a.pr_pl_no LIKE '%TP%' THEN 2 ELSE 1 END
             ELSE
                 CASE
                     WHEN pr_sc_code IN ('010',
                                         '020',
                                         '050',
                                         '051',
                                         '060',
                                         '061',
                                         '064',
                                         '100',
                                         '101',
                                         '120',
                                         '127',
                                         '128')
                     THEN
                         1
                     ELSE
                         2
                 END
         END
             pr_order,
         CASE
             WHEN a.pr_mc_code IN ('03',
                                   '04',
                                   '09',
                                   '11')
             THEN
                 a.pr_mc_name
             WHEN a.pr_mc_code IN ('070', '080')
             THEN
                 CASE
                     WHEN a.pr_pl_no LIKE '%TP%' THEN 'Third Party Only'
                     ELSE 'Comprehensive'
                 END
             ELSE
                 CASE
                     WHEN pr_sc_code IN ('010',
                                         '020',
                                         '050',
                                         '051',
                                         '060',
                                         '061',
                                         '064',
                                         '100',
                                         '101')
                     THEN
                         a.pr_sc_name
                     WHEN pr_sc_code IN ('120', '127', '128')
                     THEN
                         'Bonds'
                     ELSE
                         CASE
                             WHEN a.pr_mc_code IN ('10') THEN 'Burglary Others'
                             ELSE 'Others'
                         END
                 END
         END
             sub_class,
         COUNT (CASE WHEN pr_int_end_code = '000' THEN pr_end_no ELSE NULL END)
             policies_nb,
         SUM (
             CASE
                 WHEN pr_int_end_code = '000'
                 THEN
                     (NVL (a.pr_fc_si, 0) * a.pr_cur_rate)
                 ELSE
                     0
             END)
             si_nb,
         SUM (
             CASE
                 WHEN pr_int_end_code = '000'
                 THEN
                     ROUND (
                         (  (NVL (a.pr_fc_prem, 0) * a.pr_cur_rate)
                          + (NVL (a.pr_fc_eartquake, 0) * a.pr_cur_rate)
                          + (NVL (a.pr_fc_political, 0) * a.pr_cur_rate)),
                         0)
                 ELSE
                     0
             END)
             prem_nb,
         COUNT (
             CASE WHEN pr_int_end_code != '000' THEN pr_end_no ELSE NULL END)
             policies_rn,
         SUM (
             CASE
                 WHEN pr_int_end_code != '000'
                 THEN
                     ROUND (
                         NVL (
                             CASE
                                 WHEN a.pr_net_effect IN ('Credit')
                                 THEN
                                     (  (NVL (a.pr_fc_si, 0) * a.pr_cur_rate)
                                      * -1)
                                 ELSE
                                     (NVL (a.pr_fc_si, 0) * a.pr_cur_rate)
                             END,
                             0),
                         0)
                 ELSE
                     0
             END)
             si_rn,
         SUM (
             CASE
                 WHEN pr_int_end_code != '000'
                 THEN
                     ROUND (
                         NVL (
                             CASE
                                 WHEN a.pr_net_effect IN ('Credit')
                                 THEN
                                     (  (  (  NVL (a.pr_fc_prem, 0)
                                            * a.pr_cur_rate)
                                         + (  NVL (a.pr_fc_eartquake, 0)
                                            * a.pr_cur_rate)
                                         + (  NVL (a.pr_fc_political, 0)
                                            * a.pr_cur_rate))
                                      * -1)
                                 ELSE
                                     (  (NVL (a.pr_fc_prem, 0) * a.pr_cur_rate)
                                      + (  NVL (a.pr_fc_eartquake, 0)
                                         * a.pr_cur_rate)
                                      + (  NVL (a.pr_fc_political, 0)
                                         * a.pr_cur_rate))
                             END,
                             0),
                         0)
                 ELSE
                     0
             END)
             prem_rn
    FROM uw_premium_register a
   WHERE     pr_org_code = :p_org_code
         AND TRUNC (PR_GL_DATE) BETWEEN TRUNC ( :pr_fm_dt)
                                    AND TRUNC ( :pr_to_dt)
GROUP BY a.pr_org_code,
         a.pr_mc_code,
         INITCAP (
             CASE
                 WHEN a.pr_sc_code IN ('0804') THEN 'PSV'
                 ELSE pr_mc_name
             END),
         CASE
             WHEN a.pr_mc_code IN ('03',
                                   '04',
                                   '09',
                                   '11')
             THEN
                 1
             WHEN a.pr_mc_code IN ('070', '080')
             THEN
                 CASE WHEN a.pr_pl_no LIKE '%TP%' THEN 2 ELSE 1 END
             ELSE
                 CASE
                     WHEN pr_sc_code IN ('010',
                                         '020',
                                         '050',
                                         '051',
                                         '060',
                                         '061',
                                         '064',
                                         '100',
                                         '101',
                                         '120',
                                         '127',
                                         '128')
                     THEN
                         1
                     ELSE
                         2
                 END
         END,
         CASE
             WHEN a.pr_mc_code IN ('03',
                                   '04',
                                   '09',
                                   '11')
             THEN
                 a.pr_mc_name
             WHEN a.pr_mc_code IN ('070', '080')
             THEN
                 CASE
                     WHEN a.pr_pl_no LIKE '%TP%' THEN 'Third Party Only'
                     ELSE 'Comprehensive'
                 END
             ELSE
                 CASE
                     WHEN pr_sc_code IN ('010',
                                         '020',
                                         '050',
                                         '051',
                                         '060',
                                         '061',
                                         '064',
                                         '100',
                                         '101')
                     THEN
                         a.pr_sc_name
                     WHEN pr_sc_code IN ('120', '127', '128')
                     THEN
                         'Bonds'
                     ELSE
                         CASE
                             WHEN a.pr_mc_code IN ('10')
                             THEN
                                 'Burglary Others'
                             ELSE
                                 'Others'
                         END
                 END
         END
ORDER BY a.pr_org_code, a.pr_mc_code, pr_order`;
      const results = (await connection).execute(query, {
        p_org_code: "50",
        pr_fm_dt: new Date(fromDate),
        pr_to_dt: new Date(toDate),
      });
      const finalResults = (await results).rows?.map((row, index) => {
        return {
          class: row[2],
          subClass: row[4],
          policiesNB: row[5],
          sumInsuredNB: row[6],
          premiumNB: row[7],
          policiesRN: row[8],
          sumInsuredRN: row[9],
          premiumRN: row[10],
          policiesBF: row[8] + row[5],
          sumInsuredBF: row[9] + row[6],
          premiumBF: row[10] + row[7],
        };
      });
      const updateWorkbook = async (workbook) => {
        const worksheet = workbook.getWorksheet("59-11B");

        Object.entries(classSubclassRowMapper).forEach(
          ([classSubKey, targetRow]) => {
            const [classKey, subClassKey] = classSubKey.split("|");

            const filteredResults = finalResults.filter(
              (item) => item.class === classKey && item.subClass === subClassKey
            );
            if (filteredResults.length > 0) {
              filteredResults.forEach((dataItem) => {
                for (const [field, column] of Object.entries(cellMapper2)) {
                  const cell = worksheet.getCell(`${column}${targetRow + 1}`);
                  cell.value = dataItem[field];
                  console.log(
                    `${field} (${column}${targetRow}): ${dataItem[field]}`
                  );
                }
              });
            }
          }
        );
      };

      // Use writeFileSafely to handle file locking and write operation
      await writeFileSafely(filePath, updateWorkbook);

      // Send a success response
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
