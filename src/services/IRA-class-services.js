import pool from "../config/database.js";
import formatOracleData from "../utils/helpers.js";
import ExcelJs from "exceljs";
import { cellMapper, classSubclassRowMapper } from "./IRA-class-prem-mapper.js";
import { writeFileSafely } from "./excel-service/excel-helper.js";

//the file path
const filePath = "IRA_excel.xlsx";

export class IRAPremClass {
  constructor() {}

  static async getPremiums(req, res) {
    let connection;
    try {
      const { fromDate, toDate } = req.query;
      connection = (await pool).getConnection();
      if (connection) {
        console.log("Database connected...");
      }
      let query = `/* Formatted on 10/4/2024 11:37:34 AM (QP5 v5.336) */
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
         SUM (
             CASE
                 WHEN pr_bus_type = '1000'
                 THEN
                     (ROUND (
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
                          0))
                 ELSE
                     0
             END)
             prem_direct,
         SUM (
             CASE
                 WHEN pr_bus_type IN ('1001', '2000', '2999')
                 --          AND pr_int_aent_code = '70'
                 THEN
                     (ROUND (
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
                          0))
                 ELSE
                     0
             END)
             prem_broker,
         SUM (
             CASE
                 WHEN pr_bus_type IN ('1002')
                 THEN
                     (ROUND (
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
                          0))
                 ELSE
                     0
             END)
             prem_agent,
         SUM (
             CASE
                 WHEN pr_bus_type = '3000'
                 THEN
                     (ROUND (
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
                          0))
                 ELSE
                     0
             END)
             prem_facin
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
      const finalResults = formatOracleData(await results);
      //initiate the workbook or the excel package
      const updateWorkbook = (workbook) => {
        const worksheet = workbook.getWorksheet("59-1B (a)");

        Object.entries(classSubclassRowMapper).forEach(
          ([classSubKey, targetRow]) => {
            const [classKey, subClassKey] = classSubKey.split("|");

            const filteredResults = finalResults.filter(
              (item) =>
                item.CLASS === classKey && item.SUB_CLASS === subClassKey
            );
            if (filteredResults.length > 0) {
              filteredResults.forEach((dataItem) => {
                for (const [field, column] of Object.entries(cellMapper)) {
                  const cell = worksheet.getCell(`${column}${targetRow}`);
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
      await writeFileSafely(updateWorkbook);

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
