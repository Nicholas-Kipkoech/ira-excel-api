import pool from "../config/database.js";
import ExcelJs from "exceljs";
import formatOracleData from "../utils/helpers.js";
import {
  cellMapper6,
  cellMapper7,
  classSubclassRowMapper6,
} from "./IRA-class-prem-mapper.js";
import { writeFileSafely } from "./excel-service/excel-helper.js";

//the file path
const filePath = "IRA_excel.xlsx";

export class IRAPremiumRegisterService {
  constructor() {}

  static async getPremiums(req, res) {
    let connection;
    try {
      const { fromDate, toDate } = req.query;
      connection = (await pool).getConnection();
      if (connection) {
        console.log("Database connected...");
      }
      let query = `
     /* Formatted on 9/30/2024 11:08:52 AM (QP5 v5.336) */
  SELECT pr_class,
         pr_sub_class,
         SUM (pr_lc_prem + +pr_lc_eartquake + pr_lc_political)    total_premiums
    FROM (  SELECT pr_org_code,
                   pr_pl_index,
                   pr_end_index,
                   pr_pl_no,
                   pr_end_no,
                   pr_issue_date,
                   pr_gl_date,
                   pr_fm_dt,
                   pr_to_dt,
                   pr_mc_code,
                   pr_mc_name                           pr_class,
                   pr_sc_code,
                   pr_sc_code                           pr_sc_code_i,
                   pr_sc_name                           pr_sub_class,
                   pr_pr_code,
                   pr_pr_code || ' - ' || pr_pr_name    pr_product,
                   pr_int_aent_code,
                   pr_int_ent_code,
                   pr_int_ent_name                      pr_intermediary,
                   pr_assr_aent_code,
                   pr_assr_ent_code,
                   pr_assr_ent_name                     pr_insured,
                   pr_os_code,
                   pr_os_name                           pl_os_name,
                   pr_int_end_code,
                   CASE
                       WHEN pr_int_end_code IN ('000') THEN 1
                       WHEN pr_int_end_code IN ('110') THEN 4
                       WHEN pr_net_effect IN ('Credit') THEN 3
                       ELSE 2
                   END                                  pr_end_order,
                   CASE
                       WHEN pr_int_end_code IN ('000') THEN 'New Business'
                       WHEN pr_int_end_code IN ('110') THEN 'Renewals'
                       WHEN pr_net_effect IN ('Credit') THEN 'Refunds'
                       ELSE 'Extras'
                   END                                  pr_end_type,
                   pr_cur_code,
                   pr_cur_rate,
                   pr_net_effect,
                   NVL (
                       DECODE ( :p_currency,
                               NULL, NVL ((NVL (pr_fc_si, 0) * pr_cur_rate), 0),
                               NVL (pr_fc_si, 0)),
                       0)                               pr_lc_si,
                   NVL (
                       CASE
                           WHEN pr_net_effect IN ('Credit')
                           THEN
                               NVL (
                                   (  (DECODE (
                                           :p_currency,
                                           NULL, NVL (
                                                     (  NVL (pr_fc_prem, 0)
                                                      * pr_cur_rate),
                                                     0),
                                           NVL (pr_fc_prem, 0)))
                                    * -1),
                                   0)
                           ELSE
                               NVL (
                                   DECODE (
                                       :p_currency,
                                       NULL, NVL (
                                                 (  NVL (pr_fc_prem, 0)
                                                  * pr_cur_rate),
                                                 0),
                                       NVL (pr_fc_prem, 0)),
                                   0)
                       END,
                       0)                               pr_lc_prem,
                   NVL (
                       CASE
                           WHEN pr_net_effect IN ('Credit')
                           THEN
                               NVL (
                                   (  (DECODE (
                                           :p_currency,
                                           NULL, NVL (
                                                     (  NVL (pr_fc_eartquake, 0)
                                                      * pr_cur_rate),
                                                     0),
                                           NVL (pr_fc_eartquake, 0)))
                                    * -1),
                                   0)
                           ELSE
                               NVL (
                                   DECODE (
                                       :p_currency,
                                       NULL, NVL (
                                                 (  NVL (pr_fc_eartquake, 0)
                                                  * pr_cur_rate),
                                                 0),
                                       NVL (pr_fc_eartquake, 0)),
                                   0)
                       END,
                       0)                               pr_lc_eartquake,
                   NVL (
                       CASE
                           WHEN pr_net_effect IN ('Credit')
                           THEN
                               NVL (
                                   (  (DECODE (
                                           :p_currency,
                                           NULL, NVL (
                                                     (  NVL (pr_fc_political, 0)
                                                      * pr_cur_rate),
                                                     0),
                                           NVL (pr_fc_political, 0)))
                                    * -1),
                                   0)
                           ELSE
                               NVL (
                                   DECODE (
                                       :p_currency,
                                       NULL, NVL (
                                                 (  NVL (pr_fc_political, 0)
                                                  * pr_cur_rate),
                                                 0),
                                       NVL (pr_fc_political, 0)),
                                   0)
                       END,
                       0)                               pr_lc_political,
                   NVL (
                       CASE
                           WHEN pr_net_effect IN ('Credit')
                           THEN
                               NVL (
                                   (  (DECODE (
                                           :p_currency,
                                           NULL, NVL (
                                                     (  NVL (pr_fc_broker_comm,
                                                             0)
                                                      * pr_cur_rate),
                                                     0),
                                           NVL (pr_fc_broker_comm, 0)))
                                    * -1),
                                   0)
                           ELSE
                               NVL (
                                   DECODE (
                                       :p_currency,
                                       NULL, NVL (
                                                 (  NVL (pr_fc_broker_comm, 0)
                                                  * pr_cur_rate),
                                                 0),
                                       NVL (pr_fc_broker_comm, 0)),
                                   0)
                       END,
                       0)                               pr_lc_broker_comm,
                   NVL (
                       CASE
                           WHEN pr_net_effect IN ('Credit')
                           THEN
                               NVL (
                                   (  (DECODE (
                                           :p_currency,
                                           NULL, NVL (
                                                     (  NVL (pr_fc_broker_tax, 0)
                                                      * pr_cur_rate),
                                                     0),
                                           NVL (pr_fc_broker_tax, 0)))
                                    * -1),
                                   0)
                           ELSE
                               NVL (
                                   DECODE (
                                       :p_currency,
                                       NULL, NVL (
                                                 (  NVL (pr_fc_broker_tax, 0)
                                                  * pr_cur_rate),
                                                 0),
                                       NVL (pr_fc_broker_tax, 0)),
                                   0)
                       END,
                       0)                               pr_lc_broker_tax,
                   NVL (
                       CASE
                           WHEN pr_net_effect IN ('Credit')
                           THEN
                               NVL (
                                   (  (DECODE (
                                           :p_currency,
                                           NULL, NVL (
                                                     (  NVL (pr_fc_stamp_duty, 0)
                                                      * pr_cur_rate),
                                                     0),
                                           NVL (pr_fc_stamp_duty, 0)))
                                    * -1),
                                   0)
                           ELSE
                               NVL (
                                   DECODE (
                                       :p_currency,
                                       NULL, NVL (
                                                 (  NVL (pr_fc_stamp_duty, 0)
                                                  * pr_cur_rate),
                                                 0),
                                       NVL (pr_fc_stamp_duty, 0)),
                                   0)
                       END,
                       0)                               pr_lc_stamp_duty,
                   NVL (
                       CASE
                           WHEN pr_net_effect IN ('Credit')
                           THEN
                               NVL (
                                   (  (DECODE (
                                           :p_currency,
                                           NULL, NVL (
                                                     (  NVL (pr_fc_phc_fund, 0)
                                                      * pr_cur_rate),
                                                     0),
                                           NVL (pr_fc_phc_fund, 0)))
                                    * -1),
                                   0)
                           ELSE
                               NVL (
                                   DECODE (
                                       :p_currency,
                                       NULL, NVL (
                                                 (  NVL (pr_fc_phc_fund, 0)
                                                  * pr_cur_rate),
                                                 0),
                                       NVL (pr_fc_phc_fund, 0)),
                                   0)
                       END,
                       0)                               pr_lc_phc_fund,
                   NVL (
                       CASE
                           WHEN pr_net_effect IN ('Credit')
                           THEN
                               NVL (
                                   (  (DECODE (
                                           :p_currency,
                                           NULL, NVL (
                                                     (  NVL (pr_fc_training_levy,
                                                             0)
                                                      * pr_cur_rate),
                                                     0),
                                           NVL (pr_fc_training_levy, 0)))
                                    * -1),
                                   0)
                           ELSE
                               NVL (
                                   DECODE (
                                       :p_currency,
                                       NULL, NVL (
                                                 (  NVL (pr_fc_training_levy, 0)
                                                  * pr_cur_rate),
                                                 0),
                                       NVL (pr_fc_training_levy, 0)),
                                   0)
                       END,
                       0)                               pr_lc_training_levy,
                   NVL (
                       CASE
                           WHEN pr_net_effect IN ('Credit')
                           THEN
                               NVL (
                                   (  (DECODE (
                                           :p_currency,
                                           NULL, NVL (
                                                     (  NVL (pr_fc_pta, 0)
                                                      * pr_cur_rate),
                                                     0),
                                           NVL (pr_fc_pta, 0)))
                                    * -1),
                                   0)
                           ELSE
                               NVL (
                                   DECODE (
                                       :p_currency,
                                       NULL, NVL (
                                                 (  NVL (pr_fc_pta, 0)
                                                  * pr_cur_rate),
                                                 0),
                                       NVL (pr_fc_pta, 0)),
                                   0)
                       END,
                       0)                               pr_lc_pta,
                   NVL (
                       CASE
                           WHEN pr_net_effect IN ('Credit')
                           THEN
                               NVL (
                                   (  (DECODE (
                                           :p_currency,
                                           NULL, NVL (
                                                     (  NVL (pr_fc_aa, 0)
                                                      * pr_cur_rate),
                                                     0),
                                           NVL (pr_fc_pta, 0)))
                                    * -1),
                                   0)
                           ELSE
                               NVL (
                                   DECODE (
                                       :p_currency,
                                       NULL, NVL (
                                                 (NVL (pr_fc_aa, 0) * pr_cur_rate),
                                                 0),
                                       NVL (pr_fc_pta, 0)),
                                   0)
                       END,
                       0)                               pr_lc_aa,
                   NVL (
                       CASE
                           WHEN pr_net_effect IN ('Credit')
                           THEN
                               NVL (
                                   (  (DECODE (
                                           :p_currency,
                                           NULL, NVL (
                                                     (  NVL (pr_fc_loading, 0)
                                                      * pr_cur_rate),
                                                     0),
                                           NVL (pr_fc_loading, 0)))
                                    * -1),
                                   0)
                           ELSE
                               NVL (
                                   DECODE (
                                       :p_currency,
                                       NULL, NVL (
                                                 (  NVL (pr_fc_loading, 0)
                                                  * pr_cur_rate),
                                                 0),
                                       NVL (pr_fc_loading, 0)),
                                   0)
                       END,
                       0)                               pr_lc_loading,
                   NVL (
                       CASE
                           WHEN pr_net_effect IN ('Credit')
                           THEN
                               NVL (
                                   (  (DECODE (
                                           :p_currency,
                                           NULL, NVL (
                                                     (  NVL (pr_fc_discount, 0)
                                                      * pr_cur_rate),
                                                     0),
                                           NVL (pr_fc_discount, 0)))
                                    * -1),
                                   0)
                           ELSE
                               NVL (
                                   DECODE (
                                       :p_currency,
                                       NULL, NVL (
                                                 (  NVL (pr_fc_discount, 0)
                                                  * pr_cur_rate),
                                                 0),
                                       NVL (pr_fc_discount, 0)),
                                   0)
                       END,
                       0)                               pr_lc_discount
              FROM uw_premium_register a, all_entity b
             WHERE     a.pr_int_aent_code = b.ent_aent_code(+)
                   AND a.pr_int_ent_code = b.ent_code(+)
                   AND pr_org_code = :p_org_code
                   AND pr_bus_type != '3000'
                   AND pr_fm_dt = NVL ( :pr_fm_dt, pr_fm_dt)
                   AND pr_to_dt = NVL ( :pr_to_dt, pr_to_dt)
          ORDER BY pr_org_code, pr_pl_index, pr_end_index)
GROUP BY pr_sub_class, pr_class
  `;
      const results = (await connection).execute(query, {
        p_org_code: "50",
        p_currency: "",
        pr_fm_dt: new Date(fromDate),
        pr_to_dt: new Date(toDate),
      });
      const finalResults = formatOracleData(await results);

      const updateWorkbook = (workbook) => {
        const worksheet = workbook.getWorksheet("70-3A");

        Object.entries(classSubclassRowMapper6).forEach(
          ([classSubKey, targetRow]) => {
            const [classKey, subClassKey] = classSubKey.split("|");

            // Filter the results based on the class and subclass
            const filteredResults = finalResults.filter(
              (item) =>
                item.PR_CLASS === classKey &&
                (subClassKey === "" || item.PR_SUB_CLASS === subClassKey)
            );

            if (filteredResults.length > 0) {
              // If subClassKey is empty, calculate totals for fields where PR_CLASS is the same
              if (subClassKey === "") {
                const totals = {};

                filteredResults.forEach((dataItem) => {
                  for (const [field, column] of Object.entries(cellMapper7)) {
                    // Initialize totals if not present
                    if (!totals[field]) totals[field] = 0;

                    // Add the values from the data items
                    totals[field] += dataItem[field] || 0;
                  }
                });

                // Write totals to the targetRow
                for (const [field, column] of Object.entries(cellMapper7)) {
                  const cell = worksheet.getCell(`${column}${targetRow}`);
                  cell.value = totals[field];
                  console.log(
                    `Total ${field} (${column}${targetRow}): ${totals[field]}`
                  );
                }
              } else {
                // For non-empty subClassKey, process individual records
                filteredResults.forEach((dataItem) => {
                  for (const [field, column] of Object.entries(cellMapper7)) {
                    const cell = worksheet.getCell(`${column}${targetRow}`);
                    cell.value = dataItem[field];
                    console.log(
                      `${field} (${column}${targetRow}): ${dataItem[field]}`
                    );
                  }
                });
              }
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
