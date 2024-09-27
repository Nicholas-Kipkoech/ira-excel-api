import pool from "../config/database.js";
import ExcelJs from "exceljs";
import {
  cellMapper6,
  classSubclassRowMapper5,
} from "./IRA-class-prem-mapper.js";
import formatOracleData from "../utils/helpers.js";

//the file path
const filePath = "test_file.xlsx";

export class IRAUnearnedPremiumsService {
  constructor() {}

  static async getUnearnedPremiums(req, res) {
    let connection;
    try {
      const { fromDate, toDate } = req.query;
      connection = (await pool).getConnection();
      if (connection) {
        console.log("Database connected...");
      }
      let query = `
      /* Formatted on 9/27/2024 8:20:09 AM (QP5 v5.336) */
  SELECT DISTINCT pr_org_code,
                  pr_mc_code,
                  INITCAP (
                      CASE
                          WHEN a.pr_sc_code IN ('0804') THEN 'PSV'
                          ELSE pr_mc_name
                      END)    class,
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
                  END         pr_order,
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
                              WHEN a.pr_pl_no LIKE '%TP%'
                              THEN
                                  'Third Party Only'
                              ELSE
                                  'Comprehensive'
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
                  END         sub_class,
                  NVL (
                      ROUND (
                          SUM (
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
                              END),
                          0),
                      0)      goss_prem,
                  NVL (
                      ROUND (
                          SUM (
                              NVL (
                                  ROUND (
                                      CASE
                                          WHEN pr_bus_type NOT IN ('3000')
                                          THEN
                                                (CASE
                                                     WHEN TRUNC (
                                                                ADD_MONTHS (
                                                                    TRUNC (
                                                                        :p_asatdate,
                                                                        'YEAR'),
                                                                    -1)
                                                              + 30) BETWEEN TRUNC (
                                                                                pr_fm_dt)
                                                                        AND TRUNC (
                                                                                pr_to_dt)
                                                     THEN
                                                         (  (TO_NUMBER (
                                                                   TRUNC (
                                                                       pr_to_dt)
                                                                 - TRUNC (
                                                                         ADD_MONTHS (
                                                                             TRUNC (
                                                                                 :p_asatdate,
                                                                                 'YEAR'),
                                                                             -1)
                                                                       + 30)
                                                                 + 1))
                                                          / NULLIF (
                                                                (TO_NUMBER (
                                                                       TRUNC (
                                                                           pr_to_dt)
                                                                     - TRUNC (
                                                                           pr_fm_dt)
                                                                     + 1)),
                                                                0))
                                                     ELSE
                                                         0
                                                 END)
                                              * (CASE
                                                     WHEN a.pr_net_effect IN
                                                              ('Credit')
                                                     THEN
                                                         (  (  (  NVL (
                                                                      a.pr_fc_prem,
                                                                      0)
                                                                * a.pr_cur_rate)
                                                             + (  NVL (
                                                                      a.pr_fc_eartquake,
                                                                      0)
                                                                * a.pr_cur_rate)
                                                             + (  NVL (
                                                                      a.pr_fc_political,
                                                                      0)
                                                                * a.pr_cur_rate))
                                                          * -1)
                                                     ELSE
                                                         (  (  NVL (
                                                                   a.pr_fc_prem,
                                                                   0)
                                                             * a.pr_cur_rate)
                                                          + (  NVL (
                                                                   a.pr_fc_eartquake,
                                                                   0)
                                                             * a.pr_cur_rate)
                                                          + (  NVL (
                                                                   a.pr_fc_political,
                                                                   0)
                                                             * a.pr_cur_rate))
                                                 END)
                                          ELSE
                                              0
                                      END,
                                      0),
                                  0)),
                          0),
                      0)      gross_upr_bfwd_direct,
                  NVL (
                      ROUND (
                          SUM (
                              NVL (
                                  ROUND (
                                      CASE
                                          WHEN pr_bus_type IN ('3000')
                                          THEN
                                                (CASE
                                                     WHEN TRUNC (
                                                                ADD_MONTHS (
                                                                    TRUNC (
                                                                        :p_asatdate,
                                                                        'YEAR'),
                                                                    -1)
                                                              + 30) BETWEEN TRUNC (
                                                                                pr_fm_dt)
                                                                        AND TRUNC (
                                                                                pr_to_dt)
                                                     THEN
                                                         (  (TO_NUMBER (
                                                                   TRUNC (
                                                                       pr_to_dt)
                                                                 - TRUNC (
                                                                         ADD_MONTHS (
                                                                             TRUNC (
                                                                                 :p_asatdate,
                                                                                 'YEAR'),
                                                                             -1)
                                                                       + 30)
                                                                 + 1))
                                                          / NULLIF (
                                                                (TO_NUMBER (
                                                                       TRUNC (
                                                                           pr_to_dt)
                                                                     - TRUNC (
                                                                           pr_fm_dt)
                                                                     + 1)),
                                                                0))
                                                     ELSE
                                                         0
                                                 END)
                                              * (CASE
                                                     WHEN a.pr_net_effect IN
                                                              ('Credit')
                                                     THEN
                                                         (  (  (  NVL (
                                                                      a.pr_fc_prem,
                                                                      0)
                                                                * a.pr_cur_rate)
                                                             + (  NVL (
                                                                      a.pr_fc_eartquake,
                                                                      0)
                                                                * a.pr_cur_rate)
                                                             + (  NVL (
                                                                      a.pr_fc_political,
                                                                      0)
                                                                * a.pr_cur_rate))
                                                          * -1)
                                                     ELSE
                                                         (  (  NVL (
                                                                   a.pr_fc_prem,
                                                                   0)
                                                             * a.pr_cur_rate)
                                                          + (  NVL (
                                                                   a.pr_fc_eartquake,
                                                                   0)
                                                             * a.pr_cur_rate)
                                                          + (  NVL (
                                                                   a.pr_fc_political,
                                                                   0)
                                                             * a.pr_cur_rate))
                                                 END)
                                          ELSE
                                              0
                                      END,
                                      0),
                                  0)),
                          0),
                      0)      gross_upr_bfwd_facin,
                  NVL (
                      ROUND (
                          SUM (
                              NVL (
                                  ROUND (
                                        (CASE
                                             WHEN TRUNC ( :p_asatdate, 'Year') BETWEEN TRUNC (
                                                                                           bp_fm_date)
                                                                                   AND TRUNC (
                                                                                           bp_to_date)
                                             THEN
                                                 (  (TO_NUMBER (
                                                           TRUNC (bp_to_date)
                                                         - TRUNC ( :p_asatdate,
                                                                  'Year')
                                                         + 1))
                                                  / NULLIF (
                                                        (TO_NUMBER (
                                                               TRUNC (
                                                                   bp_to_date)
                                                             - TRUNC (
                                                                   bp_fm_date)
                                                             + 1)),
                                                        0))
                                             ELSE
                                                 0
                                         END)
                                      * (NVL (outward_prem, 0)),
                                      0),
                                  0)),
                          0),
                      0)      gross_upr_bfwd_outward,
                  NVL (
                      ROUND (
                          SUM (
                              NVL (
                                  ROUND (
                                      CASE
                                          WHEN pr_bus_type NOT IN ('3000')
                                          THEN
                                                (CASE
                                                     WHEN TRUNC ( :p_asatdate) BETWEEN TRUNC (
                                                                                           pr_fm_dt)
                                                                                   AND TRUNC (
                                                                                           pr_to_dt)
                                                     THEN
                                                         (  (TO_NUMBER (
                                                                   TRUNC (
                                                                       pr_to_dt)
                                                                 - TRUNC (
                                                                       :p_asatdate)
                                                                 + 1))
                                                          / NULLIF (
                                                                (TO_NUMBER (
                                                                       TRUNC (
                                                                           pr_to_dt)
                                                                     - TRUNC (
                                                                           pr_fm_dt)
                                                                     + 1)),
                                                                0))
                                                     ELSE
                                                         0
                                                 END)
                                              * (CASE
                                                     WHEN a.pr_net_effect IN
                                                              ('Credit')
                                                     THEN
                                                         (  (  (  NVL (
                                                                      a.pr_fc_prem,
                                                                      0)
                                                                * a.pr_cur_rate)
                                                             + (  NVL (
                                                                      a.pr_fc_eartquake,
                                                                      0)
                                                                * a.pr_cur_rate)
                                                             + (  NVL (
                                                                      a.pr_fc_political,
                                                                      0)
                                                                * a.pr_cur_rate))
                                                          * -1)
                                                     ELSE
                                                         (  (  NVL (
                                                                   a.pr_fc_prem,
                                                                   0)
                                                             * a.pr_cur_rate)
                                                          + (  NVL (
                                                                   a.pr_fc_eartquake,
                                                                   0)
                                                             * a.pr_cur_rate)
                                                          + (  NVL (
                                                                   a.pr_fc_political,
                                                                   0)
                                                             * a.pr_cur_rate))
                                                 END)
                                          ELSE
                                              0
                                      END,
                                      0),
                                  0)),
                          0),
                      0)      gross_upr_cfwd_direct,
                  NVL (
                      ROUND (
                          SUM (
                              NVL (
                                  ROUND (
                                      CASE
                                          WHEN pr_bus_type IN ('3000')
                                          THEN
                                                (CASE
                                                     WHEN TRUNC ( :p_asatdate) BETWEEN TRUNC (
                                                                                           pr_fm_dt)
                                                                                   AND TRUNC (
                                                                                           pr_to_dt)
                                                     THEN
                                                         (  (TO_NUMBER (
                                                                   TRUNC (
                                                                       pr_to_dt)
                                                                 - TRUNC (
                                                                       :p_asatdate)
                                                                 + 1))
                                                          / NULLIF (
                                                                (TO_NUMBER (
                                                                       TRUNC (
                                                                           pr_to_dt)
                                                                     - TRUNC (
                                                                           pr_fm_dt)
                                                                     + 1)),
                                                                0))
                                                     ELSE
                                                         0
                                                 END)
                                              * (CASE
                                                     WHEN a.pr_net_effect IN
                                                              ('Credit')
                                                     THEN
                                                         (  (  (  NVL (
                                                                      a.pr_fc_prem,
                                                                      0)
                                                                * a.pr_cur_rate)
                                                             + (  NVL (
                                                                      a.pr_fc_eartquake,
                                                                      0)
                                                                * a.pr_cur_rate)
                                                             + (  NVL (
                                                                      a.pr_fc_political,
                                                                      0)
                                                                * a.pr_cur_rate))
                                                          * -1)
                                                     ELSE
                                                         (  (  NVL (
                                                                   a.pr_fc_prem,
                                                                   0)
                                                             * a.pr_cur_rate)
                                                          + (  NVL (
                                                                   a.pr_fc_eartquake,
                                                                   0)
                                                             * a.pr_cur_rate)
                                                          + (  NVL (
                                                                   a.pr_fc_political,
                                                                   0)
                                                             * a.pr_cur_rate))
                                                 END)
                                          ELSE
                                              0
                                      END,
                                      0),
                                  0)),
                          0),
                      0)      gross_upr_cfwd_facin,
                  NVL (
                      ROUND (
                          SUM (
                              NVL (
                                  ROUND (
                                        (CASE
                                             WHEN TRUNC ( :p_asatdate) BETWEEN TRUNC (
                                                                                   bp_fm_date)
                                                                           AND TRUNC (
                                                                                   bp_to_date)
                                             THEN
                                                 (  (TO_NUMBER (
                                                           TRUNC (bp_to_date)
                                                         - TRUNC ( :p_asatdate)
                                                         + 1))
                                                  / NULLIF (
                                                        (TO_NUMBER (
                                                               TRUNC (
                                                                   bp_to_date)
                                                             - TRUNC (
                                                                   bp_fm_date)
                                                             + 1)),
                                                        0))
                                             ELSE
                                                 0
                                         END)
                                      * (NVL (outward_prem, 0)),
                                      0),
                                  0)),
                          0),
                      0)      gross_upr_cfwd_outward
    FROM uw_premium_register a,
         (SELECT DISTINCT
                 bh_org_code,
                 bh_pol_index,
                 bh_pol_end_index,
                 bh_pol_no,
                 bh_pol_end_no,
                 pr_fm_dt    bp_fm_date,
                 pr_to_dt    bp_to_date,
                 ROUND (
                     CASE
                         WHEN    p.pr_net_effect IN ('Credit')
                              OR bh_status IN ('Reversed')
                         THEN
                             (NVL (outward_prem, 0) * -1)
                         ELSE
                             NVL (outward_prem, 0)
                     END,
                     0)      outward_prem
            FROM ri_batch_header    a,
                 ri_batch_cover_risk b,
                 uw_premium_register p,
                 (  SELECT bl_org_code,
                           bl_batch_no,
                           bl_cr_index,
                             NVL (
                                 SUM (
                                     DECODE (bl_line_type_int,
                                             'Surplus 1', bl_lc_prem)),
                                 0)
                           + NVL (
                                 SUM (
                                     DECODE (bl_line_type_int,
                                             'Surplus 2', bl_lc_prem)),
                                 0)
                           + NVL (
                                 SUM (
                                     DECODE (bl_line_type_int,
                                             'FAC Out', bl_lc_prem)),
                                 0)
                           + NVL (
                                 SUM (
                                     DECODE (bl_line_type_int, 'QS', bl_lc_prem)),
                                 0)    outward_prem
                      FROM ri_batch_lines
                     WHERE bl_line_type_int NOT IN ('Balance', 'Retention')
                  GROUP BY bl_org_code, bl_batch_no, bl_cr_index) d
           WHERE     a.bh_org_code = b.cr_org_code
                 AND a.bh_batch_no = b.cr_batch_no
                 AND b.cr_org_code = d.bl_org_code
                 AND b.cr_batch_no = d.bl_batch_no
                 AND b.cr_index = d.bl_cr_index
                 AND a.bh_org_code = p.pr_org_code(+)
                 AND a.bh_pol_index = p.pr_pl_index(+)
                 AND a.bh_pol_end_index = p.pr_end_index(+)
                 AND a.bh_status = 'Completed'
                 AND a.bh_org_code = :p_org_code --   AND TRUNC (a.bh_gl_date) <= TRUNC (:p_asatdate)
                                                ) b
   WHERE     a.pr_org_code = b.bh_org_code(+)
         AND a.pr_pl_index = b.bh_pol_index(+)
         AND a.pr_end_index = b.bh_pol_end_index(+)
         --    AND TRUNC (pr_gl_date) <= TRUNC (:p_asatdate)
         AND pr_org_code = :p_org_code
GROUP BY pr_org_code,
         pr_mc_code,
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
ORDER BY 1, 2, 4`;
      const results = (await connection).execute(query, {
        p_org_code: "50",
        p_asatdate: new Date(toDate),
      });
      const finalResults = formatOracleData(await results);
      //initiate the workbook or the excel package
      const workbook = new ExcelJs.Workbook();

      workbook.xlsx
        .readFile(filePath)
        .then(() => {
          const worksheet = workbook.getWorksheet("59-1B (c)");

          Object.entries(classSubclassRowMapper5).forEach(
            ([classSubKey, targetRow]) => {
              const [classKey, subClassKey] = classSubKey.split("|");

              const filteredResults = finalResults.filter(
                (item) =>
                  item.CLASS === classKey && item.SUB_CLASS === subClassKey
              );
              if (filteredResults.length > 0) {
                filteredResults.forEach((dataItem) => {
                  for (const [field, column] of Object.entries(cellMapper6)) {
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
        })
        .then(async () => {
          await workbook.xlsx.writeFile(filePath);
          return console.log("Data written successfully");
        })
        .catch((err) => {
          console.error("Error modifying the Excel file:", err);
        });

      return res.status(200).json({ results: finalResults });
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
