import pool from "../config/database.js";
import ExcelJs from "exceljs";
import {
  cellMapper4,
  cellMapper5,
  classSubclassRowMapper3,
  classSubclassRowMapper4,
} from "./IRA-class-prem-mapper.js";
import formatOracleData from "../utils/helpers.js";
import { writeFileSafely } from "./excel-service/excel-helper.js";

//the file path
const filePath = "IRA_excel.xlsx";

export class IRAReinsurancePremiumsService {
  constructor() {}

  static async getReinsurancePremiums(req, res) {
    let connection;
    try {
      const { fromDate, toDate } = req.query;
      connection = (await pool).getConnection();
      if (connection) {
        console.log("Database connected...");
      }
      let query = `/* Formatted on 9/26/2024 6:33:48 PM (QP5 v5.336) */
  SELECT a.bh_org_code,
         b.cr_mc_code,
         INITCAP (
             CASE
                 WHEN b.cr_sc_code IN ('0804')
                 THEN
                     'PSV'
                 ELSE
                     pkg_system_admin.get_class_name (cr_org_code,
                                                      b.cr_mc_code)
             END)    class,
         CASE
             WHEN b.cr_mc_code IN ('03',
                                   '04',
                                   '09',
                                   '11')
             THEN
                 1
             WHEN b.cr_mc_code IN ('070', '080')
             THEN
                 CASE WHEN a.bh_pol_no LIKE '%TP%' THEN 2 ELSE 1 END
             ELSE
                 CASE
                     WHEN cr_sc_code IN ('010',
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
             WHEN b.cr_mc_code IN ('03',
                                   '04',
                                   '09',
                                   '11')
             THEN
                 pkg_system_admin.get_class_name (cr_org_code, b.cr_mc_code)
             WHEN b.cr_mc_code IN ('070', '080')
             THEN
                 CASE
                     WHEN a.bh_pol_no LIKE '%TP%' THEN 'Third Party Only'
                     ELSE 'Comprehensive'
                 END
             ELSE
                 CASE
                     WHEN cr_sc_code IN ('010',
                                         '020',
                                         '050',
                                         '051',
                                         '060',
                                         '061',
                                         '064',
                                         '100',
                                         '101')
                     THEN
                         pkg_system_admin.get_subclass_name (cr_org_code,
                                                             b.cr_sc_code)
                     WHEN cr_sc_code IN ('120', '127', '128')
                     THEN
                         'Bonds'
                     ELSE
                         CASE
                             WHEN b.cr_mc_code IN ('10') THEN 'Burglary Others'
                             ELSE 'Others'
                         END
                 END
         END         sub_class,
           NVL (
               SUM (
                   ROUND (
                       CASE
                           WHEN    p.pr_net_effect IN ('Credit')
                                OR bh_status IN ('Reversed')
                           THEN
                               (NVL (surp1_prem, 0) * -1)
                           ELSE
                               NVL (surp1_prem, 0)
                       END,
                       0)),
               0)
         + NVL (
               SUM (
                   ROUND (
                       CASE
                           WHEN    p.pr_net_effect IN ('Credit')
                                OR bh_status IN ('Reversed')
                           THEN
                               (NVL (surp2_prem, 0) * -1)
                           ELSE
                               NVL (surp2_prem, 0)
                       END,
                       0)),
               0)    prem_surplus,
         NVL (
             SUM (
                 ROUND (
                     CASE
                         WHEN    p.pr_net_effect IN ('Credit')
                              OR bh_status IN ('Reversed')
                         THEN
                             (NVL (facout_prem, 0) * -1)
                         ELSE
                             NVL (facout_prem, 0)
                     END,
                     0)),
             0)      prem_fac,
         NVL (
             SUM (
                 ROUND (
                     CASE
                         WHEN    p.pr_net_effect IN ('Credit')
                              OR bh_status IN ('Reversed')
                         THEN
                             (NVL (qs_prem, 0) * -1)
                         ELSE
                             NVL (qs_prem, 0)
                     END,
                     0)),
             0)      prem_qs,
         0           prem_xol
    FROM ri_batch_header    a,
         ri_batch_cover_risk b,
         ri_batch_policy    c,
         uw_premium_register p,
         (  SELECT bl_org_code,
                   bl_batch_no,
                   bl_cr_index,
                   NVL (SUM (DECODE (bl_line_type_int, 'Surplus 1', bl_lc_prem)),
                        0)                                                       surp1_prem,
                   NVL (SUM (DECODE (bl_line_type_int, 'Surplus 2', bl_lc_prem)),
                        0)                                                       surp2_prem,
                   NVL (SUM (DECODE (bl_line_type_int, 'FAC Out', bl_lc_prem)),
                        0)                                                       facout_prem,
                   NVL (SUM (DECODE (bl_line_type_int, 'QS', bl_lc_prem)), 0)    qs_prem
              FROM ri_batch_lines
             WHERE bl_line_type_int NOT IN ('Balance')
          GROUP BY bl_org_code, bl_batch_no, bl_cr_index) d,
         (  SELECT trn_org_code,
                   trn_ri_batch_no,
                   trn_ri_cr_index,
                   NVL (
                       SUM (
                           DECODE (
                               trn_type,
                               'RI.005', NVL (
                                             DECODE (
                                                 trn_flex06,
                                                 'QS', NVL (trn_doc_lc_amt, 0)),
                                             0),
                               'RI.005R', (  NVL (
                                                 DECODE (
                                                     trn_flex06,
                                                     'QS', NVL (trn_doc_lc_amt,
                                                                0)),
                                                 0)
                                           * -1))),
                       0)    qs_comm,
                   NVL (
                       SUM (
                           DECODE (
                               trn_type,
                               'RI.017', NVL (
                                             DECODE (
                                                 trn_flex06,
                                                 'Surplus 1', NVL (
                                                                  trn_doc_lc_amt,
                                                                  0)),
                                             0),
                               'RI.017R', (  NVL (
                                                 DECODE (
                                                     trn_flex06,
                                                     'Surplus 1', NVL (
                                                                      trn_doc_lc_amt,
                                                                      0)),
                                                 0)
                                           * -1))),
                       0)    surp1_comm,
                   NVL (
                       SUM (
                           DECODE (
                               trn_type,
                               'RI.017', NVL (
                                             DECODE (
                                                 trn_flex06,
                                                 'Surplus 2', NVL (
                                                                  trn_doc_lc_amt,
                                                                  0)),
                                             0),
                               'RI.017R', (  NVL (
                                                 DECODE (
                                                     trn_flex06,
                                                     'Surplus 2', NVL (
                                                                      trn_doc_lc_amt,
                                                                      0)),
                                                 0)
                                           * -1))),
                       0)    surp2_comm,
                   NVL (
                       SUM (
                           DECODE (
                               trn_type,
                               'RI.021', NVL (
                                             DECODE (
                                                 trn_flex06,
                                                 'FAC Out', NVL (trn_doc_lc_amt,
                                                                 0)),
                                             0),
                               'RI.021R', (  NVL (
                                                 DECODE (
                                                     trn_flex06,
                                                     'FAC Out', NVL (
                                                                    trn_doc_lc_amt,
                                                                    0)),
                                                 0)
                                           * -1))),
                       0)    fac_comm
              FROM gl_transactions
             WHERE trn_module_code = 'RI' AND trn_ent_code IS NOT NULL
          GROUP BY trn_org_code, trn_ri_batch_no, trn_ri_cr_index) e,
         (  SELECT DISTINCT
                   pl_org_code,
                   pl_pl_index,
                   pl_end_index,
                   LISTAGG (pl_eq_zone, '/') WITHIN GROUP (ORDER BY pl_eq_zone)    pl_eq_zone
              FROM (SELECT DISTINCT
                           pl_org_code,
                           pl_pl_index,
                           pl_end_index,
                           pkg_system_admin.get_system_desc ('UW_EQ_ZONE',
                                                             pl_eq_zone)    pl_eq_zone
                      FROM uh_policy_risks
                     WHERE pl_eq_zone IS NOT NULL)
          GROUP BY pl_org_code, pl_pl_index, pl_end_index) f
   WHERE     a.bh_org_code = b.cr_org_code
         AND a.bh_batch_no = b.cr_batch_no
         AND a.bh_org_code = c.bp_org_code
         AND a.bh_batch_no = c.bp_batch_no
         AND b.cr_org_code = d.bl_org_code
         AND b.cr_batch_no = d.bl_batch_no
         AND b.cr_index = d.bl_cr_index
         AND a.bh_org_code = p.pr_org_code(+)
         AND a.bh_pol_index = p.pr_pl_index(+)
         AND a.bh_pol_end_index = p.pr_end_index(+)
         AND b.cr_org_code = e.trn_org_code(+)
         AND b.cr_batch_no = e.trn_ri_batch_no(+)
         AND b.cr_index = e.trn_ri_cr_index(+)
         AND a.bh_org_code = f.pl_org_code(+)
         AND a.bh_pol_index = f.pl_pl_index(+)
         AND a.bh_pol_end_index = f.pl_end_index(+)
         AND a.bh_status = 'Completed'
         AND a.bh_org_code = :p_org_code
         AND DECODE (b.cr_cv_code, '043', '04', b.cr_mc_code) =
             NVL ( :p_class, DECODE (b.cr_cv_code, '043', '04', b.cr_mc_code))
         AND DECODE (b.cr_cv_code, '043', '043', b.cr_sc_code) =
             NVL ( :p_subclass,
                  DECODE (b.cr_cv_code, '043', '043', b.cr_sc_code))
         AND TRUNC (a.bh_gl_date) BETWEEN TRUNC (NVL ( :p_fm_dt, a.bh_gl_date))
                                      AND TRUNC (NVL ( :p_to_dt, a.bh_gl_date))
GROUP BY a.bh_org_code,
         b.cr_mc_code,
         INITCAP (
             CASE
                 WHEN b.cr_sc_code IN ('0804')
                 THEN
                     'PSV'
                 ELSE
                     pkg_system_admin.get_class_name (cr_org_code,
                                                      b.cr_mc_code)
             END),
         CASE
             WHEN b.cr_mc_code IN ('03',
                                   '04',
                                   '09',
                                   '11')
             THEN
                 1
             WHEN b.cr_mc_code IN ('070', '080')
             THEN
                 CASE WHEN a.bh_pol_no LIKE '%TP%' THEN 2 ELSE 1 END
             ELSE
                 CASE
                     WHEN cr_sc_code IN ('010',
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
             WHEN b.cr_mc_code IN ('03',
                                   '04',
                                   '09',
                                   '11')
             THEN
                 pkg_system_admin.get_class_name (cr_org_code, b.cr_mc_code)
             WHEN b.cr_mc_code IN ('070', '080')
             THEN
                 CASE
                     WHEN a.bh_pol_no LIKE '%TP%' THEN 'Third Party Only'
                     ELSE 'Comprehensive'
                 END
             ELSE
                 CASE
                     WHEN cr_sc_code IN ('010',
                                         '020',
                                         '050',
                                         '051',
                                         '060',
                                         '061',
                                         '064',
                                         '100',
                                         '101')
                     THEN
                         pkg_system_admin.get_subclass_name (cr_org_code,
                                                             b.cr_sc_code)
                     WHEN cr_sc_code IN ('120', '127', '128')
                     THEN
                         'Bonds'
                     ELSE
                         CASE
                             WHEN b.cr_mc_code IN ('10')
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
        p_class: "",
        p_subclass: "",
        p_fm_dt: new Date(fromDate),
        p_to_dt: new Date(toDate),
      });
      const finalResults = formatOracleData(await results);
      const updateWorkbook = (workbook) => {
        const worksheet = workbook.getWorksheet("59-1B (b)");

        Object.entries(classSubclassRowMapper4).forEach(
          ([classSubKey, targetRow]) => {
            const [classKey, subClassKey] = classSubKey.split("|");

            const filteredResults = finalResults.filter(
              (item) =>
                item.CLASS === classKey && item.SUB_CLASS === subClassKey
            );
            if (filteredResults.length > 0) {
              filteredResults.forEach((dataItem) => {
                for (const [field, column] of Object.entries(cellMapper5)) {
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
      await writeFileSafely(filePath, updateWorkbook);

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
