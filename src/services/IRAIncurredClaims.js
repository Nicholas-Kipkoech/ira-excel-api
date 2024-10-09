import pool from "../config/database.js";
import ExcelJs from "exceljs";
import {
  cellMapper4,
  classSubclassRowMapper3,
} from "./IRA-class-prem-mapper.js";
import formatOracleData from "../utils/helpers.js";
import { writeFileSafely } from "./excel-service/excel-helper.js";

//the file path
const filePath = "IRA_excel.xlsx";

export class IRAIncurredClaimsService {
  constructor() {}

  static async getIncuredClaims(req, res) {
    let connection;
    try {
      const { fromDate, toDate } = req.query;
      connection = (await pool).getConnection();
      if (connection) {
        console.log("Database connected...");
      }
      let query = `/* Formatted on 9/26/2024 11:56:58 AM (QP5 v5.336) */
  SELECT 1
             order_no,
         hd_org_code
             org_code,
         cr_mc_code
             mc_code,
         INITCAP (
             CASE
                 WHEN cr_sc_code IN ('0804') THEN 'PSV'
                 ELSE pkg_system_admin.get_class_name (hd_org_code, cr_mc_code)
             END)
             class,
         CASE
             WHEN cr_mc_code IN ('03',
                                 '04',
                                 '09',
                                 '11')
             THEN
                 1
             WHEN cr_mc_code IN ('070', '080')
             THEN
                 CASE WHEN pl_no LIKE '%TP%' THEN 2 ELSE 1 END
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
         END
             cm_order,
         CASE
             WHEN cr_mc_code IN ('03',
                                 '04',
                                 '09',
                                 '11')
             THEN
                 pkg_system_admin.get_class_name (hd_org_code, cr_mc_code)
             WHEN cr_mc_code IN ('070', '080')
             THEN
                 CASE
                     WHEN pl_no LIKE '%TP%' THEN 'Third Party Only'
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
                         pkg_system_admin.get_subclass_name (hd_org_code,
                                                             cr_sc_code)
                     WHEN cr_sc_code IN ('120', '127', '128')
                     THEN
                         'Bonds'
                     ELSE
                         CASE
                             WHEN cr_mc_code IN ('10') THEN 'Burglary Others'
                             ELSE 'Others'
                         END
                 END
         END
             sub_class,
         SUM (NVL (hd_paid, 0) - NVL (hd_receipts, 0))
             total,
         SUM (NVL (hd_outward_paid, 0) - NVL (hd_outward_receipts, 0))
             outward_recovery,
         SUM (NVL (ROUND (CASE
                              WHEN    pr_bus_type IN ('1000',
                                                      '1001',
                                                      '1002',
                                                      '2000',
                                                      '2999')
                                   OR pr_bus_type IS NULL
                              THEN
                                  NVL (hd_paid, 0) - NVL (hd_receipts, 0)
                              ELSE
                                  0
                          END,
                          0),
                   0))
             direct,
         SUM (
             NVL (
                 ROUND (
                     CASE
                         WHEN pr_bus_type IN ('3000')
                         THEN
                             NVL (hd_paid, 0) - NVL (hd_receipts, 0)
                         ELSE
                             0
                     END,
                     0),
                 0))
             facin
    FROM (SELECT a.hd_org_code,
                 NVL (b.cm_pl_no, p.pr_pl_no)            pl_no,
                 TRUNC (a.hd_gl_date)                    gl_date,
                 pr_bus_type,
                 cr_mc_code,
                 cr_sc_code,
                 (NVL (e.do_fc_amt, 0) * hd_cur_rate)    hd_paid,
                 0                                       hd_receipts,
                 ROUND (
                     (  (NVL (e.do_fc_amt, 0) * hd_cur_rate)
                      * (SELECT ABS (SUM (r.rs_percent))
                           FROM uw_policy_ri_shares r
                          WHERE     r.rs_org_code = c.cr_org_code
                                AND r.rs_ri_batch_no = c.cr_ri_batch_no
                                AND r.rs_ri_cr_index = c.cr_ri_cr_index
                                AND r.rs_type = 'Final'
                                AND r.rs_line_type_int NOT IN
                                        ('Balance', 'Retention'))
                      / 100))                            hd_outward_paid,
                 0                                       hd_outward_receipts
            FROM ap_payments_header a,
                 cm_claims          b,
                 cm_claims_risks    c,
                 ap_payment_docs    e,
                 uw_premium_register p,
                 (SELECT DISTINCT ln_org_code, ln_pmt_no, ln_chq_no
                    FROM ap_cheques_header, ap_cheques_lines
                   WHERE     hd_org_code = ln_org_code
                         AND hd_no = ln_hd_no
                         AND hd_status = 'Completed'
                         AND ln_chq_status = 'Written') d
           WHERE     a.hd_org_code = e.do_org_code
                 AND a.hd_no = e.do_hd_no
                 AND e.do_doc_type = 'Claim'
                 AND e.do_org_code = b.cm_org_code
                 AND e.do_doc_no = b.cm_no
                 AND b.cm_org_code = c.cr_org_code
                 AND b.cm_index = c.cr_cm_index
                 AND a.hd_org_code = d.ln_org_code(+)
                 AND a.hd_no = d.ln_pmt_no(+)
                 AND b.cm_org_code = p.pr_org_code(+)
                 AND b.cm_pl_index = p.pr_pl_index(+)
                 AND b.cm_end_index = p.pr_end_index(+)
          UNION ALL
          SELECT c.hd_org_code,
                 NVL (b.cm_pl_no, p.pr_pl_no)
                     pl_no,
                 TRUNC (a.trn_doc_gl_dt)
                     gl_date,
                 pr_bus_type,
                 f.cr_mc_code,
                 f.cr_sc_code,
                 (DECODE (trn_drcr_flag,
                          'C', (NVL (trn_doc_fc_amt, 0) * trn_cur_rate),
                          ((NVL (trn_doc_fc_amt, 0) * trn_cur_rate) * -1)))
                     hd_paid,
                 0
                     hd_receipts,
                 ROUND (
                     (  (DECODE (
                             trn_drcr_flag,
                             'C', (NVL (trn_doc_fc_amt, 0) * trn_cur_rate),
                             ((NVL (trn_doc_fc_amt, 0) * trn_cur_rate) * -1)))
                      * (SELECT ABS (SUM (r.rs_percent))
                           FROM uw_policy_ri_shares r
                          WHERE     r.rs_org_code = f.cr_org_code
                                AND r.rs_ri_batch_no = f.cr_ri_batch_no
                                AND r.rs_ri_cr_index = f.cr_ri_cr_index
                                AND r.rs_type = 'Final'
                                AND r.rs_line_type_int NOT IN
                                        ('Balance', 'Retention'))
                      / 100))
                     hd_outward_paid,
                 0
                     hd_outward_receipts
            FROM gl_transactions    a,
                 cm_claims          b,
                 gl_je_header       c,
                 cm_claims_risks    f,
                 uw_premium_register p
           WHERE     trn_ent_code IS NOT NULL
                 AND trn_doc_type = 'GL-JOURNAL'
                 --  AND trn_flex01 = 'CREDIT NOTE'
                 AND hd_org_code = cm_org_code(+)
                 AND hd_batch_no = cm_no(+)
                 AND trn_org_code = hd_org_code
                 AND trn_doc_no = hd_no
                 AND hd_type = 'CREDIT NOTE'
                 AND b.cm_org_code = f.cr_org_code
                 AND b.cm_index = f.cr_cm_index
                 AND b.cm_org_code = p.pr_org_code(+)
                 AND b.cm_pl_index = p.pr_pl_index(+)
                 AND b.cm_end_index = p.pr_end_index(+)
          UNION ALL
          SELECT a.hd_org_code,
                 NVL (b.cm_pl_no, q.pr_pl_no)    pl_no,
                 TRUNC (a.hd_gl_date)            gl_date,
                 pr_bus_type,
                 p.cr_mc_code,
                 p.cr_sc_code,
                 0                               hd_paid,
                 NVL (e.do_lc_amount, 0)         hd_receipts,
                 0                               hd_outward_paid,
                 ROUND (
                     (  NVL (e.do_lc_amount, 0)
                      * (SELECT ABS (SUM (r.rs_percent))
                           FROM uw_policy_ri_shares r
                          WHERE     r.rs_org_code = p.cr_org_code
                                AND r.rs_ri_batch_no = p.cr_ri_batch_no
                                AND r.rs_ri_cr_index = p.cr_ri_cr_index
                                AND r.rs_type = 'Final'
                                AND r.rs_line_type_int NOT IN
                                        ('Balance', 'Retention'))
                      / 100))                    hd_outward_share
            FROM ar_receipts_header a,
                 ar_receipt_lines   j,
                 ar_receipt_docs    e,
                 cm_claims          b,
                 cm_claims_risks    p,
                 uw_premium_register q
           WHERE     a.hd_org_code = j.ln_org_code
                 AND a.hd_no = j.ln_hd_no
                 AND j.ln_org_code = e.do_org_code
                 AND j.ln_hd_no = e.do_hd_no
                 AND e.do_doc_type = 'Claim'
                 AND a.hd_source IN
                         ('Claim Excess', 'Claim Salvage', 'T.P Recovery')
                 AND e.do_org_code = b.cm_org_code
                 AND e.do_doc_no = b.cm_no
                 AND b.cm_org_code = p.cr_org_code
                 AND b.cm_index = p.cr_cm_index
                 AND b.cm_org_code = q.pr_org_code(+)
                 AND b.cm_pl_index = q.pr_pl_index(+)
                 AND b.cm_end_index = q.pr_end_index(+))
   WHERE     hd_org_code = :p_org_code
         AND cr_mc_code = NVL ( :p_class, cr_mc_code)
         AND cr_sc_code = NVL ( :p_subclass, cr_sc_code)
         AND TRUNC (gl_date) BETWEEN TRUNC (NVL ( :p_fm_dt, gl_date))
                                 AND TRUNC (NVL ( :p_to_dt, gl_date))
GROUP BY 1,
         hd_org_code,
         cr_mc_code,
         INITCAP (
             CASE
                 WHEN cr_sc_code IN ('0804')
                 THEN
                     'PSV'
                 ELSE
                     pkg_system_admin.get_class_name (hd_org_code,
                                                      cr_mc_code)
             END),
         CASE
             WHEN cr_mc_code IN ('03',
                                 '04',
                                 '09',
                                 '11')
             THEN
                 1
             WHEN cr_mc_code IN ('070', '080')
             THEN
                 CASE WHEN pl_no LIKE '%TP%' THEN 2 ELSE 1 END
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
             WHEN cr_mc_code IN ('03',
                                 '04',
                                 '09',
                                 '11')
             THEN
                 pkg_system_admin.get_class_name (hd_org_code, cr_mc_code)
             WHEN cr_mc_code IN ('070', '080')
             THEN
                 CASE
                     WHEN pl_no LIKE '%TP%' THEN 'Third Party Only'
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
                         pkg_system_admin.get_subclass_name (hd_org_code,
                                                             cr_sc_code)
                     WHEN cr_sc_code IN ('120', '127', '128')
                     THEN
                         'Bonds'
                     ELSE
                         CASE
                             WHEN cr_mc_code IN ('10') THEN 'Burglary Others'
                             ELSE 'Others'
                         END
                 END
         END
UNION ALL
  SELECT DISTINCT 2                              order_no,
                  cm_org_code                    org_code,
                  cr_mc_code                     mc_code,
                  INITCAP (
                      CASE
                          WHEN cr_sc_code IN ('0804')
                          THEN
                              'PSV'
                          ELSE
                              pkg_system_admin.get_class_name (cm_org_code,
                                                               cr_mc_code)
                      END)                       class,
                  CASE
                      WHEN cr_mc_code IN ('03',
                                          '04',
                                          '09',
                                          '11')
                      THEN
                          1
                      WHEN cr_mc_code IN ('070', '080')
                      THEN
                          CASE
                              WHEN NVL (cm_pl_no, pr_pl_no) LIKE '%TP%' THEN 2
                              ELSE 1
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
                                                  '101',
                                                  '120',
                                                  '127',
                                                  '128')
                              THEN
                                  1
                              ELSE
                                  2
                          END
                  END                            cm_order,
                  CASE
                      WHEN cr_mc_code IN ('03',
                                          '04',
                                          '09',
                                          '11')
                      THEN
                          pkg_system_admin.get_class_name (cm_org_code,
                                                           cr_mc_code)
                      WHEN cr_mc_code IN ('070', '080')
                      THEN
                          CASE
                              WHEN NVL (cm_pl_no, pr_pl_no) LIKE '%TP%'
                              THEN
                                  'Third Party Only'
                              ELSE
                                  'Comprehensive'
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
                                  pkg_system_admin.get_subclass_name (
                                      cm_org_code,
                                      cr_sc_code)
                              WHEN cr_sc_code IN ('120', '127', '128')
                              THEN
                                  'Bonds'
                              ELSE
                                  CASE
                                      WHEN cr_mc_code IN ('10')
                                      THEN
                                          'Burglary Others'
                                      ELSE
                                          'Others'
                                  END
                          END
                  END                            sub_class,
                  SUM (NVL (reserve_amnt, 0))    total,
                  ROUND (
                      NVL (
                          SUM (
                              (  (  (NVL (reserve_amnt, 0))
                                  * (SELECT ABS (SUM (r.rs_percent))
                                       FROM uw_policy_ri_shares r
                                      WHERE     r.rs_org_code = b.cr_org_code
                                            AND r.rs_ri_batch_no =
                                                b.cr_ri_batch_no
                                            AND r.rs_ri_cr_index =
                                                b.cr_ri_cr_index
                                            AND r.rs_type = 'Final'
                                            AND r.rs_line_type_int NOT IN
                                                    ('Balance', 'Retention')))
                               / 100)),
                          0))                    outward_recovery,
                  SUM (NVL (ROUND (CASE
                                       WHEN    pr_bus_type IN ('1000',
                                                               '1001',
                                                               '1002',
                                                               '2000',
                                                               '2999')
                                            OR pr_bus_type IS NULL
                                       THEN
                                           NVL (reserve_amnt, 0)
                                       ELSE
                                           0
                                   END,
                                   0),
                            0))                  direct,
                  SUM (
                      NVL (
                          ROUND (
                              CASE
                                  WHEN pr_bus_type IN ('3000')
                                  THEN
                                      NVL (reserve_amnt, 0)
                                  ELSE
                                      0
                              END,
                              0),
                          0))                    facin
    --   g.ch_status
    FROM cm_claims          a,
         cm_claims_risks    b,
         uw_premium_register f,
         (  SELECT eh_org_code,
                   eh_cm_index,
                   NVL (SUM (NVL (cm_closing_value, 0)), 0)     reserve_amnt
              FROM (  SELECT DISTINCT
                             d.eh_org_code,
                             d.eh_cm_index,
                             d.eh_ce_index,
                             d.eh_status,
                             NVL (d.eh_new_lc_amount, 0)     cm_closing_value
                        FROM cm_estimates_history d
                       WHERE     d.created_on =
                                 (SELECT DISTINCT MAX (g.created_on)
                                    FROM cm_estimates_history g
                                   WHERE     TRUNC (g.created_on) <=
                                             TRUNC ( :p_to_dt)
                                         AND g.eh_org_code = d.eh_org_code
                                         AND g.eh_cm_index = d.eh_cm_index
                                         AND g.eh_ce_index = d.eh_ce_index)
                             AND TRUNC (d.created_on) <= TRUNC ( :p_to_dt)
                             AND d.eh_status NOT IN ('Closed', 'Fully Paid')
                    ORDER BY d.eh_cm_index, d.eh_ce_index)
          GROUP BY eh_org_code, eh_cm_index) e,
         (SELECT DISTINCT a.ch_org_code, a.ch_cm_index, a.ch_status
            FROM cm_claims_history a
           WHERE a.created_on =
                 (SELECT DISTINCT MAX (b.created_on)
                    FROM cm_claims_history b
                   WHERE     TRUNC (b.created_on) <= TRUNC ( :p_to_dt)
                         AND b.ch_org_code = a.ch_org_code
                         AND b.ch_cm_index = a.ch_cm_index)) g
   WHERE     a.cm_org_code = :p_org_code
         AND a.cm_org_code = b.cr_org_code
         AND a.cm_index = b.cr_cm_index
         AND a.cm_org_code = f.pr_org_code(+)
         AND a.cm_pl_index = f.pr_pl_index(+)
         AND a.cm_end_index = f.pr_end_index(+)
         AND b.cr_org_code = e.eh_org_code(+)
         AND b.cr_cm_index = e.eh_cm_index(+)
         AND a.cm_org_code = g.ch_org_code
         AND a.cm_index = g.ch_cm_index
         AND g.ch_status NOT IN ('Closed', 'Closed - No Claim')
         AND a.cm_register = 'Y'
GROUP BY cm_org_code,
         cr_mc_code,
         INITCAP (
             CASE
                 WHEN cr_sc_code IN ('0804')
                 THEN
                     'PSV'
                 ELSE
                     pkg_system_admin.get_class_name (cm_org_code,
                                                      cr_mc_code)
             END),
         CASE
             WHEN cr_mc_code IN ('03',
                                 '04',
                                 '09',
                                 '11')
             THEN
                 1
             WHEN cr_mc_code IN ('070', '080')
             THEN
                 CASE
                     WHEN NVL (cm_pl_no, pr_pl_no) LIKE '%TP%' THEN 2
                     ELSE 1
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
             WHEN cr_mc_code IN ('03',
                                 '04',
                                 '09',
                                 '11')
             THEN
                 pkg_system_admin.get_class_name (cm_org_code, cr_mc_code)
             WHEN cr_mc_code IN ('070', '080')
             THEN
                 CASE
                     WHEN NVL (cm_pl_no, pr_pl_no) LIKE '%TP%'
                     THEN
                         'Third Party Only'
                     ELSE
                         'Comprehensive'
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
                         pkg_system_admin.get_subclass_name (cm_org_code,
                                                             cr_sc_code)
                     WHEN cr_sc_code IN ('120', '127', '128')
                     THEN
                         'Bonds'
                     ELSE
                         CASE
                             WHEN cr_mc_code IN ('10') THEN 'Burglary Others'
                             ELSE 'Others'
                         END
                 END
         END
UNION ALL
  SELECT DISTINCT 3                              order_no,
                  cm_org_code                    org_code,
                  cr_mc_code                     mc_code,
                  INITCAP (
                      CASE
                          WHEN cr_sc_code IN ('0804')
                          THEN
                              'PSV'
                          ELSE
                              pkg_system_admin.get_class_name (cm_org_code,
                                                               cr_mc_code)
                      END)                       class,
                  CASE
                      WHEN cr_mc_code IN ('03',
                                          '04',
                                          '09',
                                          '11')
                      THEN
                          1
                      WHEN cr_mc_code IN ('070', '080')
                      THEN
                          CASE
                              WHEN NVL (cm_pl_no, pr_pl_no) LIKE '%TP%' THEN 2
                              ELSE 1
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
                                                  '101',
                                                  '120',
                                                  '127',
                                                  '128')
                              THEN
                                  1
                              ELSE
                                  2
                          END
                  END                            cm_order,
                  CASE
                      WHEN cr_mc_code IN ('03',
                                          '04',
                                          '09',
                                          '11')
                      THEN
                          pkg_system_admin.get_class_name (cm_org_code,
                                                           cr_mc_code)
                      WHEN cr_mc_code IN ('070', '080')
                      THEN
                          CASE
                              WHEN NVL (cm_pl_no, pr_pl_no) LIKE '%TP%'
                              THEN
                                  'Third Party Only'
                              ELSE
                                  'Comprehensive'
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
                                  pkg_system_admin.get_subclass_name (
                                      cm_org_code,
                                      cr_sc_code)
                              WHEN cr_sc_code IN ('120', '127', '128')
                              THEN
                                  'Bonds'
                              ELSE
                                  CASE
                                      WHEN cr_mc_code IN ('10')
                                      THEN
                                          'Burglary Others'
                                      ELSE
                                          'Others'
                                  END
                          END
                  END                            sub_class,
                  SUM (NVL (reserve_amnt, 0))    total,
                  ROUND (
                      NVL (
                          SUM (
                              (  (  (NVL (reserve_amnt, 0))
                                  * (SELECT SUM (r.rs_percent)
                                       FROM uw_policy_ri_shares r
                                      WHERE     r.rs_org_code = b.cr_org_code
                                            AND r.rs_ri_batch_no =
                                                b.cr_ri_batch_no
                                            AND r.rs_ri_cr_index =
                                                b.cr_ri_cr_index
                                            AND r.rs_type = 'Final'
                                            AND r.rs_line_type_int NOT IN
                                                    ('Balance', 'Retention')))
                               / 100)),
                          0))                    outward_recovery,
                  SUM (NVL (ROUND (CASE
                                       WHEN    pr_bus_type IN ('1000',
                                                               '1001',
                                                               '1002',
                                                               '2000',
                                                               '2999')
                                            OR pr_bus_type IS NULL
                                       THEN
                                           NVL (reserve_amnt, 0)
                                       ELSE
                                           0
                                   END,
                                   0),
                            0))                  direct,
                  SUM (
                      NVL (
                          ROUND (
                              CASE
                                  WHEN pr_bus_type IN ('3000')
                                  THEN
                                      NVL (reserve_amnt, 0)
                                  ELSE
                                      0
                              END,
                              0),
                          0))                    facin
    --   g.ch_status
    FROM cm_claims          a,
         cm_claims_risks    b,
         uw_premium_register f,
         (  SELECT eh_org_code,
                   eh_cm_index,
                   NVL (SUM (NVL (cm_closing_value, 0)), 0)     reserve_amnt
              FROM (  SELECT DISTINCT
                             d.eh_org_code,
                             d.eh_cm_index,
                             d.eh_ce_index,
                             d.eh_status,
                             NVL (d.eh_new_lc_amount, 0)     cm_closing_value
                        FROM cm_estimates_history d
                       WHERE     d.created_on =
                                 (SELECT DISTINCT MAX (g.created_on)
                                    FROM cm_estimates_history g
                                   WHERE     TRUNC (g.created_on) <=
                                             TRUNC (
                                                   ADD_MONTHS (
                                                       TRUNC ( :p_to_dt, 'YEAR'),
                                                       -1)
                                                 + 30)
                                         AND g.eh_org_code = d.eh_org_code
                                         AND g.eh_cm_index = d.eh_cm_index
                                         AND g.eh_ce_index = d.eh_ce_index)
                             AND TRUNC (d.created_on) <=
                                 TRUNC (
                                       ADD_MONTHS (TRUNC ( :p_to_dt, 'YEAR'), -1)
                                     + 30)
                             AND d.eh_status NOT IN ('Closed', 'Fully Paid')
                    ORDER BY d.eh_cm_index, d.eh_ce_index)
          GROUP BY eh_org_code, eh_cm_index) e,
         (SELECT DISTINCT a.ch_org_code, a.ch_cm_index, a.ch_status
            FROM cm_claims_history a
           WHERE a.created_on =
                 (SELECT DISTINCT MAX (b.created_on)
                    FROM cm_claims_history b
                   WHERE     TRUNC (b.created_on) <=
                             TRUNC (
                                   ADD_MONTHS (TRUNC ( :p_to_dt, 'YEAR'), -1)
                                 + 30)
                         AND b.ch_org_code = a.ch_org_code
                         AND b.ch_cm_index = a.ch_cm_index)) g
   WHERE     a.cm_org_code = :p_org_code
         AND a.cm_org_code = b.cr_org_code
         AND a.cm_index = b.cr_cm_index
         AND a.cm_org_code = f.pr_org_code(+)
         AND a.cm_pl_index = f.pr_pl_index(+)
         AND a.cm_end_index = f.pr_end_index(+)
         AND b.cr_org_code = e.eh_org_code(+)
         AND b.cr_cm_index = e.eh_cm_index(+)
         AND a.cm_org_code = g.ch_org_code
         AND a.cm_index = g.ch_cm_index
         AND g.ch_status NOT IN ('Closed', 'Closed - No Claim')
         AND a.cm_register = 'Y'
GROUP BY cm_org_code,
         cr_mc_code,
         INITCAP (
             CASE
                 WHEN cr_sc_code IN ('0804')
                 THEN
                     'PSV'
                 ELSE
                     pkg_system_admin.get_class_name (cm_org_code,
                                                      cr_mc_code)
             END),
         CASE
             WHEN cr_mc_code IN ('03',
                                 '04',
                                 '09',
                                 '11')
             THEN
                 1
             WHEN cr_mc_code IN ('070', '080')
             THEN
                 CASE
                     WHEN NVL (cm_pl_no, pr_pl_no) LIKE '%TP%' THEN 2
                     ELSE 1
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
             WHEN cr_mc_code IN ('03',
                                 '04',
                                 '09',
                                 '11')
             THEN
                 pkg_system_admin.get_class_name (cm_org_code, cr_mc_code)
             WHEN cr_mc_code IN ('070', '080')
             THEN
                 CASE
                     WHEN NVL (cm_pl_no, pr_pl_no) LIKE '%TP%'
                     THEN
                         'Third Party Only'
                     ELSE
                         'Comprehensive'
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
                         pkg_system_admin.get_subclass_name (cm_org_code,
                                                             cr_sc_code)
                     WHEN cr_sc_code IN ('120', '127', '128')
                     THEN
                         'Bonds'
                     ELSE
                         CASE
                             WHEN cr_mc_code IN ('10') THEN 'Burglary Others'
                             ELSE 'Others'
                         END
                 END
         END
ORDER BY
    1,
    2,
    3,
    5`;
      const results = (await connection).execute(query, {
        p_org_code: "50",
        p_class: "",
        p_subclass: "",
        p_fm_dt: new Date(fromDate),
        p_to_dt: new Date(toDate),
      });
      const finalResults = formatOracleData(await results);
      //initiate the workbook or the excel package
      const updateWorkbook = (workbook) => {
        const worksheet = workbook.getWorksheet("59-3B");

        Object.entries(classSubclassRowMapper3).forEach(
          ([classSubKey, targetRow]) => {
            const [classKey, subClassKey] = classSubKey.split("|");

            const filteredResults = finalResults.filter(
              (item) =>
                item.CLASS === classKey && item.SUB_CLASS === subClassKey
            );
            if (filteredResults.length > 0) {
              filteredResults.forEach((dataItem) => {
                for (const [field, column] of Object.entries(cellMapper4)) {
                  const cell = worksheet.getCell(`${column}${targetRow}`);
                  cell.value += dataItem[field];
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
