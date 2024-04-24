using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Media.Effects;
using OfficeOpenXml;

namespace BatchALPHAPartMasterUpdate
{
    public class Program
    {
        static void Main(string[] args)
        {
            ConnectDB oConnBCS = new ConnectDB("DBBCS");
            ConnectDB oConnSCM = new ConnectDB("DBSCM");

            Console.WriteLine("START - UPDATE PART MASTER TO DCI DATABASE. " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            #region
            string source_file = @"D:\RPA\051\051PARTMASTER.xlsx";
            string path = @"C:\\temp\";
            string filename = DateTime.Now.ToString("yyyyMMddHHmmss")+".xlsx";

            if (File.Exists(source_file))
            {
                if (!Directory.Exists(path))
                {
                    Console.WriteLine("CREATE DIRECTORY FOR TEMP FILE.");
                    Directory.CreateDirectory(path);
                }
                else
                {
                    if (File.Exists(path + filename))
                    {
                        Console.WriteLine("DELETE OLD TEMP FILE.");
                        File.Delete(path + filename);
                    }
                    else
                    {
                        Console.WriteLine("CREATE TEMP FILE.");
                        File.Copy(source_file, path + filename);
                    }
                }


                // DELETE OLD DATA
                SqlCommand sqlDelete1 = new SqlCommand();
                sqlDelete1.CommandText = @"DELETE FROM AL_Part";
                sqlDelete1.CommandTimeout = 180;
                oConnSCM.ExecuteCommand(sqlDelete1);
                Console.WriteLine("DELETED OLD DATA IN DATABASE. (AL_PART)");

                int endRowIndex = 10000;
                byte[] bin = File.ReadAllBytes(path + filename);

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (MemoryStream stream = new MemoryStream(bin))
                using (ExcelPackage excelPackage = new ExcelPackage(stream))
                {
                    ExcelWorksheet xlWorkSheet = excelPackage.Workbook.Worksheets[0];
                    endRowIndex = xlWorkSheet.Dimension.End.Row;
                    for (int row = 8; row <= endRowIndex; row++)
                    {
                        if (xlWorkSheet.Cells[row, 1].Value != null)
                        {
                            string p_DrawingNo = "";
                            string p_CM = "";
                            string p_Description = "";
                            string p_Route = "";
                            string p_CATMAT = "";
                            string p_OrderType = "";
                            string p_WHUnit = "";
                            string p_CnvCode = "";
                            decimal p_CnvWT = 0;
                            string p_IVUnit = "";
                            decimal p_QtyBox = 0;
                            string p_VenderCode = "";
                            string p_VenderName = "";
                            int p_LeadTime = 0;
                            int p_PdLeadTime = 0;
                            string p_Remark1 = "";

                            try { p_DrawingNo = xlWorkSheet.Cells[row, 2].Value.ToString().Trim(); } catch { }
                            try { p_CM = xlWorkSheet.Cells[row, 5].Value.ToString().Trim(); } catch { }
                            try { p_Description = xlWorkSheet.Cells[row, 4].Value.ToString().Trim(); } catch { }
                            try { p_Route = xlWorkSheet.Cells[row, 10].Value.ToString().Trim(); } catch { }
                            try { p_CATMAT = xlWorkSheet.Cells[row, 13].Value.ToString().Trim(); } catch { }
                            try { p_OrderType = xlWorkSheet.Cells[row, 14].Value.ToString().Trim(); } catch { }
                            try { p_WHUnit = xlWorkSheet.Cells[row, 9].Value.ToString().Trim(); } catch { }
                            try { p_CnvCode = xlWorkSheet.Cells[row, 25].Value.ToString().Trim(); } catch { }
                            try {
                                if (xlWorkSheet.Cells[row, 26].Value != null)
                                {
                                    p_CnvWT = Convert.ToDecimal(xlWorkSheet.Cells[row, 26].Value.ToString().Trim());
                                }
                                else
                                {
                                    p_CnvWT = 0;
                                }
                            } catch {
                                
                            }
                            try { p_IVUnit = xlWorkSheet.Cells[row, 27].Value.ToString().Trim(); } catch { }
                            try {
                                if (xlWorkSheet.Cells[row, 19].Value != null)
                                {
                                    p_QtyBox = Convert.ToDecimal(xlWorkSheet.Cells[row, 19].Value.ToString().Trim());
                                }
                                else
                                {
                                    p_QtyBox = 0;
                                }
                                
                            } catch { }
                            try { p_VenderCode = xlWorkSheet.Cells[row, 11].Value.ToString().Trim(); } catch { }
                            try { p_VenderName = xlWorkSheet.Cells[row, 12].Value.ToString().Trim(); } catch { }
                            try {
                                if (xlWorkSheet.Cells[row, 15].Value != null)
                                {
                                    p_LeadTime = Convert.ToInt16(xlWorkSheet.Cells[row, 15].Value.ToString().Trim());
                                }
                                else
                                {
                                    p_LeadTime = 0;
                                }
                            } catch { }
                            try {
                                if (xlWorkSheet.Cells[row, 16].Value != null)
                                {
                                    p_PdLeadTime = Convert.ToInt16(xlWorkSheet.Cells[row, 16].Value.ToString().Trim());
                                }
                                else
                                {
                                    p_PdLeadTime = 0;
                                }
                            } catch { }
                            try { p_Remark1 = xlWorkSheet.Cells[row, 2].Value.ToString().Trim(); } catch { }

                            SqlCommand sqlSelect1 = new SqlCommand();
                            sqlSelect1.CommandText = @"SELECT * FROM AL_Part WHERE DrawingNo = @DrawingNo";
                            sqlSelect1.Parameters.Add(new SqlParameter("@DrawingNo", p_DrawingNo));
                            sqlSelect1.CommandTimeout = 180;
                            DataTable dtPART = oConnSCM.Query(sqlSelect1);
                            if (dtPART.Rows.Count == 0)
                            {
                                SqlCommand sqlInsert = new SqlCommand();
                                sqlInsert.CommandText = @"INSERT INTO AL_Part (DrawingNo, CM, Description, Route, OrderType, CATMAT, WHUnit, CnvCode, CnvWT, IVUnit, QtyBox, VenderCode, VenderName, LeadTime, PdLeadTime, UpdateDate, UpdateBy)
                                    VALUES (@DrawingNo, @CM, @Description, @Route, @OrderType, @CATMAT, @WHUnit, @CnvCode, @CnvWT, 
                                        @IVUnit, @QtyBox, @VenderCode, @VenderName, @LeadTime, @PdLeadTime, @UpdateDate, @UpdateBy)";
                                sqlInsert.Parameters.Add(new SqlParameter("@DrawingNo", p_DrawingNo));
                                sqlInsert.Parameters.Add(new SqlParameter("@CM", p_CM));
                                sqlInsert.Parameters.Add(new SqlParameter("@Description", p_Description));
                                sqlInsert.Parameters.Add(new SqlParameter("@Route", p_Route));
                                sqlInsert.Parameters.Add(new SqlParameter("@OrderType", p_OrderType));
                                sqlInsert.Parameters.Add(new SqlParameter("@CATMAT", p_CATMAT));
                                sqlInsert.Parameters.Add(new SqlParameter("@WHUnit", p_WHUnit));
                                sqlInsert.Parameters.Add(new SqlParameter("@CnvCode", p_CnvCode));
                                sqlInsert.Parameters.Add(new SqlParameter("@CnvWT", p_CnvWT));
                                sqlInsert.Parameters.Add(new SqlParameter("@IVUnit", p_IVUnit));
                                sqlInsert.Parameters.Add(new SqlParameter("@QtyBox", p_QtyBox));
                                sqlInsert.Parameters.Add(new SqlParameter("@VenderCode", p_VenderCode));
                                sqlInsert.Parameters.Add(new SqlParameter("@VenderName", p_VenderName));
                                sqlInsert.Parameters.Add(new SqlParameter("@LeadTime", p_LeadTime));
                                sqlInsert.Parameters.Add(new SqlParameter("@PdLeadTime", p_PdLeadTime));
                                sqlInsert.Parameters.Add(new SqlParameter("@UpdateDate", DateTime.Now));
                                sqlInsert.Parameters.Add(new SqlParameter("@UpdateBy", "BATCH"));
                                sqlInsert.CommandTimeout = 180;
                                oConnSCM.ExecuteCommand(sqlInsert);
                            }

                            decimal progress = ((Convert.ToDecimal(row) / Convert.ToDecimal(endRowIndex)) * 100);
                            Console.WriteLine("SAVED " + progress.ToString("N2") + "%");
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("NO PART MASTER EXCEL FILE.");
            }
            #endregion
            Console.WriteLine("END - UPDATE PART MASTER TO DCI DATABASE. " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            Console.WriteLine("START - UPDATE MODELS TO AL_PART. " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            #region
            SqlCommand sqlSelectB = new SqlCommand();
            sqlSelectB.CommandText = @"SELECT * FROM AL_Part 
                WHERE (DrawingNo NOT LIKE '1Y%' AND DrawingNo NOT LIKE '2Y%' AND DrawingNo NOT LIKE 'J%' AND DrawingNo NOT LIKE 'M%'
                    AND DrawingNo NOT LIKE 'D%' AND DrawingNo NOT LIKE 'Q%')
                ORDER BY DrawingNo ASC";
            sqlSelectB.CommandTimeout = 180;
            DataTable dtPart = oConnSCM.Query(sqlSelectB);
            if (dtPart.Rows.Count > 0)
            {
                int count = 0;
                foreach (DataRow drow in dtPart.Rows)
                {
                    string _partno = drow["DrawingNo"].ToString();

                    SqlCommand sqlSelectModel = new SqlCommand();
                    sqlSelectModel.CommandText = @"SELECT t.PARTNO, SUBSTRING(STRING_AGG(Models + ',',''),1, LEN(STRING_AGG(Models + ',',''))-1) as Models FROM (
                            SELECT PARTNO, SUBSTRING([MODEL], 1,8) Models 
                            FROM [RES_PART_LIST]
                             WHERE YM = @YM AND PARTNO = @PARTNO
                             GROUP BY PARTNO, SUBSTRING([MODEL], 1, 8)
                         ) as t
	                     group by t.PARTNO";
                    sqlSelectModel.Parameters.Add(new SqlParameter("@YM", DateTime.Now.ToString("yyyyMM")));
                    sqlSelectModel.Parameters.Add(new SqlParameter("@PARTNO", _partno));
                    sqlSelectModel.CommandTimeout = 180;
                    DataTable dtModels = oConnBCS.Query(sqlSelectModel);
                    if (dtModels.Rows.Count > 0)
                    {
                        string models = dtModels.Rows[0]["Models"].ToString();

                        if (models.Length >= 255)
                        {
                            models = models.Substring(0, 255);
                        }

                        SqlCommand sqlUpdate = new SqlCommand();
                        sqlUpdate.CommandText = "UPDATE AL_Part SET Remark1 = @MODEL WHERE DrawingNo = @PARTNO";
                        sqlUpdate.Parameters.Add(new SqlParameter("@MODEL", models));
                        sqlUpdate.Parameters.Add(new SqlParameter("@PARTNO", _partno));
                        sqlUpdate.CommandTimeout = 180;
                        oConnSCM.ExecuteCommand(sqlUpdate);
                    }
                    else
                    {
                        SqlCommand sqlSelectModel2 = new SqlCommand();
                        sqlSelectModel2.CommandText = @"SELECT t.CHILD_PART, SUBSTRING(STRING_AGG(Models + ',',''),1, LEN(STRING_AGG(Models + ',',''))-1) as Models FROM (
                            SELECT CHILD_PART, SUBSTRING([MODEL], 1,8) Models 
                            FROM [CST_PRD_STRUCTURE]
                             WHERE YM = @YM AND CHILD_PART = @PARTNO
                             GROUP BY CHILD_PART, SUBSTRING([MODEL], 1, 8)
                         ) as t
	                     group by t.CHILD_PART";
                        sqlSelectModel2.Parameters.Add(new SqlParameter("@YM", DateTime.Now.ToString("yyyyMM")));
                        sqlSelectModel2.Parameters.Add(new SqlParameter("@PARTNO", _partno));
                        sqlSelectModel2.CommandTimeout = 180;
                        DataTable dtModel2 = oConnBCS.Query(sqlSelectModel2);
                        if (dtModel2.Rows.Count > 0)
                        {
                            string models = dtModel2.Rows[0]["Models"].ToString();

                            if (models.Length >= 255)
                            {
                                models = models.Substring(0, 255);
                            }

                            SqlCommand sqlUpdate = new SqlCommand();
                            sqlUpdate.CommandText = "UPDATE AL_Part SET Remark1 = @MODEL WHERE DrawingNo = @PARTNO";
                            sqlUpdate.Parameters.Add(new SqlParameter("@MODEL", models));
                            sqlUpdate.Parameters.Add(new SqlParameter("@PARTNO", _partno));
                            sqlUpdate.CommandTimeout = 180;
                            oConnSCM.ExecuteCommand(sqlUpdate);
                        }
                    }

                    decimal progress = ((Convert.ToDecimal(count) / Convert.ToDecimal(dtPart.Rows.Count)) * 100);
                    count++;
                    Console.WriteLine("SAVED " + progress.ToString("N2") + "%");
                }
            }
            #endregion
            Console.WriteLine("END - UPDATE MODELS TO AL_PART. " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            Console.WriteLine("DELETE OLD TEMP FILE.");
            File.Delete(path + filename);
        }
    }
}
