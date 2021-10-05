using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.Extractor;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using FastReport;
using FastReport.Data;
using TKITDLL;
using System.Data.OleDb;

namespace TKBUSINESS
{
    public partial class frmREPORTCLIENTCODES : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        SqlTransaction tran;
        int result;

        public frmREPORTCLIENTCODES()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void InsertExcelRecords(string FilePathName)
        {
            try
            {
                DELETETBCLIENTCODES();

                string conString = string.Empty;
                string extension = Path.GetExtension(FilePathName);
                switch (extension)
                {
                    case ".xls": //Excel 97-03
                        conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'";
                        break;
                    case ".xlsx": //Excel 07 or higher
                        conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'";
                        break;

                }

                conString = string.Format(conString, FilePathName);
                using (OleDbConnection excel_con = new OleDbConnection(conString))
                {
                    excel_con.Open();
                    string sheet1 = excel_con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["TABLE_NAME"].ToString();
                    DataTable dtExcelData = new DataTable();

                    //[OPTIONAL]: It is recommended as otherwise the data will be considered as String by default.
                    //dtExcelData.Columns.AddRange(new DataColumn[1] { new DataColumn("CODES", typeof(string)) });

                    using (OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [" + sheet1 + "]", excel_con))
                    {
                        oda.Fill(dtExcelData);
                    }
                    excel_con.Close();

                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);
                   
                    using (SqlConnection con = new SqlConnection(sqlsb.ConnectionString))
                    {
                        using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                        {
                            //Set the database table name
                            sqlBulkCopy.DestinationTableName = "[TKBUSINESS].[dbo].[TBCLIENTCODES]";

                            //[OPTIONAL]: Map the Excel columns with that of the database table
                            sqlBulkCopy.ColumnMappings.Add("CODE", "CODE");
               
                            con.Open();
                            sqlBulkCopy.WriteToServer(dtExcelData);
                            con.Close();

                            MessageBox.Show("完成");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Data has not been Imported due to :{0}", ex.Message), "Not Imported", MessageBoxButtons.OK, MessageBoxIcon.Warning);
           
            }


        
        }

        public void DELETETBCLIENTCODES()
        {
            try
            {

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                
                sbSql.AppendFormat(@"   
                                    DELETE [TKBUSINESS].[dbo].[TBCLIENTCODES]
                                    ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void SETFASTREPORT()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);



            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();

            report1.Load(@"REPORT\銷貨單號比對.frx");


            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            StringBuilder SB = new StringBuilder();          

            SB.AppendFormat(@"
                            SELECT [TBCLIENTCODES].CODE AS '單號',TG014 AS '發票號碼',TG020 AS '備註',TH001 AS '銷貨單',TH002 AS '銷貨單號',TH003 AS '銷貨序號',TH004 AS '品號',TH005 AS '品名',(TH008+TH024) AS '數量'
                            FROM [TKBUSINESS].[dbo].[TBCLIENTCODES],[TK].dbo.COPTG,[TK].dbo.COPTH
                            WHERE TG001=TH001 AND TG002=TH002
                            AND TG020 LIKE '%'+[TBCLIENTCODES].CODE+'%'
                            ORDER BY [TBCLIENTCODES].CODE,TG020,TH001,TH002,TH003,TH004
                            ");

            return SB;

        }
        #endregion

        #region BUTTON
        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select file";
            dialog.InitialDirectory = ".\\";
            dialog.Filter = "xls files (*.*)|*.xls|xlsx files (*.*)|*.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                InsertExcelRecords(dialog.FileName);
                //MessageBox.Show(dialog.FileName);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        #endregion


    }
}
