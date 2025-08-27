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

namespace TKBUSINESS
{
    public partial class frmREPORTSCOPTGTH : Form
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
        Report report1 = new Report();

        string tablename = null;
        int rownum = 0;

        public frmREPORTSCOPTGTH()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT(string SDATES, string EDATES)
        {
            string P1 = SDATES;
            string P2 = EDATES;

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);


            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            Report report1 = new Report();

            SQL1 = SETSQL_SETFASTREPORT(SDATES, EDATES);
            report1.Load(@"REPORT\銷貨單理貨表.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();
            
            report1.SetParameterValue("P1", P1);
            report1.SetParameterValue("P2", P2);

            report1.Preview = previewControl4;
            report1.Show();
        }

        public StringBuilder SETSQL_SETFASTREPORT(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SBQUERY = new StringBuilder();
            StringBuilder SBQUERY2 = new StringBuilder();

          

            SB.AppendFormat(@" 
                           SELECT 
                            TH004 AS '品號'
                            ,TH005 AS '品名'
                            ,TH006 AS '規格'
                            ,TH007 AS '庫別ID'
                            ,MC002 AS '庫別'
                            ,TH017 AS '批號'
                            ,SUM(TH008+TH024) AS '數量'
                            FROM [TK].dbo.COPTG,[TK].dbo.COPTH,[TK].dbo.CMSMC
                            WHERE TG001=TH001 AND TG002=TH002
                            AND MC001=TH007
                            AND TH004 NOT LIKE '599%' 
                            AND TG001 IN ('A233','A234')
                            AND TG003>='{0}' AND TG003<='{1}'
                            GROUP BY TH004,TH005,TH006,TH007,MC002,TH017
                            ORDER BY MC002,TH004,TH005,TH006,TH007
 
                            ", SDATES, EDATES);

            return SB;

        }
        public void SETFASTREPORT_DETAILS(string SDATES, string EDATES)
        {
            string P1 = SDATES;
            string P2 = EDATES;

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);


            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            Report report1 = new Report();

            SQL1 = SETSQL_SETFASTREPORT_DETAILS(SDATES, EDATES);
            report1.Load(@"REPORT\銷貨單理貨明細表.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.SetParameterValue("P1", P1);
            report1.SetParameterValue("P2", P2);

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL_SETFASTREPORT_DETAILS(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SBQUERY = new StringBuilder();
            StringBuilder SBQUERY2 = new StringBuilder();



            SB.AppendFormat(@"                            
                            SELECT 
                            TG003 AS '出貨日'
                            ,TH001 AS '單別'
                            ,TH002 AS '單號'
                            ,TH003 AS '序號'
                            ,TH004 AS '品號'
                            ,TH005 AS '品名'
                            ,TH006 AS '規格'
                            ,TH007 AS '庫別ID'
                            ,MC002 AS '庫別'
                            ,TH017 AS '批號'
                            ,(TH008+TH024) AS '數量'
                            FROM [TK].dbo.COPTG,[TK].dbo.COPTH,[TK].dbo.CMSMC
                            WHERE TG001=TH001 AND TG002=TH002
                            AND MC001=TH007
                            AND TH004 NOT LIKE '599%'
                            AND TG001 IN ('A233','A234')
                            AND TG003>='{0}' AND TG003<='{1}'
                            ORDER BY TG003,TH001,TH002,TH003

                            ", SDATES, EDATES);

            return SB;

        }

        #endregion

        #region BUTTON


        private void button1_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");
            string EDATES = dateTimePicker2.Value.ToString("yyyyMMdd");

            SETFASTREPORT(SDATES, EDATES);

            SETFASTREPORT_DETAILS(SDATES, EDATES);
        }
        #endregion
    }
}
