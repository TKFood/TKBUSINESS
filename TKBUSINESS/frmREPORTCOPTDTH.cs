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

namespace TKBUSINESS
{
    public partial class frmREPORTCOPTDTH : Form
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

        public frmREPORTCOPTDTH()
        {
            InitializeComponent();
        }


        #region FUNCTION
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();

            if(comboBox2.Text.Equals("國內"))
            {
                report1.Load(@"REPORT\訂單達交報表-國內.frx");
            }
            else if (comboBox2.Text.Equals("國外"))
            {
                report1.Load(@"REPORT\訂單達交報表-國外.frx");
            }


            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
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

            if (comboBox2.Text.Equals("國內"))
            {
                if (comboBox1.Text.Equals("未準時出貨"))
                {
                    SB.AppendFormat(" SELECT TC053 AS '客戶',TD013 AS '預交日',TD001 AS '單別',TD002 AS '單號',TD003 AS '序號',TD004 AS '品號',TD005 AS '品名',TD008 AS '訂單數量',TD024 AS '贈品量',TD010 AS '單位'");
                    SB.AppendFormat(" ,(SELECT ISNULL(SUM(TH008),0) FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' AND TG003<=DATEADD(day,5,TD013)) AS '預交日前的已交數量'");
                    SB.AppendFormat(" ,(SELECT ISNULL(SUM(TH008),0) FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' ) AS '總已交數量'");
                    SB.AppendFormat(" ,(SELECT TOP 1 ISNULL(TG003,'') FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' ) AS '銷貨的第一天'");
                    SB.AppendFormat(" ,(SELECT ISNULL(SUM(TH024),0) FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' AND TG003<=DATEADD(day,5,TD013)) AS '預交日前的贈品已交量'");
                    SB.AppendFormat(" ,(SELECT ISNULL(SUM(TH024),0) FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' ) AS '總贈品已交量'");
                    SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
                    SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
                    SB.AppendFormat(" AND (TD004 LIKE '4%' OR TD004 LIKE '5%')");
                    SB.AppendFormat(" AND TD021='Y'");
                    SB.AppendFormat(" ");
                    SB.AppendFormat(" AND TD001 IN ('A221')");
                    SB.AppendFormat(" AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    SB.AppendFormat(" AND ((TD008>(SELECT ISNULL(SUM(TH008),0) FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' AND TG003<=DATEADD(day,5,TD013))))");
                    SB.AppendFormat(" ORDER BY TC001,TC053,TD013,TD005");
                    SB.AppendFormat(" ");
                    SB.AppendFormat(" ");
                }
                else
                {
                    SB.AppendFormat(" SELECT TC053 AS '客戶',TD013 AS '預交日',TD001 AS '單別',TD002 AS '單號',TD003 AS '序號',TD004 AS '品號',TD005 AS '品名',TD008 AS '訂單數量',TD024 AS '贈品量',TD010 AS '單位'");
                    SB.AppendFormat(" ,(SELECT ISNULL(SUM(TH008),0) FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' AND TG003<=DATEADD(day,5,TD013)) AS '預交日前的已交數量'");
                    SB.AppendFormat(" ,(SELECT ISNULL(SUM(TH008),0) FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' ) AS '總已交數量'");
                    SB.AppendFormat(" ,(SELECT TOP 1 ISNULL(TG003,'') FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' ) AS '銷貨的第一天'");
                    SB.AppendFormat(" ,(SELECT ISNULL(SUM(TH024),0) FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' AND TG003<=DATEADD(day,5,TD013)) AS '預交日前的贈品已交量'");
                    SB.AppendFormat(" ,(SELECT ISNULL(SUM(TH024),0) FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' ) AS '總贈品已交量'");
                    SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
                    SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
                    SB.AppendFormat(" AND (TD004 LIKE '4%' OR TD004 LIKE '5%')");
                    SB.AppendFormat(" AND TD021='Y'");
                    SB.AppendFormat(" ");
                    SB.AppendFormat(" AND TD001 IN ('A221')");
                    SB.AppendFormat(" AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));                   
                    SB.AppendFormat(" ORDER BY TC001,TC053,TD013,TD005");
                    SB.AppendFormat(" ");
                    SB.AppendFormat(" ");

                }
            }

            else if (comboBox2.Text.Equals("國外"))
            {
                if (comboBox1.Text.Equals("未準時出貨"))
                {
                    SB.AppendFormat(" SELECT TC053 AS '客戶',TD013 AS '預交日',TD001 AS '單別',TD002 AS '單號',TD003 AS '序號',TD004 AS '品號',TD005 AS '品名',TD008 AS '訂單數量',TD024 AS '贈品量',TD010 AS '單位'");
                    SB.AppendFormat(" ,(SELECT ISNULL(SUM(TH008),0) FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' AND TG003<=DATEADD(day,7,TD013)) AS '預交日前的已交數量'");
                    SB.AppendFormat(" ,(SELECT ISNULL(SUM(TH008),0) FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' ) AS '總已交數量'");
                    SB.AppendFormat(" ,(SELECT TOP 1 ISNULL(TG003,'') FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' ) AS '銷貨的第一天'");
                    SB.AppendFormat(" ,(SELECT ISNULL(SUM(TH024),0) FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' AND TG003<=DATEADD(day,7,TD013)) AS '預交日前的贈品已交量'");
                    SB.AppendFormat(" ,(SELECT ISNULL(SUM(TH024),0) FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' ) AS '總贈品已交量'");
                    SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
                    SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
                    SB.AppendFormat(" AND (TD004 LIKE '4%' OR TD004 LIKE '5%')");
                    SB.AppendFormat(" AND TD021='Y'");
                    SB.AppendFormat(" ");
                    SB.AppendFormat(" AND TD001 IN ('A222')");
                    SB.AppendFormat(" AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    SB.AppendFormat(" AND ((TD008>(SELECT ISNULL(SUM(TH008),0) FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' AND TG003<=DATEADD(day,7,TD013))))");
                    SB.AppendFormat(" ORDER BY TC001,TC053,TD013,TD005");
                    SB.AppendFormat(" ");
                    SB.AppendFormat(" ");
                }
                else
                {
                    SB.AppendFormat(" SELECT TC053 AS '客戶',TD013 AS '預交日',TD001 AS '單別',TD002 AS '單號',TD003 AS '序號',TD004 AS '品號',TD005 AS '品名',TD008 AS '訂單數量',TD024 AS '贈品量',TD010 AS '單位'");
                    SB.AppendFormat(" ,(SELECT ISNULL(SUM(TH008),0) FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' AND TG003<=DATEADD(day,7,TD013)) AS '預交日前的已交數量'");
                    SB.AppendFormat(" ,(SELECT ISNULL(SUM(TH008),0) FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' ) AS '總已交數量'");
                    SB.AppendFormat(" ,(SELECT TOP 1 ISNULL(TG003,'') FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' ) AS '銷貨的第一天'");
                    SB.AppendFormat(" ,(SELECT ISNULL(SUM(TH024),0) FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' AND TG003<=DATEADD(day,7,TD013)) AS '預交日前的贈品已交量'");
                    SB.AppendFormat(" ,(SELECT ISNULL(SUM(TH024),0) FROM [TK].dbo.COPTH,[TK].dbo.COPTG WHERE TG001=TH001 AND TG002=TH002 AND  TH014=TD001 AND TH015=TD002 AND TH016=TD003 AND TH020='Y' ) AS '總贈品已交量'");
                    SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
                    SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
                    SB.AppendFormat(" AND (TD004 LIKE '4%' OR TD004 LIKE '5%')");
                    SB.AppendFormat(" AND TD021='Y'");
                    SB.AppendFormat(" ");
                    SB.AppendFormat(" AND TD001 IN ('A222')");
                    SB.AppendFormat(" AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    SB.AppendFormat(" ORDER BY TC001,TC053,TD013,TD005");
                    SB.AppendFormat(" ");
                    SB.AppendFormat(" ");

                }
            }

            
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");

            return SB;

        }

        #endregion

        #region BUTTON
        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        #endregion
    }
}
