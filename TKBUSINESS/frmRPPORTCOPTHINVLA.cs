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
    public partial class frmRPPORTCOPTHINVLA : Form
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

        public frmRPPORTCOPTHINVLA()
        {
            InitializeComponent();

            SETDATES();
            SETPCTS();
        }


        #region FUNCTION

        public void SETDATES()
        {
            DateTime FirstDay = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime LastDay = new DateTime(DateTime.Now.AddMonths(1).Year, DateTime.Now.AddMonths(1).Month, 1).AddDays(-1);

            dateTimePicker1.Value = FirstDay;
            dateTimePicker2.Value = LastDay;


        }

        public void SETPCTS()
        {
            textBox1.Text = "40";
            textBox2.Text = "20";
        }
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();
            report1.Load(@"REPORT\2021銷貨單價+成本+利潤.frx");

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

            SB.AppendFormat(@" 
                            SELECT TG006 AS '部門代',ME002 AS '部門',TG005 AS '業務',MV002 AS '業務員',TG004 AS '客代',TG007 AS '客戶',TH004 AS '品號',TH005 AS '品名',CONVERT(DECIMAL(16,2),AVG(TH037/LA011)) AS '平均銷貨單價',AVG(LA012) AS '平均成本',AVG(LA012)*{0}*{1} AS '目標成本利潤',CONVERT(DECIMAL(16,2),(AVG(TH037/LA011)-(AVG(LA012)))) AS '單價成本差',CONVERT(DECIMAL(16,2),(AVG(TH037/LA011)-(AVG(LA012)*{0}*{1}))) AS '目標利潤單價成本差'
                            FROM(
                            SELECT TG001,TG002,TG006,ME002,TG005,MV002,TG004,TG007,TH004,TH005,TH037,LA011,LA012
                            FROM [TK].dbo.COPTG,[TK].dbo.COPTH,[TK].dbo.INVLA,[TK].dbo.CMSMV,[TK].dbo.CMSME
                            WHERE TG001=TH001 AND TG002=TH002 
                            AND LA006=TH001 AND LA007=TH002 AND LA008=TH003
                            AND TG006=MV001
                            AND TG005=ME001
                            AND (TH004 LIKE '4%' OR TH004 LIKE '5%')
                            AND (TG004 LIKE '2%' OR TG004 LIKE '3%' OR TG004 LIKE 'A%' OR TG004 LIKE 'B%')
                            AND TH037>0
                            AND LA011>0
                            AND TG003>='{2}' AND TG003<='{3}'
                            ) AS TEMP
                            GROUP BY TG006,ME002,TG005,MV002,TG004,TG007,TH004,TH005
                            ORDER BY CONVERT(DECIMAL(16,2),(AVG(TH037/LA011)-(AVG(LA012)*{0}*{1}))),ME002,TG005
                            ",((Convert.ToDecimal(textBox1.Text)+100)/100), ((Convert.ToDecimal(textBox2.Text) + 100) / 100),dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));

            return SB;

        }



        #endregion

        #region BUTTON



        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        #endregion
    }
}
