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
    public partial class frmRECOPTH : Form
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

        public frmRECOPTH()
        {
            InitializeComponent();
        }

        #region FUNCTION
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
            report1.Load(@"REPORT\銷售月報表-客戶.frx");

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

            SB.AppendFormat(" SELECT SUBSTRING(TG003,1,6) AS 'YM',TG004,TG007,TH004,TH005,SUM(TH037) AS 'MONEY',SUM(LA011) AS 'NUM',MB004");
            SB.AppendFormat(" FROM [TK].dbo.COPTG,[TK].dbo.COPTH,[TK].dbo.INVLA,[TK].dbo.INVMB");
            SB.AppendFormat(" WHERE TG001=TH001 AND TG002=TH002");
            SB.AppendFormat(" AND LA006=TH001 AND LA007=TH002 AND LA008=TH003");
            SB.AppendFormat(" AND TH004=MB001");
            SB.AppendFormat(" AND TG023='Y'");
            SB.AppendFormat(" AND (TH004 LIKE '4%' OR TH004 LIKE '5%' )");
            SB.AppendFormat(" AND TG003>='{0}' AND TG003<='{1}'", dateTimePicker1.Value.ToString("yyyyMM")+"01", dateTimePicker2.Value.ToString("yyyyMM") + "31");
            SB.AppendFormat(" AND TG004 IN ('{0}')",textBox1.Text);
            SB.AppendFormat(" GROUP BY SUBSTRING(TG003,1,6),TG004,TG007,TH004,TH005,MB004");
            SB.AppendFormat(" ORDER BY SUBSTRING(TG003,1,6),TH004");
            SB.AppendFormat("     ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");

            return SB;

        }


        public void SETFASTREPORT2()
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

            SQL1 = SETSQL2();
            Report report2 = new Report();
            report2.Load(@"REPORT\銷售月報表-業務.frx");

            report2.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report2.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report2.Preview = previewControl2;
            report2.Show();
        }

        public StringBuilder SETSQL2()
        {
            StringBuilder SB = new StringBuilder();

          
            SB.AppendFormat(" SELECT TG006,SUBSTRING(TG003,1,6) AS 'YM',TH004,TH005,SUM(TH037) AS 'MONEY',SUM(LA011) AS 'NUM',MB004");
            SB.AppendFormat(" FROM [TK].dbo.COPTG,[TK].dbo.COPTH,[TK].dbo.INVLA,[TK].dbo.INVMB");
            SB.AppendFormat(" WHERE TG001=TH001 AND TG002=TH002");
            SB.AppendFormat(" AND LA006=TH001 AND LA007=TH002 AND LA008=TH003");
            SB.AppendFormat(" AND TH004=MB001");
            SB.AppendFormat(" AND TG023='Y'");
            SB.AppendFormat(" AND (TH004 LIKE '4%' OR TH004 LIKE '5%' )");
            SB.AppendFormat(" AND TG003>='{0}' AND TG003<='{1}'", dateTimePicker3.Value.ToString("yyyyMM") + "01", dateTimePicker4.Value.ToString("yyyyMM") + "31");
            SB.AppendFormat(" AND TG006='{0}'",textBox2.Text);
            SB.AppendFormat(" GROUP BY TG006,SUBSTRING(TG003,1,6),TH004,TH005,MB004");
            SB.AppendFormat(" ORDER BY TG006, SUBSTRING(TG003,1,6),TH004");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");

            return SB;

        }

        public void SETFASTREPORT3()
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

            SQL1 = SETSQL3();
            Report report1 = new Report();
            report1.Load(@"REPORT\銷售月報表-多客戶.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL3()
        {
            string CID = null;
            string ST = textBox3.Text;
            string[] SARRARY = ST.Split(',');

            foreach (string S in SARRARY)
            {
                CID = CID + "'" + S + "',";
            }
            CID = CID + "''";

            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT SUBSTRING(TG003,1,6) AS 'YM',TH004,TH005,SUM(TH037) AS 'MONEY',SUM(LA011) AS 'NUM',MB004 ");
            SB.AppendFormat(" FROM [TK].dbo.COPTG,[TK].dbo.COPTH,[TK].dbo.INVLA,[TK].dbo.INVMB ");
            SB.AppendFormat(" WHERE TG001=TH001 AND TG002=TH002 AND LA006=TH001 AND LA007=TH002 AND LA008=TH003 AND TH004=MB001 AND TG023='Y'");
            SB.AppendFormat(" AND (TH004 LIKE '4%' OR TH004 LIKE '5%' ) AND TG003>='{0}' AND TG003<='{1}' ", dateTimePicker1.Value.ToString("yyyyMM") + "01", dateTimePicker2.Value.ToString("yyyyMM") + "31");
            SB.AppendFormat(" AND TG004 IN ({0} )", CID.ToString());
            SB.AppendFormat(" GROUP BY SUBSTRING(TG003,1,6),TH004,TH005,MB004 ORDER BY SUBSTRING(TG003,1,6),TH004  ");
            SB.AppendFormat("     ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");

            return SB;

        }


        public void SETFASTREPORT4(string SDAY,string EDAY,string TG005)
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

            SQL1 = SETSQL4(SDAY, EDAY, TG005);
            Report report1 = new Report();
            report1.Load(@"REPORT\銷售月報表-部門.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl3;
            report1.Show();
        }

        public StringBuilder SETSQL4(string SDAY, string EDAY,string TG005)
        {           

            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                             SELECT TG005,TG006,SUBSTRING(TG003,1,6) AS 'YM',SUM(TH037) AS 'MONEY',SUM(LA011) AS 'NUM'
                             FROM [TK].dbo.COPTG,[TK].dbo.COPTH,[TK].dbo.INVLA,[TK].dbo.INVMB
                             WHERE TG001=TH001 AND TG002=TH002
                             AND LA006=TH001 AND LA007=TH002 AND LA008=TH003
                             AND TH004=MB001
                             AND TG023='Y'
                             AND (TH004 LIKE '4%' OR TH004 LIKE '5%' )
                             AND TG003>='{0}' AND TG003<='{1}'
                             AND TG005='{2}'
                             GROUP BY TG005,TG006,SUBSTRING(TG003,1,6)
                             ORDER BY TG005,TG006, SUBSTRING(TG003,1,6)

                            ", SDAY, EDAY, TG005);

            return SB;

        }

        public void SETFASTREPORT5(string SDAY,  string TG006)
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

            SQL1 = SETSQL5(SDAY, TG006);
            Report report1 = new Report();
            report1.Load(@"REPORT\銷售年報-業務員-客戶.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl4;
            report1.Show();
        }

        public StringBuilder SETSQL5(string SDAY, string TG006)
        {

            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT SUBSTRING(TG003,1,6) AS YM,TG004,MA002,TG006,MV002,SUM(TH037) AS MM
                            FROM [TK].dbo.COPTG,[TK].dbo.COPTH, [TK].dbo.CMSMV ,[TK].dbo.COPMA
                            WHERE TG001=TH001 AND TG002=TH002
                            AND MV001=TG006
                            AND MA001=TG004
                            AND TG023='Y'
                            AND TG003 LIKE '{0}%'
                            AND TG006='{1}'
                            GROUP BY  SUBSTRING(TG003,1,6),TG004,MA002,TG006,MV002

                            ", SDAY, TG006);

            return SB;

        }

        public void SETFASTREPORT6(string SDAYS, string EDAYS, string LA001)
        {
            SDAYS = SDAYS + "01";
            EDAYS = EDAYS + "31";

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);



            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL6(SDAYS, EDAYS, LA001);
            Report report1 = new Report();
            report1.Load(@"REPORT\銷售月報表-商品.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl5;
            report1.Show();
        }

        public StringBuilder SETSQL6(string SDAYS,string EDAYS, string LA001)
        {
            string ST = LA001;
            string[] SARRARY = ST.Split(',');
            string LA001ID = null;

            if(SARRARY.Length>=2)
            {
                foreach (string S in SARRARY)
                {
                    LA001ID = LA001ID + "'" + S + "',";
                }
                LA001ID = LA001ID + "''";
            }
            else
            {
                LA001ID = "'" + LA001 + "'";

            }
           

            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT SUBSTRING(LA004,1,6) AS 'YM',LA001,MB002,SUM(LA011) LA011,SUM(LA013) LA013
                            FROM [TK].dbo.INVLA WITH(NOLOCK)
                            LEFT JOIN [TK].dbo.CMSMQ WITH(NOLOCK) ON LA006=MQ001
                            ,[TK].dbo.INVMB WITH(NOLOCK)
                            WHERE LA001=MB001
                            AND (MQ008='2' OR (ISNULL(MQ008,'')='' AND LA005=-1))
                            AND LA004>='{0}' AND LA004<='{1}'
                            AND LA001 IN ({2})

                            GROUP BY  LA001,MB002,LA005,SUBSTRING(LA004,1,6)
                            ORDER BY SUBSTRING(LA004,1,6),LA001

                            ", SDAYS, EDAYS, LA001ID);

            return SB;

        }


        #endregion

        #region BUTTON

        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT3();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT4(dateTimePicker5.Value.ToString("yyyyMM") + "01", dateTimePicker6.Value.ToString("yyyyMM") + "31",textBox4.Text.Trim());
        }
        private void button5_Click(object sender, EventArgs e)
        {
            SETFASTREPORT5(dateTimePicker7.Value.ToString("yyyy") , textBox5.Text.Trim());
        }
        private void button7_Click(object sender, EventArgs e)
        {
            SETFASTREPORT6(dateTimePicker8.Value.ToString("yyyyMM"), dateTimePicker9.Value.ToString("yyyyMM"), textBox7.Text.Trim());
        }

        #endregion


    }
}
