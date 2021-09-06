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
using System.Text.RegularExpressions;
using FastReport;
using FastReport.Data;
using TKITDLL;

namespace TKBUSINESS
{
    public partial class frmREPORTCOP : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;

        int result;

        public Report report1 { get; private set; }

        public Report report2 { get; private set; }


        public frmREPORTCOP()
        {
            InitializeComponent();

            combobox1load();
        }

        #region FUNCTION
        public void combobox1load()
        {

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            String Sequel = "SELECT MV001,MV002  FROM [TK].dbo.CMSMV WHERE MV001 IN ('140078','140049','160155','090002','160048')";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MV001", typeof(string));
            dt.Columns.Add("MV002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MV001";
            comboBox1.DisplayMember = "MV002";
            sqlConn.Close();



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

            

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\業務商品排名表.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL();
            Table.SelectCommand = SQL;
            report1.Preview = previewControl1;
            report1.Show();

        }
        public string SETFASETSQL()
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            FASTSQL.AppendFormat(@"  SELECT ME002 AS '部門',MV002 AS '業務員',TH004 AS '品號',TH005 AS '品名',SUM(TH037) AS '金額',SUM(NUM) AS '數量',TH009 AS '單位',SUM(TH037)/SUM(NUM) AS '平均售價'");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.COPTG,[TK].dbo.[VCOPTHINVMD],[TK].dbo.CMSMV,[TK].dbo.CMSME ");
            FASTSQL.AppendFormat(@"  WHERE TG001=TH001 AND TG002=TH002");
            FASTSQL.AppendFormat(@"  AND TG006=MV001");
            FASTSQL.AppendFormat(@"  AND TG005=ME001");
            FASTSQL.AppendFormat(@"  AND TG003>='{0}' AND TG003<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            FASTSQL.AppendFormat(@"  AND TG006='{0}'",comboBox1.SelectedValue,ToString());
            FASTSQL.AppendFormat(@"  GROUP BY ME002,MV002,TH004,TH005,TH009");
            FASTSQL.AppendFormat(@"  HAVING SUM(TH037)>0 ");
            FASTSQL.AppendFormat(@"  ORDER BY ME002,SUM(TH037) DESC");
            FASTSQL.AppendFormat(@"  ");
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }

        public void SETFASTREPORT2()
        {

            string SQL;
            report2 = new Report();
            report2.Load(@"REPORT\業務客戶排名表.frx");

            report2.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            TableDataSource Table = report2.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL2();
            Table.SelectCommand = SQL;
            report2.Preview = previewControl2;
            report2.Show();

        }
        public string SETFASETSQL2()
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            FASTSQL.AppendFormat(@"  SELECT  ME002 AS '部門', MV002 AS '業務員',TG007 AS '客戶',SUM(TH037) AS '金額',SUM(NUM) AS '數量'");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.COPTG,[TK].dbo.[VCOPTHINVMD],[TK].dbo.CMSMV,[TK].dbo.CMSME");
            FASTSQL.AppendFormat(@"  WHERE TG001=TH001 AND TG002=TH002");
            FASTSQL.AppendFormat(@"  AND TG006=MV001");
            FASTSQL.AppendFormat(@"  AND TG005=ME001");
            FASTSQL.AppendFormat(@"  AND TG003>='{0}' AND TG003<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            FASTSQL.AppendFormat(@"  AND TG006='{0}'", comboBox1.SelectedValue, ToString());
            FASTSQL.AppendFormat(@"  GROUP BY ME002,MV002,TG007");
            FASTSQL.AppendFormat(@"  HAVING SUM(TH037)>0");
            FASTSQL.AppendFormat(@"  ORDER BY ME002,SUM(TH037) DESC ");
            FASTSQL.AppendFormat(@"  ");
            FASTSQL.AppendFormat(@"  ");
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }

        #endregion

        #region       


        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
            SETFASTREPORT2();
        }

        #endregion
    }
}
