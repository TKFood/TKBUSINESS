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
    public partial class frmREPORTSASLA : Form
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


        public frmREPORTSASLA()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT(string MA001, string SDATES, string EDATES)
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

            SQL1 = SETSQL(MA001, SDATES,  EDATES);
            Report report1 = new Report();
            report1.Load(@"REPORT\銷售統計月報.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL(string MA001,string SDATES,string EDATES)
        {
            StringBuilder SB = new StringBuilder();

           
            if(!string.IsNullOrEmpty(MA001))
            {
                SB.AppendFormat(@" 
                            SELECT YM,LA006,MA002,LA005,MB002,NUMS,MMS
                            FROM (
                            SELECT SUBSTRING(CONVERT(NVARCHAR,LA015,112),1,6) AS 'YM',LA006,LA005,SUM(LA016-LA019) AS 'NUMS',SUM(LA017-LA020-LA022-LA023) AS 'MMS'
                            FROM [TK].dbo.SASLA
                            WHERE CONVERT(NVARCHAR,LA015,112)>='{0}' AND CONVERT(NVARCHAR,LA015,112)<='{1}' 
                            AND LA005 NOT LIKE '1%'
                            AND LA005 NOT LIKE '2%'
                            AND LA005 NOT LIKE '3%'
                            AND LA006 IN (
                            SELECT LA006
                            FROM [TK].dbo.SASLA
                            WHERE CONVERT(NVARCHAR,LA015,112)>='{0}' AND CONVERT(NVARCHAR,LA015,112)<='{1}' 
                            AND LA006 LIKE '2%'
                            GROUP BY LA006

                            )
                            GROUP BY SUBSTRING(CONVERT(NVARCHAR,LA015,112),1,6),LA006,LA005
                            ) AS TEMP
                            LEFT JOIN [TK].dbo.COPMA ON MA001=LA006
                            LEFT JOIN [TK].dbo.INVMB ON MB001=LA005
                            WHERE (LA006 LIKE '%{2}%' OR MA002 LIKE '%{2}%' )

                            ", SDATES, EDATES, MA001);

            }
            else
            {
                SB.AppendFormat(@" 
                            SELECT YM,LA006,MA002,LA005,MB002,NUMS,MMS
                            FROM (
                            SELECT SUBSTRING(CONVERT(NVARCHAR,LA015,112),1,6) AS 'YM',LA006,LA005,SUM(LA016-LA019) AS 'NUMS',SUM(LA017-LA020-LA022-LA023) AS 'MMS'
                            FROM [TK].dbo.SASLA
                            WHERE CONVERT(NVARCHAR,LA015,112)>='{0}' AND CONVERT(NVARCHAR,LA015,112)<='{1}' 
                            AND LA005 NOT LIKE '1%'
                            AND LA005 NOT LIKE '2%'
                            AND LA005 NOT LIKE '3%'
                            AND LA006 IN (
                            SELECT LA006
                            FROM [TK].dbo.SASLA
                            WHERE CONVERT(NVARCHAR,LA015,112)>='{0}' AND CONVERT(NVARCHAR,LA015,112)<='{1}' 
                           
                            GROUP BY LA006

                            )
                            GROUP BY SUBSTRING(CONVERT(NVARCHAR,LA015,112),1,6),LA006,LA005
                            ) AS TEMP
                            LEFT JOIN [TK].dbo.COPMA ON MA001=LA006
                            LEFT JOIN [TK].dbo.INVMB ON MB001=LA005

                            ",  SDATES, EDATES);

            }


            return SB;

        }

        #endregion

        #region BUTTON

        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(textBox1.Text, dateTimePicker1.Value.ToString("yyyyMM") + "01", dateTimePicker2.Value.ToString("yyyyMM") + "31");
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        #endregion

    }
}
