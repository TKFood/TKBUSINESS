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
    public partial class frmREPORTSASLARANKS : Form
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

        public frmREPORTSASLARANKS()
        {
            InitializeComponent();

            SETDATES();
            comboBox1load();
        }

        #region FUNCTION
        public void comboBox1load()
        {
            DataTable dt = new DataTable();
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);


            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"
                                SELECT [NAMES]
                                FROM [TKBUSINESS].[dbo].[TBPARA]
                                WHERE [KINDS]='SASLALA007'
                                ORDER BY VALUE
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
           
            sqlConn.Open();
           
            dt.Columns.Add("NAMES", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.DisplayMember = "NAMES";
            sqlConn.Close();


        }
        public void SETDATES()
        {
            DateTime currentDate = DateTime.Now;
            DateTime firstDayOfLastMonth = new DateTime(currentDate.Year, currentDate.Month - 1, 1);
            dateTimePicker1.Value = firstDayOfLastMonth;
            dateTimePicker2.Value = firstDayOfLastMonth;
        }
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            CHECKDATE_SDAYS(dateTimePicker1.Value);
        }
        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            CHECKDATE_EDAYS(dateTimePicker2.Value);
        }
        public void CHECKDATE_SDAYS(DateTime SDAYS)
        {
            DateTime currentDate = DateTime.Now;
            DateTime firstDayOfMonth = new DateTime(currentDate.Year, currentDate.Month, 1);
            DateTime firstDayOfLastMonth = new DateTime(currentDate.Year, currentDate.Month-1, 1);

            DateTime yourDateTime = SDAYS;// 設定您想要檢查的 DateTime 物件

            if (yourDateTime > firstDayOfMonth)
            {
                // 您的 DateTime 大於或等於本月第一天
                dateTimePicker1.Value = firstDayOfLastMonth;
                MessageBox.Show("日期只能在上個月之前");
            }
            else
            {
               
            }
        }
        public void CHECKDATE_EDAYS(DateTime EDAYS)
        {
            DateTime currentDate = DateTime.Now;
            DateTime firstDayOfMonth = new DateTime(currentDate.Year, currentDate.Month, 1);
            DateTime firstDayOfLastMonth = new DateTime(currentDate.Year, currentDate.Month - 1, 1);

            DateTime yourDateTime = EDAYS;// 設定您想要檢查的 DateTime 物件

            if (yourDateTime > firstDayOfMonth)
            {
                // 您的 DateTime 大於或等於本月第一天
                dateTimePicker2.Value = firstDayOfLastMonth;
                MessageBox.Show("日期只能在上個月之前");
            }
            else
            {

            }
        }

        public void SETFASTREPORT(string SDATES, string EDATES,string KINDS)
        {

            SDATES = SDATES + "01";
            EDATES = EDATES + "31";
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);


            StringBuilder SQL1 = new StringBuilder();     


            if(KINDS.Equals("門市部"))
            {
                SQL1 = SETSQL1(SDATES, EDATES, KINDS);
            }
            else if (KINDS.Equals("全公司"))
            {
                SQL1 = SETSQL3(SDATES, EDATES, KINDS);
            }
            else
            {
                SQL1 = SETSQL2(SDATES, EDATES, KINDS);
            }
            
    
            Report report1 = new Report();
            report1.Load(@"REPORT\產品貢獻度排名.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();
   

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1(string SDATES, string EDATES, string KINDS)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT 
                            LA005 AS '品號'
                            ,LA007 AS '部門'
                            ,ME002 AS '部門名'
                            ,MB002 AS '品名'
                            ,NUMS AS '銷售數量'
                            ,MONEYS AS '銷售金額'
                            ,COSTS AS '成本'
                            ,EARNSMONEYS AS '毛利'
                            ,EARNSMONEYSRATES AS '毛利率'
                            ,MONEYSPCTS AS '個別銷售'
                            ,EARNSMONEYSPCTS AS '毛利貢獻'
                            ,RANKS AS '貢獻比'
                            ,ROW_NUMBER() OVER (ORDER BY  RANKS DESC) AS '貢獻比排名'
                            FROM 
                            (
                            SELECT *
                            ,(MONEYS-COSTS) AS EARNSMONEYS
                            ,(CASE WHEN MONEYS>0 AND COSTS>0 THEN ((MONEYS-COSTS)/COSTS) ELSE 0 END ) AS EARNSMONEYSRATES
                            ,(MONEYS/SUM(MONEYS) OVER ()) AS MONEYSPCTS
                            ,((MONEYS-COSTS)/SUM((MONEYS-COSTS)) OVER ()) AS EARNSMONEYSPCTS
                            ,((MONEYS/SUM(MONEYS) OVER ())*((MONEYS-COSTS)/SUM((MONEYS-COSTS)) OVER ())) AS RANKS
                            FROM
                            (
                            SELECT LA005,'' LA007,'門市' ME002,MB002,SUM(LA016-LA019+LA025) AS NUMS,SUM(LA017-LA020-LA022-LA023) AS MONEYS,SUM(LA024) AS COSTS
                            FROM [TK].dbo.SASLA,[TK].dbo.INVMB,[TK].dbo.CMSME
                            WHERE 1=1
                            AND LA005=MB001
                            AND LA007=ME001
                            AND (LA005 LIKE '4%' OR  LA005 LIKE '5%')
                            AND LA005 NOT LIKE '599%'
                            AND ((MB002 NOT LIKE '%試吃%') OR (MB002  LIKE '%試吃%' AND (LA017-LA020-LA022-LA023)>0)) 
                            AND CONVERT(NVARCHAR,LA015,112)>='{0}'
                            AND CONVERT(NVARCHAR,LA015,112)<='{1}'
                            AND LA007 IN  (
                            SELECT [NAMES]
                            FROM [TKBUSINESS].[dbo].[TBPARA]
                            WHERE [KINDS]='{2}'
                            )
                            GROUP BY LA005,MB002
                            ) AS TEMP
                            ) AS TMEP2
                            ORDER BY RANKS DESC
                               
 
                            ", SDATES, EDATES, KINDS);

            return SB;

        }

        public StringBuilder SETSQL2(string SDATES, string EDATES, string KINDS)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"
                            SELECT 
                            LA005 AS '品號'
                            ,LA007 AS '部門'
                            ,ME002 AS '部門名'
                            ,MB002 AS '品名'
                            ,NUMS AS '銷售數量'
                            ,MONEYS AS '銷售金額'
                            ,COSTS AS '成本'
                            ,EARNSMONEYS AS '毛利'
                            ,EARNSMONEYSRATES AS '毛利率'
                            ,MONEYSPCTS AS '個別銷售'
                            ,EARNSMONEYSPCTS AS '毛利貢獻'
                            ,RANKS AS '貢獻比'
                            ,ROW_NUMBER() OVER (ORDER BY  RANKS DESC) AS '貢獻比排名'
                            FROM 
                            (
                            SELECT *
                            ,(MONEYS-COSTS) AS EARNSMONEYS
                            ,(CASE WHEN MONEYS>0 AND COSTS>0 THEN ((MONEYS-COSTS)/COSTS) ELSE 0 END ) AS EARNSMONEYSRATES
                            ,(MONEYS/SUM(MONEYS) OVER ()) AS MONEYSPCTS
                            ,((MONEYS-COSTS)/SUM((MONEYS-COSTS)) OVER ()) AS EARNSMONEYSPCTS
                            ,((MONEYS/SUM(MONEYS) OVER ())*((MONEYS-COSTS)/SUM((MONEYS-COSTS)) OVER ())) AS RANKS
                            FROM
                            (
                            SELECT LA005,LA007,ME002,MB002,SUM(LA016-LA019+LA025) AS NUMS,SUM(LA017-LA020-LA022-LA023) AS MONEYS,SUM(LA024) AS COSTS
                            FROM [TK].dbo.SASLA,[TK].dbo.INVMB,[TK].dbo.CMSME
                            WHERE 1=1
                            AND LA005=MB001
                            AND LA007=ME001
                            AND (LA005 LIKE '4%' OR  LA005 LIKE '5%')
                            AND LA005 NOT LIKE '599%'
                            AND ((MB002 NOT LIKE '%試吃%') OR (MB002  LIKE '%試吃%' AND (LA017-LA020-LA022-LA023)>0)) 
                            AND CONVERT(NVARCHAR,LA015,112)>='{0}'
                            AND CONVERT(NVARCHAR,LA015,112)<='{1}'
                            AND LA007 IN  (
                            SELECT [NAMES]
                            FROM [TKBUSINESS].[dbo].[TBPARA]
                            WHERE [KINDS]='{2}'
                            )
                            GROUP BY LA005,LA007,ME002,MB002
                            ) AS TEMP
                            ) AS TMEP2
                            ORDER BY RANKS DESC


                             ", SDATES, EDATES, KINDS);

            return SB;

        }

        public StringBuilder SETSQL3(string SDATES, string EDATES, string KINDS)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"
                            SELECT 
                            LA005 AS '品號'
                            ,LA007 AS '部門'
                            ,ME002 AS '部門名'
                            ,MB002 AS '品名'
                            ,NUMS AS '銷售數量'
                            ,MONEYS AS '銷售金額'
                            ,COSTS AS '成本'
                            ,EARNSMONEYS AS '毛利'
                            ,EARNSMONEYSRATES AS '毛利率'
                            ,MONEYSPCTS AS '個別銷售'
                            ,EARNSMONEYSPCTS AS '毛利貢獻'
                            ,RANKS AS '貢獻比'
                            ,ROW_NUMBER() OVER (ORDER BY  RANKS DESC) AS '貢獻比排名'
                            FROM 
                            (
                            SELECT *
                            ,(MONEYS-COSTS) AS EARNSMONEYS
                            ,(CASE WHEN MONEYS>0 AND COSTS>0 THEN ((MONEYS-COSTS)/COSTS) ELSE 0 END ) AS EARNSMONEYSRATES
                            ,(MONEYS/SUM(MONEYS) OVER ()) AS MONEYSPCTS
                            ,((MONEYS-COSTS)/SUM((MONEYS-COSTS)) OVER ()) AS EARNSMONEYSPCTS
                            ,((MONEYS/SUM(MONEYS) OVER ())*((MONEYS-COSTS)/SUM((MONEYS-COSTS)) OVER ())) AS RANKS
                            FROM
                            (
                            SELECT LA005,'' LA007,'全公司' ME002,MB002,SUM(LA016-LA019+LA025) AS NUMS,SUM(LA017-LA020-LA022-LA023) AS MONEYS,SUM(LA024) AS COSTS
                            FROM [TK].dbo.SASLA,[TK].dbo.INVMB,[TK].dbo.CMSME
                            WHERE 1=1
                            AND LA005=MB001
                            AND LA007=ME001
                            AND (LA005 LIKE '4%' OR  LA005 LIKE '5%')
                            AND LA005 NOT LIKE '599%'
                            AND ((MB002 NOT LIKE '%試吃%') OR (MB002  LIKE '%試吃%' AND (LA017-LA020-LA022-LA023)>0)) 
                            AND CONVERT(NVARCHAR,LA015,112)>='{0}'
                            AND CONVERT(NVARCHAR,LA015,112)<='{1}'
                         
                            GROUP BY LA005,LA007,ME002,MB002
                            ) AS TEMP
                            ) AS TMEP2
                            ORDER BY RANKS DESC


                             ", SDATES, EDATES, KINDS);

            return SB;

        }
        #endregion
        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMM"), dateTimePicker2.Value.ToString("yyyyMM"),comboBox1.Text.ToString());
        }

        #endregion

       
    }
}
