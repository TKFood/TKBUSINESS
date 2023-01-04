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
    public partial class frmREPORTSALESWEEK : Form
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

        public frmREPORTSALESWEEK()
        {
            InitializeComponent();

            SETDATE();
        }

        #region FUNCTION

        #endregion
        public void SETDATE()
        {
            DateTime FirstDay = DateTime.Now.AddDays(-DateTime.Now.Day + 1);
            DateTime LastDay = DateTime.Now.AddMonths(1).AddDays(-DateTime.Now.AddMonths(1).Day);

            dateTimePicker3.Value = FirstDay;
            dateTimePicker4.Value = LastDay;
        }

        public void SETFASTREPORT(string SDATES, string EDATES)
        {
            string YEARSMOTNS = SDATES.Substring(0,6);
            DataTable DT = SEARCH_TK_ZTARGETMONEYS(YEARSMOTNS);
            string P1 = DT.Rows[0]["YEARSMOTNS"].ToString();
            string P2 = DT.Rows[0]["INTARGETMONEYS"].ToString();
            P2 = String.Format("{0:#,##0;(#,##0);0}", Convert.ToInt32(P2));
            string P3 = DT.Rows[0]["OUTTARGETMONEYS"].ToString();
            P3 = String.Format("{0:#,##0;(#,##0);0}", Convert.ToInt32(P3));

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


            SQL1 = SETSQL(SDATES, EDATES);
            SQL2 = SETSQL2(SDATES, EDATES); 
            Report report1 = new Report();
            report1.Load(@"REPORT\業務-每月訂單-週報用.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();
            TableDataSource table1 = report1.GetDataSource("Table1") as TableDataSource;
            table1.SelectCommand = SQL2.ToString();

            report1.SetParameterValue("P1", P1);
            report1.SetParameterValue("P2", P2);
            report1.SetParameterValue("P3", P3);

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL(string SDATES,string EDATES)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 

                                SELECT *
                                FROM 
                                (
                                SELECT TC001 AS '訂單單別',TC002 AS '訂單單號',MA002 AS '客戶簡稱'
                                ,CASE WHEN TC016='1' THEN '應稅內含' WHEN TC016='2' THEN '應稅外加' END  AS '課稅別'
                                ,ME002 AS '部門',TD005 AS '品名',TD008 AS 	'訂單數量',TD009 AS '已交數量',TD024 AS	'贈品數量',TD025 AS	'贈品已交量',(TD008-TD009) AS  '未出數量',TD010 AS 	'單位',TD011 AS  '單價',(TD008-TD009)*TD011 AS '未出貨金額',TD013 AS'預交日'
                                FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                WHERE TC001=TD001 AND TC002=TD002
                                AND TC004=MA001
                                AND TC005=ME001
                                AND TC027 IN ('Y','N')
                                AND TD016 IN ('N')
                                AND TD013>='{0}' AND TD013<='{1}'
                                AND TC005 IN ('117700','117100','117200','117400')
                                AND TC001 NOT IN ('A223')
                                UNION ALL
                                SELECT TC001 AS '訂單單別',TC002 AS '訂單單號',MA002 AS '客戶簡稱'
                                ,CASE WHEN TC016='1' THEN '應稅內含' WHEN TC016='2' THEN '應稅外加' END  AS '課稅別'
                                ,ME002 AS '部門',TD005 AS '品名',TD008 AS 	'訂單數量',TD009 AS '已交數量',TD024 AS	'贈品數量',TD025 AS	'贈品已交量',(TD008-TD009) AS  '未出數量',TD010 AS 	'單位',TD011 AS  '單價',(TD008-TD009)*TD011 AS '未出貨金額',TD013 AS'預交日'
                                FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                WHERE TC001=TD001 AND TC002=TD002
                                AND TC004=MA001
                                AND TC005=ME001
                                AND TC027 IN ('Y','N')
                                AND TD016 IN ('N')
                                 AND TD013>='{0}' AND TD013<='{1}'
                                AND TC005 IN ('117700','117100','117200','117400')
                                AND TC001  IN ('A223')
                                AND TC004 NOT IN ('2248500100')
                                AND TC004 NOT IN ('2248500100')
                                ) AS TEMP 
                                ORDER BY 訂單單別,未出貨金額 DESC


                            ", SDATES,EDATES);

            return SB;

        }

        public StringBuilder SETSQL2(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 

                                SELECT *
                                FROM 
                                (
                                SELECT TC001 AS '訂單單別',TC002 AS '訂單單號',MA002 AS '客戶簡稱'
                                ,CASE WHEN TC016='1' THEN '應稅內含' WHEN TC016='2' THEN '應稅外加' END  AS '課稅別'
                                ,ME002 AS '部門',TD005 AS '品名',TD008 AS 	'訂單數量',TD009 AS '已交數量',TD024 AS	'贈品數量',TD025 AS	'贈品已交量',(TD008-TD009) AS  '未出數量',TD010 AS 	'單位',TD011 AS  '單價',(TD008-TD009)*TD011 AS '未出貨金額',TD013 AS'預交日'
                                FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                WHERE TC001=TD001 AND TC002=TD002
                                AND TC004=MA001
                                AND TC005=ME001
                                AND TC027 IN ('Y','N')
                                AND TD016 IN ('N')
                                AND TD013>='{0}' AND TD013<='{1}'
                                AND TC005 IN ('117800','117500','117600')
                                AND TC001 NOT IN ('A223')
                                UNION ALL
                                SELECT TC001 AS '訂單單別',TC002 AS '訂單單號',MA002 AS '客戶簡稱'
                                ,CASE WHEN TC016='1' THEN '應稅內含' WHEN TC016='2' THEN '應稅外加' END  AS '課稅別'
                                ,ME002 AS '部門',TD005 AS '品名',TD008 AS 	'訂單數量',TD009 AS '已交數量',TD024 AS	'贈品數量',TD025 AS	'贈品已交量',(TD008-TD009) AS  '未出數量',TD010 AS 	'單位',TD011 AS  '單價',(TD008-TD009)*TD011 AS '未出貨金額',TD013 AS'預交日'
                                FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                WHERE TC001=TD001 AND TC002=TD002
                                AND TC004=MA001
                                AND TC005=ME001
                                AND TD016 IN ('N')
                                AND TD013>='{0}' AND TD013<='{1}'
                                AND TC005 IN ('117800','117500','117600')
                                AND TC001  IN ('A223')

                                ) AS TEMP 
                                ORDER BY 訂單單別,未出貨金額 DESC

                             ", SDATES, EDATES);

            return SB;

        }

        public DataTable SEARCH_TK_ZTARGETMONEYS(string YEARSMOTNS)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();
            

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds1.Clear();

                sbSql.AppendFormat(@" 
                                    SELECT [YEARSMOTNS]
                                    ,[INTARGETMONEYS]
                                    ,[OUTTARGETMONEYS]
                                    FROM [TK].[dbo].[ZTARGETMONEYS]
                                    WHERE [YEARSMOTNS]='{0}'
                                    ", YEARSMOTNS);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    
                    return ds1.Tables["ds1"];

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
        }

        #endregion
    }
}
