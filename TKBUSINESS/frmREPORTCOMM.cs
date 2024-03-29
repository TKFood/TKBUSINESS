﻿using System;
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
    public partial class frmREPORTCOMM : Form
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
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;
        int result;

        public frmREPORTCOMM()
        {
            InitializeComponent();

            textBox1.Text = DateTime.Now.Year.ToString();
            textBox2.Text = DateTime.Now.Year.ToString();
            textBox3.Text = (DateTime.Now.Month-1).ToString().PadLeft(2, '0');
            textBox4.Text = DateTime.Now.Year.ToString();
            textBox5.Text = (DateTime.Now.Month - 1).ToString().PadLeft(2, '0');
            textBox6.Text = DateTime.Now.Year.ToString();
            textBox7.Text = (DateTime.Now.Month - 1).ToString().PadLeft(2, '0');
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
            report1.Load(@"REPORT\預算跟銷售的業務員、商品比較表.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", textBox1.Text);
            //report1.SetParameterValue("P2", textBox2.Text);
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            StringBuilder SB = new StringBuilder();
            //MessageBox.Show(strLineData);

            string THISYEARS = textBox1.Text.Trim();
            string LASTYEARS = (Convert.ToInt32(textBox1.Text.Trim()) - 1).ToString();

            SB.AppendFormat(@" 
                            DECLARE @THISYEARS nvarchar(10)
                            DECLARE @LASTYEARS nvarchar(10)
                            DECLARE @MONTHS nvarchar(10)
                            SET @THISYEARS='{0}'
                            SET @LASTYEARS='{1}'
                            SET @MONTHS=''

                            SELECT ID3 AS '業務員代',MV002 AS '業務員',@THISYEARS AS '年度'
                            ,SUM(PRE202101+PRE202102+PRE202103+PRE202104+PRE202105+PRE202106+PRE202107+PRE202108+PRE202109+PRE202110+PRE202111+PRE202112)  AS '年度預算'
                            ,SUM((IN202101-OUT202101)+(IN202102-OUT202102)+(IN202103-OUT202103)+(IN202104-OUT202104)+(IN202105-OUT202105)+(IN202106-OUT202106)+(IN202107-OUT202107)+(IN202108-OUT202108)+(IN202109-OUT202109)+(IN202110-OUT202110)+(IN202111-OUT202111)+(IN202112-OUT202112)) AS '本年銷售金額'
                            ,((SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG023='Y' AND TG001 NOT IN ('A233','A234') AND TG006=ID3 AND TG003 LIKE @LASTYEARS+'%')-(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI019='Y' AND TI001 NOT IN ('A243','A246') AND TI006=ID3 AND TI003 LIKE @LASTYEARS+'%')) AS '去年銷售金額'
                            ,(SUM((IN202101-OUT202101)+(IN202102-OUT202102)+(IN202103-OUT202103)+(IN202104-OUT202104)+(IN202105-OUT202105)+(IN202106-OUT202106)+(IN202107-OUT202107)+(IN202108-OUT202108)+(IN202109-OUT202109)+(IN202110-OUT202110)+(IN202111-OUT202111)+(IN202112-OUT202112)))/(SUM(PRE202101+PRE202102+PRE202103+PRE202104+PRE202105+PRE202106+PRE202107+PRE202108+PRE202109+PRE202110+PRE202111+PRE202112)) AS '預算累積達成率'
                            FROM(
                            SELECT ID1,MA002,ID3,MV002
                           ,(SELECT ISNULL(SUM(MN005),0) FROM [TK].dbo.COPMM,[TK].dbo.COPMN WHERE MM001=MN001 AND MM002=MN002 AND MM003=ID1 AND MM001=@THISYEARS AND MN003='01') AS 'PRE202101'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG004=ID1 AND TG023='Y' AND TG001 NOT IN ('A233','A234') AND TG003 LIKE @THISYEARS+'01%') 'IN202101'
                            ,(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI004=ID1 AND TI019='Y' AND TI001 NOT IN ('A243','A246') AND TI003 LIKE @THISYEARS+'01%') 'OUT202101'
                            ,(SELECT ISNULL(SUM(MN005),0) FROM [TK].dbo.COPMM,[TK].dbo.COPMN WHERE MM001=MN001 AND MM002=MN002 AND MM003=ID1 AND MM001=@THISYEARS AND MN003='02') AS 'PRE202102'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG004=ID1 AND TG023='Y' AND TG001 NOT IN ('A233','A234') AND TG003 LIKE @THISYEARS+'02%') 'IN202102'
                            ,(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI004=ID1 AND TI019='Y' AND TI001 NOT IN ('A243','A246') AND TI003 LIKE @THISYEARS+'02%') 'OUT202102'
                            ,(SELECT ISNULL(SUM(MN005),0) FROM [TK].dbo.COPMM,[TK].dbo.COPMN WHERE MM001=MN001 AND MM002=MN002 AND MM003=ID1 AND MM001=@THISYEARS AND MN003='03') AS 'PRE202103'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG004=ID1 AND TG023='Y' AND TG001 NOT IN ('A233','A234') AND TG003 LIKE @THISYEARS+'03%') 'IN202103'
                            ,(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI004=ID1 AND TI019='Y' AND TI001 NOT IN ('A243','A246') AND TI003 LIKE @THISYEARS+'03%') 'OUT202103'
                            ,(SELECT ISNULL(SUM(MN005),0) FROM [TK].dbo.COPMM,[TK].dbo.COPMN WHERE MM001=MN001 AND MM002=MN002 AND MM003=ID1 AND MM001=@THISYEARS AND MN003='04') AS 'PRE202104'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG004=ID1 AND TG023='Y' AND TG001 NOT IN ('A233','A234') AND TG003 LIKE @THISYEARS+'04%') 'IN202104'
                            ,(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI004=ID1 AND TI019='Y' AND TI001 NOT IN ('A243','A246') AND TI003 LIKE @THISYEARS+'04%') 'OUT202104'
                            ,(SELECT ISNULL(SUM(MN005),0) FROM [TK].dbo.COPMM,[TK].dbo.COPMN WHERE MM001=MN001 AND MM002=MN002 AND MM003=ID1 AND MM001=@THISYEARS AND MN003='05') AS 'PRE202105'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG004=ID1 AND TG023='Y' AND TG001 NOT IN ('A233','A234') AND TG003 LIKE @THISYEARS+'05%') 'IN202105'
                            ,(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI004=ID1 AND TI019='Y' AND TI001 NOT IN ('A243','A246') AND TI003 LIKE @THISYEARS+'05%') 'OUT202105'
                            ,(SELECT ISNULL(SUM(MN005),0) FROM [TK].dbo.COPMM,[TK].dbo.COPMN WHERE MM001=MN001 AND MM002=MN002 AND MM003=ID1 AND MM001=@THISYEARS AND MN003='06') AS 'PRE202106'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG004=ID1 AND TG023='Y' AND TG001 NOT IN ('A233','A234') AND TG003 LIKE @THISYEARS+'06%') 'IN202106'
                            ,(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI004=ID1 AND TI019='Y' AND TI001 NOT IN ('A243','A246') AND TI003 LIKE @THISYEARS+'06%') 'OUT202106'
                            ,(SELECT ISNULL(SUM(MN005),0) FROM [TK].dbo.COPMM,[TK].dbo.COPMN WHERE MM001=MN001 AND MM002=MN002 AND MM003=ID1 AND MM001=@THISYEARS AND MN003='07') AS 'PRE202107'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG004=ID1 AND TG023='Y' AND TG001 NOT IN ('A233','A234') AND TG003 LIKE @THISYEARS+'07%') 'IN202107'
                            ,(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI004=ID1 AND TI019='Y' AND TI001 NOT IN ('A243','A246') AND TI003 LIKE @THISYEARS+'07%') 'OUT202107'
                            ,(SELECT ISNULL(SUM(MN005),0) FROM [TK].dbo.COPMM,[TK].dbo.COPMN WHERE MM001=MN001 AND MM002=MN002 AND MM003=ID1 AND MM001=@THISYEARS AND MN003='08') AS 'PRE202108'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG004=ID1 AND TG023='Y' AND TG001 NOT IN ('A233','A234') AND TG003 LIKE @THISYEARS+'08%') 'IN202108'
                            ,(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI004=ID1 AND TI019='Y' AND TI001 NOT IN ('A243','A246') AND TI003 LIKE @THISYEARS+'08%') 'OUT202108'
                            ,(SELECT ISNULL(SUM(MN005),0) FROM [TK].dbo.COPMM,[TK].dbo.COPMN WHERE MM001=MN001 AND MM002=MN002 AND MM003=ID1 AND MM001=@THISYEARS AND MN003='09') AS 'PRE202109'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG004=ID1 AND TG023='Y' AND TG001 NOT IN ('A233','A234') AND TG003 LIKE @THISYEARS+'09%') 'IN202109'
                            ,(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI004=ID1 AND TI019='Y' AND TI001 NOT IN ('A243','A246') AND TI003 LIKE @THISYEARS+'09%') 'OUT202109'
                            ,(SELECT ISNULL(SUM(MN005),0) FROM [TK].dbo.COPMM,[TK].dbo.COPMN WHERE MM001=MN001 AND MM002=MN002 AND MM003=ID1 AND MM001=@THISYEARS AND MN003='10') AS 'PRE202110'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG004=ID1 AND TG023='Y' AND TG001 NOT IN ('A233','A234') AND TG003 LIKE @THISYEARS+'10%') 'IN202110'
                            ,(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI004=ID1 AND TI019='Y' AND TI001 NOT IN ('A243','A246') AND TI003 LIKE @THISYEARS+'10%') 'OUT202110'
                            ,(SELECT ISNULL(SUM(MN005),0) FROM [TK].dbo.COPMM,[TK].dbo.COPMN WHERE MM001=MN001 AND MM002=MN002 AND MM003=ID1 AND MM001=@THISYEARS AND MN003='11') AS 'PRE202111'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG004=ID1 AND TG023='Y' AND TG001 NOT IN ('A233','A234') AND TG003 LIKE @THISYEARS+'11%') 'IN202111'
                            ,(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI004=ID1 AND TI019='Y' AND TI001 NOT IN ('A243','A246') AND TI003 LIKE @THISYEARS+'11%') 'OUT202111'
                            ,(SELECT ISNULL(SUM(MN005),0) FROM [TK].dbo.COPMM,[TK].dbo.COPMN WHERE MM001=MN001 AND MM002=MN002 AND MM003=ID1 AND MM001=@THISYEARS AND MN003='12') AS 'PRE202112'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG004=ID1 AND TG023='Y' AND TG001 NOT IN ('A233','A234') AND TG003 LIKE @THISYEARS+'12%') 'IN202112'
                            ,(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI004=ID1 AND TI019='Y' AND TI001 NOT IN ('A243','A246') AND TI003 LIKE @THISYEARS+'12%') 'OUT202112'

                            FROM [TK].dbo.ZSLAES
                            LEFT JOIN [TK].dbo.CMSMV ON MV001=ID3
                            LEFT JOIN [TK].dbo.COPMA ON MA001=ID1
                            WHERE YEARS=@THISYEARS
                            ) AS TEMP
                            GROUP BY ID3,MV002
                            UNION ALL
                            SELECT MM003,MA002,MN001,SUM(MN005) AS 'PRE'
                            ,((SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG023='Y' AND TG001 IN ('A233','A234') AND TG003 LIKE MN001+'%')-(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI019='Y' AND TI001 IN ('A243','A246') AND TI003 LIKE MN001+'%')) '2021'
                            ,((SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG023='Y' AND TG001 IN ('A233','A234') AND TG003 LIKE '2020%')-(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI019='Y' AND TI001 IN ('A243','A246') AND TI003 LIKE '{1}%'))AS '電商去年實收'
                            ,((SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG023='Y' AND TG001 IN ('A233','A234') AND TG003 LIKE MN001+'%')-(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI019='Y' AND TI001 IN ('A243','A246') AND TI003 LIKE MN001+'%'))/(SUM(MN005))  AS '預算累積達成率'
                            FROM [TK].dbo.COPMM,[TK].dbo.COPMN,[TK].dbo.COPMA 
                            WHERE MM001=MN001 AND MM002=MN002 
                            AND MA001=MM003
                            AND MM001=@THISYEARS
                            AND MM003='44900001'
                            GROUP BY MM003,MA002,MN001

                            ", THISYEARS, LASTYEARS);

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
            Report report1 = new Report();
            report1.Load(@"REPORT\各業務員各月總預算及業績.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", textBox1.Text);
            //report1.SetParameterValue("P2", textBox2.Text);
            report1.Preview = previewControl2;
            report1.Show();
        }

        public StringBuilder SETSQL2()
        {
            StringBuilder SB = new StringBuilder();
            //MessageBox.Show(strLineData);

            string THISYEARS = textBox2.Text.Trim();
            string LASTYEARS = (Convert.ToInt32(textBox2.Text.Trim()) - 1).ToString();
            string THISMONTHS= textBox3.Text.Trim();

            SB.AppendFormat(@" 
                            DECLARE @THISYEARS nvarchar(10)
                            DECLARE @LASTYEARS nvarchar(10)
                            DECLARE @MONTHS nvarchar(10)
                            SET @THISYEARS='{0}'
                            SET @LASTYEARS='{1}'
                            SET @MONTHS='{2}'

                            SELECT @THISYEARS AS '年度',@MONTHS  AS '月份',ID3  AS '業務代',MV002  AS '業務員'
                            ,SUM(PRE2021MOTH) AS '本月預算'
                            ,SUM((IN2021MOTH-OUT2021MOTH)) AS '本月實收'
                            ,SUM((IN2020MOTH-OUT2020MOTH)) AS '去年同期實收'
                            ,((SUM((IN2021MOTH-OUT2021MOTH)))-SUM(PRE2021MOTH)) AS '實收跟預算的差異'
                            FROM(
                            SELECT ID1,MA002,ID3,MV002
                            ,(SELECT ISNULL(SUM(MN005),0) FROM [TK].dbo.COPMM,[TK].dbo.COPMN WHERE MM001=MN001 AND MM002=MN002 AND MM003=ID1 AND MM001=@THISYEARS AND MN003=@MONTHS) AS 'PRE2021MOTH'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG004=ID1 AND TG023='Y' AND TG001 NOT IN ('A233','A234') AND TG003 LIKE @THISYEARS+@MONTHS+'%') 'IN2021MOTH'
                            ,(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI004=ID1 AND TI019='Y' AND TI001 NOT IN ('A243','A246') AND TI003 LIKE @THISYEARS+@MONTHS+'%') 'OUT2021MOTH'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG004=ID1 AND TG023='Y' AND TG001 NOT IN ('A233','A234') AND TG003 LIKE @LASTYEARS+@MONTHS+'%') 'IN2020MOTH'
                            ,(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI004=ID1 AND TI019='Y' AND TI001 NOT IN ('A243','A246') AND TI003 LIKE @LASTYEARS+@MONTHS+'%') 'OUT2020MOTH'

                            FROM [TK].dbo.ZSLAES
                            LEFT JOIN [TK].dbo.CMSMV ON MV001=ID3
                            LEFT JOIN [TK].dbo.COPMA ON MA001=ID1
                            ) AS TEMP
                            GROUP BY ID3,MV002
                            UNION ALL
                            SELECT   @THISYEARS AS '年度',@MONTHS  AS '月份',MM003,MA002,PRE,(IN2021-OUT2021),(IN2020-OUT2020),((IN2021-OUT2021)-PRE)
                            FROM (
                            SELECT MM003,MA002,MN001,MN003,SUM(MN005) AS 'PRE'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG023='Y' AND TG001 IN ('A233','A234') AND TG003 LIKE MN001+MN003+'%') 'IN2021'
                            ,(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI019='Y' AND TI001 IN ('A243','A246') AND TI003 LIKE MN001+MN003+'%') 'OUT2021'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG023='Y' AND TG001 IN ('A233','A234') AND TG003 LIKE @LASTYEARS+MN003+'%') 'IN2020'
                            ,(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI019='Y' AND TI001 IN ('A243','A246') AND TI003 LIKE @LASTYEARS+MN003+'%') 'OUT2020'
                            FROM [TK].dbo.COPMM,[TK].dbo.COPMN,[TK].dbo.COPMA 
                            WHERE MM001=MN001 AND MM002=MN002 
                            AND MA001=MM003
                            AND MM001=@THISYEARS
                            AND MM003='44900001'
                            AND MN003=@MONTHS
                            GROUP BY MM003,MA002,MN001,MN003
                            )AS TEMP


                            ", THISYEARS, LASTYEARS,THISMONTHS);

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
            report1.Load(@"REPORT\各客戶月份業績達成.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", textBox1.Text);
            //report1.SetParameterValue("P2", textBox2.Text);
            report1.Preview = previewControl3;
            report1.Show();
        }

        public StringBuilder SETSQL3()
        {
            StringBuilder SB = new StringBuilder();
            //MessageBox.Show(strLineData);

            string THISYEARS = textBox2.Text.Trim();
            string LASTYEARS = (Convert.ToInt32(textBox2.Text.Trim()) - 1).ToString();
            string THISMONTHS = textBox3.Text.Trim();

            SB.AppendFormat(@" 
                            DECLARE @THISYEARS nvarchar(10)
                            DECLARE @LASTYEARS nvarchar(10)
                            DECLARE @MONTHS nvarchar(10)
                            SET @THISYEARS='{0}'
                            SET @LASTYEARS='{1}'
                            SET @MONTHS='{2}'


                            SELECT @THISYEARS AS '年度',@MONTHS  AS '月份',ID1  AS '客戶代',MA002  AS '客戶',ID3  AS '業務代',MV002  AS '業務員'
                            ,PRE2021 AS '本月預算',(IN2021MONTH-OUT2021MONTH) AS '本月實收',(IN2020MONTH-OUT2020MONTH) AS '去年同期實收'
                            ,((IN2021MONTH-OUT2021MONTH)-PRE2021) AS '實收跟預算的差異'
                            FROM
                            (
                            SELECT ID1,MA002,ID3,MV002
                            ,(SELECT ISNULL(SUM(MN005),0) FROM [TK].dbo.COPMM,[TK].dbo.COPMN WHERE MM001=MN001 AND MM002=MN002 AND MM003=ID1 AND MM001=@THISYEARS AND MN003=@MONTHS) AS 'PRE2021'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG004=ID1 AND TG023='Y' AND TG001 NOT IN ('A233','A234') AND TG003 LIKE @THISYEARS+@MONTHS+'%') 'IN2021MONTH'
                            ,(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI004=ID1 AND TI019='Y' AND TI001 NOT IN ('A243','A246') AND TI003 LIKE @THISYEARS+@MONTHS+'%') 'OUT2021MONTH'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG004=ID1 AND TG023='Y' AND TG001 NOT IN ('A233','A234') AND TG003 LIKE @LASTYEARS+@MONTHS+'%') 'IN2020MONTH'
                            ,(SELECT ISNULL(SUM(TJ033),0) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI004=ID1 AND TI019='Y' AND TI001 NOT IN ('A243','A246') AND TI003 LIKE @LASTYEARS+@MONTHS+'%') 'OUT2020MONTH'

                            FROM [TK].dbo.ZSLAES
                            LEFT JOIN [TK].dbo.CMSMV ON MV001=ID3
                            LEFT JOIN [TK].dbo.COPMA ON MA001=ID1
                            ) AS TEMP
                            ORDER BY MV002,((IN2021MONTH-OUT2021MONTH)-PRE2021),ID1



                            ", THISYEARS, LASTYEARS, THISMONTHS);

            return SB;

        }

        public void SETFASTREPORT4()
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

            SQL1 = SETSQL4();
            Report report1 = new Report();
            report1.Load(@"REPORT\客戶月份的預算及銷售商品差異表.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", textBox1.Text);
            //report1.SetParameterValue("P2", textBox2.Text);
            report1.Preview = previewControl4;
            report1.Show();
        }

        public StringBuilder SETSQL4()
        {
            StringBuilder SB = new StringBuilder();
            //MessageBox.Show(strLineData);

            string THISYEARS = textBox6.Text.Trim();
            string LASTYEARS = (Convert.ToInt32(textBox6.Text.Trim()) - 1).ToString();
            string THISMONTHS = textBox7.Text.Trim();
            string MM003 = textBox8.Text.Trim();

            SB.AppendFormat(@" 
                            
                            DECLARE @THISYEARS nvarchar(10)
                            DECLARE @LASTYEARS nvarchar(10)
                            DECLARE @MONTHS nvarchar(10)
                            DECLARE @MM003 nvarchar(20)
                            SET @THISYEARS='{0}'
                            SET @LASTYEARS='{1}'
                            SET @MONTHS='{2}'
                            SET @MM003='{3}'

                            SELECT MM001 AS '年度',MN003 AS '月份',MM002 AS '預算編號',MM003 AS '客戶代號',MA002 AS '客戶',MM017 AS '品號',MB002 AS '品名',MN004 AS '預算數量',MN005 AS '預算金額'
                            ,(SELECT ISNULL(SUM(LA011),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH,[TK].dbo.INVLA WHERE TG001=TH001 AND TG002=TH002 AND LA006=TH001 AND LA007=TH002 AND LA008=TH003 AND TG004=MM003 AND TH004=MM017 AND TG003 LIKE @THISYEARS+@MONTHS+'%') AS '銷貨數量'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG004=MM003 AND TH004=MM017 AND TG003 LIKE @THISYEARS+@MONTHS+'%') AS '銷貨金額'
                            ,((SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG004=MM003 AND TH004=MM017 AND TG003 LIKE @THISYEARS+@MONTHS+'%')-MN005)  AS '銷貨跟預算的差異金額'
                            FROM [TK].dbo.COPMM,[TK].dbo.COPMN,[TK].dbo.COPMA ,[TK].dbo.INVMB
                            WHERE MM001=MN001 AND MM002=MN002
                            AND MA001=MM003
                            AND MB001=MM017
                            AND MM001=@THISYEARS
                            AND MM003=@MM003
                            AND MN003=@MONTHS


                            ", THISYEARS, LASTYEARS, THISMONTHS, MM003);

            return SB;

        }

        #endregion




        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT3();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT4();
        }


        #endregion


    }
}
