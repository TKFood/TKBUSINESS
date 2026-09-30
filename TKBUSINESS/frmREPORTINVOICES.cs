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
    public partial class frmREPORTINVOICES : Form
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


        public frmREPORTINVOICES()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT(string TA014)
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
            report1.Load(@"REPORT\發票OMO.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL(TA014);
            Table.SelectCommand = SQL;
            report1.Preview = previewControl1;
            report1.Show();
        }

        public string SETFASETSQL(string TA014)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            FASTSQL.AppendFormat(@"  
								--20260930匯出omo
								SELECT *
								FROM 
									(
									SELECT 
									TA001+'001' AS '交易序號1',	
									'' AS '交易序號2',	
									'' AS '交易序號3',	
									'' AS '交易序號4',	
									'' AS '交易序號5',	
									'' AS '品牌會員編號',
									TA002 AS '門市店號',	
									TA007 AS '店員編號',	
									'Normal' AS '商品類型',
									'商品一批' AS '商品名稱',	
									'商品一批' AS '商品料號'	,
									(CASE 
									WHEN ISDATE(TA001) = 1 
									THEN CAST(YEAR(CAST(TA001 AS DATETIME)) AS VARCHAR(4)) + '/' + 
										CAST(MONTH(CAST(TA001 AS DATETIME)) AS VARCHAR(2)) + '/' + 
										CAST(DAY(CAST(TA001 AS DATETIME)) AS VARCHAR(2)) + ' 10:00:00' 
									ELSE NULL 
									END)  AS '交易完成日期',
									TA017 AS '實付金額(含稅)',	
									TA017 AS '商品總售價(含稅)',	
									TA017 AS '商品單價(含稅)',	
									1 AS '購買數量',	
									0 AS '折抵金額',	
									0 AS '折扣活動總折扣金額'	,
									0 AS '折價券折扣金額',	
									0 AS 'VIP折扣金額',	
									'' AS '退貨勾稽交易序號1',
									'' AS '退貨勾稽交易序號2',
									'' AS '退貨勾稽交易序號3',
									'' AS '退貨勾稽交易序號4',
									'' AS '退貨勾稽交易序號5',
									TA014

									FROM [TK].dbo.POSTA
									WHERE TA014 LIKE '%{0}%'

									UNION ALL
									SELECT 
									TG003+'001' AS '交易序號1',	
									'' AS '交易序號2',	
									'' AS '交易序號3',	
									'' AS '交易序號4',	
									'' AS '交易序號5',	
									'' AS '品牌會員編號',
									'' AS '門市店號',	
									'' AS '店員編號',	
									'Normal' AS '商品類型',
									'商品一批' AS '商品名稱',	
									'商品一批' AS '商品料號'	,
									(CASE 
									WHEN ISDATE(TG003) = 1 
									THEN CAST(YEAR(CAST(TG003 AS DATETIME)) AS VARCHAR(4)) + '/' + 
										CAST(MONTH(CAST(TG003 AS DATETIME)) AS VARCHAR(2)) + '/' + 
										CAST(DAY(CAST(TG003 AS DATETIME)) AS VARCHAR(2)) + ' 10:00:00' 
									ELSE NULL 
									END)  AS '交易完成日期',
									TG045+TG046 AS '實付金額(含稅)',	
									TG045+TG046 AS '商品總售價(含稅)',	
									TG045+TG046 AS '商品單價(含稅)',	
									1 AS '購買數量',	
									0 AS '折抵金額',	
									0 AS '折扣活動總折扣金額'	,
									0 AS '折價券折扣金額',	
									0 AS 'VIP折扣金額',	
									'' AS '退貨勾稽交易序號1',
									'' AS '退貨勾稽交易序號2',
									'' AS '退貨勾稽交易序號3',
									'' AS '退貨勾稽交易序號4',
									'' AS '退貨勾稽交易序號5',
									TG014

									 FROM [TK].dbo.COPTG
									 WHERE  TG014  LIKE '%{0}%'
								) AS TEMP
								", TA014);

            return FASTSQL.ToString();
        }

        #endregion

        #region  
        private void button1_Click(object sender, EventArgs e)
        {
			string TA014 = textBox1.Text.ToString();
            SETFASTREPORT(TA014);
        }
        #endregion
    }
}
