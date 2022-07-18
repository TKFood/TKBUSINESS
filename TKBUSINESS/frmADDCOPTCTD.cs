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
    public partial class frmADDCOPTCTD : Form
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


        public class DATACOPMD
        {
            public string COMPANY;
            public string CREATOR;
            public string USR_GROUP;
            public string CREATE_DATE;
            public string MODIFIER;
            public string MODI_DATE;
            public string FLAG;
            public string CREATE_TIME;
            public string MODI_TIME;
            public string TRANS_TYPE;
            public string TRANS_NAME;
            public string sync_date;
            public string sync_time;
            public string sync_mark;
            public string sync_count;
            public string DataUser;
            public string DataGroup;
            public string MD001;
            public string MD002;
            public string MD003;
            public string MD004;
            public string MD005;
            public string MD006;
            public string MD007;
            public string MD008;
            public string MD009;
            public string MD010;
            public string MD011;
            public string MD012;
            public string MD013;
            public string MD014;
            public string MD015;
            public string MD016;
            public string MD017;
            public string MD018;
            public string MD019;
            public string MD020;
            public string MD021;
            public string MD022;
            public string MD023;
            public string MD024;
            public string MD025;
            public string MD026;
            public string MD027;
            public string MD028;
            public string MD029;
            public string MD030;
            public string MD031;
            public string UDF01;
            public string UDF02;
            public string UDF03;
            public string UDF04;
            public string UDF05;
            public string UDF06;
            public string UDF07;
            public string UDF08;
            public string UDF09;
            public string UDF10;
        }

        public frmADDCOPTCTD()
        {
            InitializeComponent();

            SETTEXTBOX();
        }

        #region FUNCTION

        public void SETTEXTBOX()
        {
            textBox2.Text = "2221103200";
            textBox1.Text = SERACHCOPMA(textBox2.Text.ToString().Trim());
        }
        public void Search()
        {
            DataSet ds = new DataSet();

            try
            {
                sbSql.Clear();

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.AppendFormat(@"
                                       
                                        SELECT [SERNO]
                                        ,[預購單號]
                                        ,[收件者姓名]
                                        ,[電話(日)]
                                        ,[電話(夜)]
                                        ,[手機]
                                        ,[電子郵件]
                                        ,[預定到貨日期]
                                        ,[取貨時段]
                                        ,[郵遞區號]
                                        ,[縣市]
                                        ,[鄉鎮區]
                                        ,[收件者地址]
                                        ,[產品料號]
                                        ,[商品名稱]
                                        ,[預訂數量]
                                        ,[加油站代號]
                                        ,[加油站名]
                                        FROM [TKBUSINESS].[dbo].[TEMPCOPMAORDERRS]
                                        ORDER BY [SERNO]
                                         ");

                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;

                }
                else
                {
                    dataGridView1.DataSource = ds.Tables["ds"];
                    dataGridView1.AutoResizeColumns();
                    //rownum = ds.Tables[talbename].Rows.Count - 1;                       

                    //dataGridView1.CurrentCell = dataGridView1[0, 2];

                }



            }
            catch
            {

            }
            finally
            {

            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if(string.IsNullOrEmpty(textBox2.Text.ToString().Trim()))
            {
                textBox1.Text = SERACHCOPMA(textBox2.Text.ToString().Trim());
            }
            
        }

        public string SERACHCOPMA(string MA001)
        {
            DataSet ds = new DataSet();

            try
            {
                sbSql.Clear();

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.AppendFormat(@"
                                    SELECT MA001,MA002 
                                    FROM [TK].dbo.COPMA
                                    WHERE MA001='{0}'
                                         ",MA001);

                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 0)
                {
                    return ds.Tables["ds"].Rows[0]["MA002"].ToString().Trim();
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

            }
        }


        public void ADDCOPMD()
        {
            DATACOPMD COPMD = new DATACOPMD();
            COPMD.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
            COPMD.MD001 = "2221103200";

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


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                    INSERT INTO [TK].[dbo].[COPMD]
                                    (
                                    [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]
                                    ,[MD001],[MD002],[MD003],[MD004],[MD005],[MD006],[MD007],[MD008],[MD009],[MD010]
                                    ,[MD011],[MD012],[MD013],[MD014],[MD015],[MD016],[MD017],[MD018],[MD019],[MD020]
                                    ,[MD021],[MD022],[MD023],[MD024],[MD025],[MD026],[MD027],[MD028],[MD029],[MD030]
                                    ,[MD031]
                                    ,[UDF01],[UDF02],[UDF03],[UDF04],[UDF05],[UDF06],[UDF07],[UDF08],[UDF09],[UDF10]
                                    )
                                    SELECT 
                                    'TK' AS [COMPANY],'160115' AS [CREATOR],'117000' AS [USR_GROUP],'{0}' AS [CREATE_DATE],'' AS [MODIFIER],'' AS [MODI_DATE],1 AS [FLAG],'' AS [CREATE_TIME],'' AS [MODI_TIME],'' AS [TRANS_TYPE],'' AS [TRANS_NAME],'' AS [sync_date],'' AS [sync_time],'' AS [sync_mark],0 AS [sync_count],'' AS [DataUser],'' AS [DataGroup]
                                    ,'{1}' AS [MD001],[加油站代號] AS [MD002],[縣市]+[鄉鎮區]+[收件者地址]+','+[收件者姓名]+','+[電話(日)]+'/ '+[手機] AS [MD003],'' AS [MD004],'' AS [MD005],[加油站名] AS [MD006],[收件者姓名] AS [MD007],'' AS [MD008],'' AS [MD009],'' AS [MD010]
                                    ,'' AS [MD011],[收件者姓名] AS [MD012],0 AS [MD013],0 AS [MD014],'' AS [MD015],'' AS [MD016],'' AS [MD017],[郵遞區號] AS [MD018],'' AS [MD019],'' AS [MD020]
                                    ,0 AS [MD021],'' AS [MD022],'' AS [MD023],'' AS [MD024],0 AS [MD025],'' AS [MD026],0 AS [MD027],'' AS [MD028],'' AS [MD029],'' AS [MD030]
                                    ,0 AS [MD031]
                                    ,'' AS [UDF01],'' AS [UDF02],'' AS [UDF03],'' AS [UDF04],'' AS [UDF05],0 AS [UDF06],0 AS [UDF07],0 AS [UDF08],0 AS [UDF09],0 AS [UDF10]
                                    FROM [TKBUSINESS].[dbo].[TEMPCOPMAORDERRS] 
                                    WHERE [加油站代號] NOT IN (SELECT [MD002] FROM [TK].[dbo].[COPMD] WHERE [MD001]='2221103200')
                                      
                                        ", COPMD.CREATE_DATE, COPMD.MD001);


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  

                    MessageBox.Show("完成");
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }

        }

        public DATACOPMD SETCOPMD()
        {
            DATACOPMD COPMD = new DATACOPMD();

            COPMD.COMPANY = "TK";
            COPMD.CREATOR = "160115";
            COPMD.USR_GROUP = "117000";
            COPMD.CREATE_DATE = "20220718";
            COPMD.MODIFIER = "";
            COPMD.MODI_DATE = "";
            COPMD.FLAG = "1";
            COPMD.CREATE_TIME = "";
            COPMD.MODI_TIME = "";
            COPMD.TRANS_TYPE = "";
            COPMD.TRANS_NAME = "";
            COPMD.sync_date = "";
            COPMD.sync_time = "";
            COPMD.sync_mark = "";
            COPMD.sync_count = "0";
            COPMD.DataUser = "";
            COPMD.DataGroup = "";
            COPMD.MD001 = "";
            COPMD.MD002 = "";
            COPMD.MD003 = "";
            COPMD.MD004 = "";
            COPMD.MD005 = "";
            COPMD.MD006 = "";
            COPMD.MD007 = "";
            COPMD.MD008 = "";
            COPMD.MD009 = "";
            COPMD.MD010 = "";
            COPMD.MD011 = "";
            COPMD.MD012 = "";
            COPMD.MD013 = "0";
            COPMD.MD014 = "0";
            COPMD.MD015 = "";
            COPMD.MD016 = "";
            COPMD.MD017 = "";
            COPMD.MD018 = "";
            COPMD.MD019 = "";
            COPMD.MD020 = "";
            COPMD.MD021 = "0";
            COPMD.MD022 = "";
            COPMD.MD023 = "";
            COPMD.MD024 = "";
            COPMD.MD025 = "0";
            COPMD.MD026 = "";
            COPMD.MD027 = "0";
            COPMD.MD028 = "";
            COPMD.MD029 = "";
            COPMD.MD030 = "";
            COPMD.MD031 = "0";
            COPMD.UDF01 = "";
            COPMD.UDF02 = "";
            COPMD.UDF03 = "";
            COPMD.UDF04 = "";
            COPMD.UDF05 = "";
            COPMD.UDF06 = "0";
            COPMD.UDF07 = "0";
            COPMD.UDF08 = "0";
            COPMD.UDF09 = "0";
            COPMD.UDF10 = "0";



            return COPMD;

        }

        public void ADDCOPTCCOPTD()
        {

        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            ADDCOPMD();
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        #endregion


    }
}
