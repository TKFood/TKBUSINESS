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
using System.Data.OleDb;

namespace TKBUSINESS
{
    public partial class frmTEMP91APPCOP : Form
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

        string _path = null;


        public frmTEMP91APPCOP()
        {
            InitializeComponent();

            SETTEXTBOX();


        }


        #region FUNCTION
        public void SETTEXTBOX()
        {
            textBox2.Text = "11127673";
            textBox1.Text = SERACHCOPMA(textBox2.Text.ToString().Trim());
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox2.Text.ToString().Trim()))
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
                                         ", MA001);

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


        public void Search(string YYYYMM)
        {
            DataSet ds = new DataSet();
            YYYYMM= YYYYMM.Substring(YYYYMM.Length - 4, 4);           

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
                                    
                                    SELECT 
                                    [購物車編號]
                                    ,[主單編號]
                                    ,[訂單編號]
                                    ,[轉單日期時間]
                                    ,[預計出貨日期]
                                    ,[配送方式]
                                    ,[通路商]
                                    ,[溫層類別]
                                    ,[收件人]
                                    ,[收件人電話]
                                    ,[地址]
                                    ,[門市]
                                    ,[訂單來源]
                                    ,[商品名稱]
                                    ,[商品選項]
                                    ,[商品料號]
                                    ,[數量]
                                    ,[商品單價]
                                    ,[運費]
                                    ,[配送編號]
                                    ,[狀態日期]
                                    ,[出貨單狀態]
                                    ,[訂單狀態]
                                    ,[活動代碼]
                                    ,[活動名稱]
                                    ,[折扣金額]
                                    ,[銷售金額(折扣後)]
                                    ,[付款方式]
                                    ,[活動折扣金額]
                                    ,[折價券活動序號]
                                    ,[折價券活動名稱]
                                    ,[折價券折扣金額]
                                    ,[貨到物流中心日]
                                    ,[建議貨到期限]
                                    ,[會員編號]
                                    ,[商店備註]
                                    ,[訂購備註]
                                    ,[貨運單備註]
                                    ,[驗退原因說明]
                                    ,[訂單確認日期]
                                    ,[實體會員編號]
                                    ,[商品屬性]
                                    ,[商品贈品關聯代碼]
                                    ,[購買人]
                                    ,[購買人會員等級]
                                    ,[活動對象]
                                    ,[活動會員等級]
                                    ,[總成本]
                                    ,[是否為加價購品]
                                    ,[國碼]
                                    ,[收件國家]
                                    ,[取消原因]
                                    ,[購物車總額]
                                    ,[商品頁序號]
                                    ,[點數活動名稱]
                                    ,[折抵點數]
                                    ,[點數折扣金額]
                                    ,[已設定為不可退貨商品]
                                    ,[郵遞區號]
                                    ,[指定到貨日期]
                                    ,[指定到貨時段]
                                    ,[贈品券活動序號]
                                    ,[國家地區運費活動名稱]
                                    ,[運費折扣]
                                    ,[地區/州/省份]
                                    ,[城市]
                                    ,[鄉鎮市區]
                                    ,[街道]
                                    ,[實際出貨數量]
                                    ,[實際出貨金額]
                                    ,[配送商]
                                    ,[TS重量小計(g)]
                                    ,[運費券活動序號]
                                    ,[自訂活動代碼]
                                    ,[交期]
                                    ,[線上訂單建立類型]
                                    ,[TG001]
                                    ,[TG002]
                                    ,[TH003]
                                    FROM [TKBUSINESS].[dbo].[TEMP91APPCOP]
                                    WHERE [購物車編號] LIKE '{0}%'
                                    ORDER BY [購物車編號],[主單編號],[訂單編號]
                                         ", "TG"+YYYYMM);

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

        public void UPDATETEMP91APPCOPCOPTG001TG002()
        {
            DataSet ds = new DataSet();

            //[購物車編號]  NOT IN (SELECT TG020 FROM [TK].dbo.COPTG WHERE ISNULL(TG020,'')<>'')
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
                                     SELECT
                                     [購物車編號] 
                                     FROM [TKBUSINESS].[dbo].[TEMP91APPCOP]
                                     WHERE 1=1

                                     GROUP BY [購物車編號]

                                         ");

                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 0)
                {
                    //多筆預購訂單，每陶1次新增TC001、TC002，避免重覆
                    string TG001 = "A233";
                    string TG002 = GETMAXTG002(TG001, DateTime.Now.ToString("yyyyMMdd"));

                    int serno = Convert.ToInt16(TG002.Substring(8, 3));
                    serno = serno - 1;

                    foreach (DataRow DR in ds.Tables["ds"].Rows)
                    {
                        string 購物車編號 = DR["購物車編號"].ToString();

                        //流水號+1
                        serno = serno + 1;
                        string temp = serno.ToString();
                        temp = temp.PadLeft(3, '0');
                        TG002 = DateTime.Now.ToString("yyyyMMdd") + temp.ToString();

                        UPDATETEMP91APPCOPTG001TG002(購物車編號, TG001, TG002);


                    }

                }




            }
            catch
            {

            }
            finally
            {

            }
        }

        public string GETMAXTG002(string TG001, string TG003)
        {
            SqlDataAdapter adapter4 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
            DataSet ds4 = new DataSet();
            string TC002 = null;

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
                ds4.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TG002),'00000000000') AS TG002
                                       FROM [TK].[dbo].[COPTG] 
                                       WHERE  TG001='{0}' AND TG003='{1}'
                                    ", TG001, TG003);

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        TC002 = SETTG002(ds4.Tables["TEMPds4"].Rows[0]["TG002"].ToString());
                        return TC002;

                    }
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

        public string SETTG002(string TG002)
        {
            if (TG002.Equals("00000000000"))
            {
                return DateTime.Now.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(TG002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return DateTime.Now.ToString("yyyyMMdd") + temp.ToString();
            }
        }

        public void UPDATETEMP91APPCOPTG001TG002(string 購物車編號, string TG001, string TG002)
        {
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
                                    UPDATE [TKBUSINESS].[dbo].[TEMP91APPCOP]
                                    SET TG001='{1}',TG002='{2}'
                                    WHERE  [購物車編號]='{0}'
                                        ", 購物車編號, TG001, TG002);


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

        public void UPDATETEMP91APPCOPCOPTH003()
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

                //AND [購物車編號]  NOT IN (SELECT TG020 FROM [TK].dbo.COPTG WHERE ISNULL(TG020,'')<>'')
                sbSql.AppendFormat(@"                                    
                                    SELECT 
                                    [購物車編號]
                                    ,[主單編號]
                                    ,[訂單編號]

                                    ,[TG001]
                                    ,[TG002]
                                    ,[TH003]
                                    ,RIGHT('0000'+CAST(row_number() OVER(PARTITION BY [TG001],[TG002] ORDER BY [訂單編號]) AS nvarchar(10)),4)  AS SEQ

                                    FROM [TKBUSINESS].[dbo].[TEMP91APPCOP]
                                    WHERE 1=1

                                    AND ISNULL(TG001,'')<>'' AND ISNULL(TG002,'')<>'' 
                                    ORDER BY [購物車編號],[主單編號],[訂單編號]

                                         ");

                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 0)
                {
                 

                    foreach (DataRow DR in ds.Tables["ds"].Rows)
                    {
                        string 訂單編號 = DR["訂單編號"].ToString();
                        string TH003 = DR["SEQ"].ToString();



                        UPDATETEMP91APPCOPTH003(訂單編號, TH003);


                    }

                }




            }
            catch
            {

            }
            finally
            {

            }
        }

        public void UPDATETEMP91APPCOPTH003(string 訂單編號,string TH003)
        {
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
                                    UPDATE [TKBUSINESS].[dbo].[TEMP91APPCOP]
                                    SET TH003='{1}'
                                    WHERE  [訂單編號]='{0}'
                                        ", 訂單編號, TH003);


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


        public void ADDSENDINVMB()
        {
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

                //AND [購物車編號] NOT IN (SELECT [購物車編號] FROM [TKBUSINESS].[dbo].[TEMP91APPCOP] WHERE [商品料號]='''599010000000' GROUP BY [購物車編號])
                sbSql.AppendFormat(@" 
                                   
                                    INSERT INTO [TKBUSINESS].[dbo].[TEMP91APPCOP]
                                    (
                                    [購物車編號]
                                    ,[主單編號]
                                    ,[訂單編號]
                                    ,[轉單日期時間]
                                    ,[預計出貨日期]
                                    ,[配送方式]
                                    ,[通路商]
                                    ,[溫層類別]
                                    ,[收件人]
                                    ,[收件人電話]
                                    ,[地址]
                                    ,[門市]
                                    ,[訂單來源]
                                    ,[商品名稱]
                                    ,[商品選項]
                                    ,[商品料號]
                                    ,[數量]
                                    ,[商品單價]
                                    ,[運費]
                                    ,[配送編號]
                                    ,[狀態日期]
                                    ,[出貨單狀態]
                                    ,[訂單狀態]
                                    ,[活動代碼]
                                    ,[活動名稱]
                                    ,[折扣金額]
                                    ,[銷售金額(折扣後)]
                                    ,[付款方式]
                                    ,[活動折扣金額]
                                    ,[折價券活動序號]
                                    ,[折價券活動名稱]
                                    ,[折價券折扣金額]
                                    ,[貨到物流中心日]
                                    ,[建議貨到期限]
                                    ,[會員編號]
                                    ,[商店備註]
                                    ,[訂購備註]
                                    ,[貨運單備註]
                                    ,[驗退原因說明]
                                    ,[訂單確認日期]
                                    ,[實體會員編號]
                                    ,[商品屬性]
                                    ,[商品贈品關聯代碼]
                                    ,[購買人]
                                    ,[購買人會員等級]
                                    ,[活動對象]
                                    ,[活動會員等級]
                                    ,[總成本]
                                    ,[是否為加價購品]
                                    ,[國碼]
                                    ,[收件國家]
                                    ,[取消原因]
                                    ,[購物車總額]
                                    ,[商品頁序號]
                                    ,[點數活動名稱]
                                    ,[折抵點數]
                                    ,[點數折扣金額]
                                    ,[已設定為不可退貨商品]
                                    ,[郵遞區號]
                                    ,[指定到貨日期]
                                    ,[指定到貨時段]
                                    ,[贈品券活動序號]
                                    ,[國家地區運費活動名稱]
                                    ,[運費折扣]
                                    ,[地區/州/省份]
                                    ,[城市]
                                    ,[鄉鎮市區]
                                    ,[街道]
                                    ,[實際出貨數量]
                                    ,[實際出貨金額]
                                    ,[配送商]
                                    ,[TS重量小計(g)]
                                    ,[運費券活動序號]
                                    ,[自訂活動代碼]
                                    ,[交期]
                                    ,[線上訂單建立類型]

                                    )
                                    SELECT 
                                    [購物車編號]
                                    ,[主單編號]
                                    ,''+[購物車編號]+'A'+CONVERT(NVARCHAR,row_number() OVER(PARTITION BY [購物車編號] ORDER BY [購物車編號])) As [訂單編號]
                                    ,[轉單日期時間]
                                    ,[預計出貨日期]
                                    ,[配送方式]
                                    ,[通路商]
                                    ,[溫層類別]
                                    ,[收件人]
                                    ,[收件人電話]
                                    ,[地址]
                                    ,[門市]
                                    ,[訂單來源]
                                    ,'' [商品名稱]
                                    ,'' [商品選項]
                                    ,'''599010000000' [商品料號]
                                    ,1 [數量]
                                    ,[運費] [商品單價]
                                    ,0 [運費]
                                    ,[配送編號]
                                    ,[狀態日期]
                                    ,[出貨單狀態]
                                    ,[訂單狀態]
                                    ,[活動代碼]
                                    ,[活動名稱]
                                    ,0 [折扣金額]
                                    ,0 [銷售金額(折扣後)]
                                    ,[付款方式]
                                    ,0 [活動折扣金額]
                                    ,[折價券活動序號]
                                    ,[折價券活動名稱]
                                    ,0 [折價券折扣金額]
                                    ,[貨到物流中心日]
                                    ,[建議貨到期限]
                                    ,[會員編號]
                                    ,[商店備註]
                                    ,[訂購備註]
                                    ,[貨運單備註]
                                    ,[驗退原因說明]
                                    ,[訂單確認日期]
                                    ,[實體會員編號]
                                    ,[商品屬性]
                                    ,[商品贈品關聯代碼]
                                    ,[購買人]
                                    ,[購買人會員等級]
                                    ,[活動對象]
                                    ,[活動會員等級]
                                    ,[總成本]
                                    ,[是否為加價購品]
                                    ,[國碼]
                                    ,[收件國家]
                                    ,[取消原因]
                                    ,[購物車總額]
                                    ,'' [商品頁序號]
                                    ,[點數活動名稱]
                                    ,0 [折抵點數]
                                    ,0 [點數折扣金額]
                                    ,[已設定為不可退貨商品]
                                    ,[郵遞區號]
                                    ,[指定到貨日期]
                                    ,[指定到貨時段]
                                    ,[贈品券活動序號]
                                    ,[國家地區運費活動名稱]
                                    ,[運費折扣]
                                    ,[地區/州/省份]
                                    ,[城市]
                                    ,[鄉鎮市區]
                                    ,[街道]
                                    ,1 [實際出貨數量]
                                    ,[運費] [實際出貨金額]
                                    ,[配送商]
                                    ,0 [TS重量小計(g)]
                                    ,[運費券活動序號]
                                    ,[自訂活動代碼]
                                    ,[交期]
                                    ,[線上訂單建立類型]
                                    FROM [TKBUSINESS].[dbo].[TEMP91APPCOP]
                                    WHERE [運費]>0 


                                    GROUP  BY [購物車編號]
                                    ,[主單編號]
                                    ,[轉單日期時間]
                                    ,[預計出貨日期]
                                    ,[配送方式]
                                    ,[通路商]
                                    ,[溫層類別]
                                    ,[收件人]
                                    ,[收件人電話]
                                    ,[地址]
                                    ,[門市]
                                    ,[訂單來源]
                                    ,[運費]
                                    ,[配送編號]
                                    ,[狀態日期]
                                    ,[出貨單狀態]
                                    ,[訂單狀態]
                                    ,[活動代碼]
                                    ,[活動名稱]
                                    ,[付款方式]
                                    ,[折價券活動序號]
                                    ,[折價券活動名稱]
                                    ,[貨到物流中心日]
                                    ,[建議貨到期限]
                                    ,[會員編號]
                                    ,[商店備註]
                                    ,[訂購備註]
                                    ,[貨運單備註]
                                    ,[驗退原因說明]
                                    ,[訂單確認日期]
                                    ,[實體會員編號]
                                    ,[商品屬性]
                                    ,[商品贈品關聯代碼]
                                    ,[購買人]
                                    ,[購買人會員等級]
                                    ,[活動對象]
                                    ,[活動會員等級]
                                    ,[總成本]
                                    ,[是否為加價購品]
                                    ,[國碼]
                                    ,[收件國家]
                                    ,[取消原因]
                                    ,[購物車總額]
                                    ,[點數活動名稱]                                   
                                    ,[已設定為不可退貨商品]
                                    ,[郵遞區號]
                                    ,[指定到貨日期]
                                    ,[指定到貨時段]
                                    ,[贈品券活動序號]
                                    ,[國家地區運費活動名稱]
                                    ,[運費折扣]
                                    ,[地區/州/省份]
                                    ,[城市]
                                    ,[鄉鎮市區]
                                    ,[街道]
                                    ,[配送商]
                                    ,[運費券活動序號]
                                    ,[自訂活動代碼]
                                    ,[交期]
                                    ,[線上訂單建立類型]
                                    ORDER BY [購物車編號]

                                        ");


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

        public void ADDDISCOUNTINVMB()
        {
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

                //AND [購物車編號] NOT IN (SELECT [購物車編號] FROM [TKBUSINESS].[dbo].[TEMP91APPCOP] WHERE [商品料號]='''599030000004' GROUP BY [購物車編號])

                sbSql.AppendFormat(@" 
                                   
                                    INSERT INTO [TKBUSINESS].[dbo].[TEMP91APPCOP]
                                    (
                                    [購物車編號]
                                    ,[主單編號]
                                    ,[訂單編號]
                                    ,[轉單日期時間]
                                    ,[預計出貨日期]
                                    ,[配送方式]
                                    ,[通路商]
                                    ,[溫層類別]
                                    ,[收件人]
                                    ,[收件人電話]
                                    ,[地址]
                                    ,[門市]
                                    ,[訂單來源]
                                    ,[商品名稱]
                                    ,[商品選項]
                                    ,[商品料號]
                                    ,[數量]
                                    ,[商品單價]
                                    ,[運費]
                                    ,[配送編號]
                                    ,[狀態日期]
                                    ,[出貨單狀態]
                                    ,[訂單狀態]
                                    ,[活動代碼]
                                    ,[活動名稱]
                                    ,[折扣金額]
                                    ,[銷售金額(折扣後)]
                                    ,[付款方式]
                                    ,[活動折扣金額]
                                    ,[折價券活動序號]
                                    ,[折價券活動名稱]
                                    ,[折價券折扣金額]
                                    ,[貨到物流中心日]
                                    ,[建議貨到期限]
                                    ,[會員編號]
                                    ,[商店備註]
                                    ,[訂購備註]
                                    ,[貨運單備註]
                                    ,[驗退原因說明]
                                    ,[訂單確認日期]
                                    ,[實體會員編號]
                                    ,[商品屬性]
                                    ,[商品贈品關聯代碼]
                                    ,[購買人]
                                    ,[購買人會員等級]
                                    ,[活動對象]
                                    ,[活動會員等級]
                                    ,[總成本]
                                    ,[是否為加價購品]
                                    ,[國碼]
                                    ,[收件國家]
                                    ,[取消原因]
                                    ,[購物車總額]
                                    ,[商品頁序號]
                                    ,[點數活動名稱]
                                    ,[折抵點數]
                                    ,[點數折扣金額]
                                    ,[已設定為不可退貨商品]
                                    ,[郵遞區號]
                                    ,[指定到貨日期]
                                    ,[指定到貨時段]
                                    ,[贈品券活動序號]
                                    ,[國家地區運費活動名稱]
                                    ,[運費折扣]
                                    ,[地區/州/省份]
                                    ,[城市]
                                    ,[鄉鎮市區]
                                    ,[街道]
                                    ,[實際出貨數量]
                                    ,[實際出貨金額]
                                    ,[配送商]
                                    ,[TS重量小計(g)]
                                    ,[運費券活動序號]
                                    ,[自訂活動代碼]
                                    ,[交期]
                                    ,[線上訂單建立類型]

                                    )
                                    
                                    SELECT 
                                    [購物車編號]
                                    ,[主單編號]
                                    ,''+[購物車編號]+'B'+CONVERT(NVARCHAR,row_number() OVER(PARTITION BY [購物車編號] ORDER BY [購物車編號])) As [訂單編號]
                                    ,[轉單日期時間]
                                    ,[預計出貨日期]
                                    ,[配送方式]
                                    ,[通路商]
                                    ,[溫層類別]
                                    ,[收件人]
                                    ,[收件人電話]
                                    ,[地址]
                                    ,[門市]
                                    ,[訂單來源]
                                    ,'' [商品名稱]
                                    ,'' [商品選項]
                                    ,'''599030000004' [商品料號]
                                    ,1 [數量]
                                    ,(SUM(CONVERT(INT,[折扣金額]))-SUM(CONVERT(INT,[點數折扣金額]))) [商品單價]
                                    ,0 [運費]
                                    ,[配送編號]
                                    ,[狀態日期]
                                    ,[出貨單狀態]
                                    ,[訂單狀態]
                                    ,[活動代碼]
                                    ,[活動名稱]
                                    ,0 [折扣金額]
                                    ,0 [銷售金額(折扣後)]
                                    ,[付款方式]
                                    ,0 [活動折扣金額]
                                    ,[折價券活動序號]
                                    ,[折價券活動名稱]
                                    ,0 [折價券折扣金額]
                                    ,[貨到物流中心日]
                                    ,[建議貨到期限]
                                    ,[會員編號]
                                    ,[商店備註]
                                    ,[訂購備註]
                                    ,[貨運單備註]
                                    ,[驗退原因說明]
                                    ,[訂單確認日期]
                                    ,[實體會員編號]
                                    ,[商品屬性]
                                    ,[商品贈品關聯代碼]
                                    ,[購買人]
                                    ,[購買人會員等級]
                                    ,[活動對象]
                                    ,[活動會員等級]
                                    ,[總成本]
                                    ,[是否為加價購品]
                                    ,[國碼]
                                    ,[收件國家]
                                    ,[取消原因]
                                    ,[購物車總額]
                                    ,'' [商品頁序號]
                                    ,[點數活動名稱]
                                    ,0 [折抵點數]
                                    ,0 [點數折扣金額]
                                    ,[已設定為不可退貨商品]
                                    ,[郵遞區號]
                                    ,[指定到貨日期]
                                    ,[指定到貨時段]
                                    ,[贈品券活動序號]
                                    ,[國家地區運費活動名稱]
                                    ,[運費折扣]
                                    ,[地區/州/省份]
                                    ,[城市]
                                    ,[鄉鎮市區]
                                    ,[街道]
                                    ,1 [實際出貨數量]
                                    ,[運費] [實際出貨金額]
                                    ,[配送商]
                                    ,0 [TS重量小計(g)]
                                    ,[運費券活動序號]
                                    ,[自訂活動代碼]
                                    ,[交期]
                                    ,[線上訂單建立類型]
                                    FROM [TKBUSINESS].[dbo].[TEMP91APPCOP]
                                    WHERE [折扣金額]<>0 


                                    GROUP  BY [購物車編號]
                                    ,[主單編號]
                                    ,[轉單日期時間]
                                    ,[預計出貨日期]
                                    ,[配送方式]
                                    ,[通路商]
                                    ,[溫層類別]
                                    ,[收件人]
                                    ,[收件人電話]
                                    ,[地址]
                                    ,[門市]
                                    ,[訂單來源]
                                    ,[運費]
                                    ,[配送編號]
                                    ,[狀態日期]
                                    ,[出貨單狀態]
                                    ,[訂單狀態]
                                    ,[活動代碼]
                                    ,[活動名稱]
                                    ,[付款方式]
                                    ,[折價券活動序號]
                                    ,[折價券活動名稱]
                                    ,[貨到物流中心日]
                                    ,[建議貨到期限]
                                    ,[會員編號]
                                    ,[商店備註]
                                    ,[訂購備註]
                                    ,[貨運單備註]
                                    ,[驗退原因說明]
                                    ,[訂單確認日期]
                                    ,[實體會員編號]
                                    ,[商品屬性]
                                    ,[商品贈品關聯代碼]
                                    ,[購買人]
                                    ,[購買人會員等級]
                                    ,[活動對象]
                                    ,[活動會員等級]
                                    ,[總成本]
                                    ,[是否為加價購品]
                                    ,[國碼]
                                    ,[收件國家]
                                    ,[取消原因]
                                    ,[購物車總額]
                                    ,[點數活動名稱]                                   
                                    ,[已設定為不可退貨商品]
                                    ,[郵遞區號]
                                    ,[指定到貨日期]
                                    ,[指定到貨時段]
                                    ,[贈品券活動序號]
                                    ,[國家地區運費活動名稱]
                                    ,[運費折扣]
                                    ,[地區/州/省份]
                                    ,[城市]
                                    ,[鄉鎮市區]
                                    ,[街道]
                                    ,[配送商]
                                    ,[運費券活動序號]
                                    ,[自訂活動代碼]
                                    ,[交期]
                                    ,[線上訂單建立類型]
                                    ORDER BY [購物車編號]
                                        ");


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


        public void ADDDISCOUNTINVMB2()
        {
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

                //                                    AND [購物車編號] NOT IN (SELECT [購物車編號] FROM [TKBUSINESS].[dbo].[TEMP91APPCOP] WHERE [商品料號]='''599030000004' GROUP BY [購物車編號])
                sbSql.AppendFormat(@" 
                                   
                                    INSERT INTO [TKBUSINESS].[dbo].[TEMP91APPCOP]
                                    (
                                    [購物車編號]
                                    ,[主單編號]
                                    ,[訂單編號]
                                    ,[轉單日期時間]
                                    ,[預計出貨日期]
                                    ,[配送方式]
                                    ,[通路商]
                                    ,[溫層類別]
                                    ,[收件人]
                                    ,[收件人電話]
                                    ,[地址]
                                    ,[門市]
                                    ,[訂單來源]
                                    ,[商品名稱]
                                    ,[商品選項]
                                    ,[商品料號]
                                    ,[數量]
                                    ,[商品單價]
                                    ,[運費]
                                    ,[配送編號]
                                    ,[狀態日期]
                                    ,[出貨單狀態]
                                    ,[訂單狀態]
                                    ,[活動代碼]
                                    ,[活動名稱]
                                    ,[折扣金額]
                                    ,[銷售金額(折扣後)]
                                    ,[付款方式]
                                    ,[活動折扣金額]
                                    ,[折價券活動序號]
                                    ,[折價券活動名稱]
                                    ,[折價券折扣金額]
                                    ,[貨到物流中心日]
                                    ,[建議貨到期限]
                                    ,[會員編號]
                                    ,[商店備註]
                                    ,[訂購備註]
                                    ,[貨運單備註]
                                    ,[驗退原因說明]
                                    ,[訂單確認日期]
                                    ,[實體會員編號]
                                    ,[商品屬性]
                                    ,[商品贈品關聯代碼]
                                    ,[購買人]
                                    ,[購買人會員等級]
                                    ,[活動對象]
                                    ,[活動會員等級]
                                    ,[總成本]
                                    ,[是否為加價購品]
                                    ,[國碼]
                                    ,[收件國家]
                                    ,[取消原因]
                                    ,[購物車總額]
                                    ,[商品頁序號]
                                    ,[點數活動名稱]
                                    ,[折抵點數]
                                    ,[點數折扣金額]
                                    ,[已設定為不可退貨商品]
                                    ,[郵遞區號]
                                    ,[指定到貨日期]
                                    ,[指定到貨時段]
                                    ,[贈品券活動序號]
                                    ,[國家地區運費活動名稱]
                                    ,[運費折扣]
                                    ,[地區/州/省份]
                                    ,[城市]
                                    ,[鄉鎮市區]
                                    ,[街道]
                                    ,[實際出貨數量]
                                    ,[實際出貨金額]
                                    ,[配送商]
                                    ,[TS重量小計(g)]
                                    ,[運費券活動序號]
                                    ,[自訂活動代碼]
                                    ,[交期]
                                    ,[線上訂單建立類型]

                                    )
                                    
                                    SELECT 
                                    [購物車編號]
                                    ,[主單編號]
                                    ,''+[購物車編號]+'C'+CONVERT(NVARCHAR,row_number() OVER(PARTITION BY [購物車編號] ORDER BY [購物車編號])) As [訂單編號]
                                    ,[轉單日期時間]
                                    ,[預計出貨日期]
                                    ,[配送方式]
                                    ,[通路商]
                                    ,[溫層類別]
                                    ,[收件人]
                                    ,[收件人電話]
                                    ,[地址]
                                    ,[門市]
                                    ,[訂單來源]
                                    ,'' [商品名稱]
                                    ,'' [商品選項]
                                    ,'''599030000003' [商品料號]
                                    ,SUM(CONVERT(INT,[點數折扣金額]))*-1 [數量]
                                    ,-1 [商品單價]
                                    ,0 [運費]
                                    ,[配送編號]
                                    ,[狀態日期]
                                    ,[出貨單狀態]
                                    ,[訂單狀態]
                                    ,[活動代碼]
                                    ,[活動名稱]
                                    ,0 [折扣金額]
                                    ,0 [銷售金額(折扣後)]
                                    ,[付款方式]
                                    ,0 [活動折扣金額]
                                    ,[折價券活動序號]
                                    ,[折價券活動名稱]
                                    ,0 [折價券折扣金額]
                                    ,[貨到物流中心日]
                                    ,[建議貨到期限]
                                    ,[會員編號]
                                    ,[商店備註]
                                    ,[訂購備註]
                                    ,[貨運單備註]
                                    ,[驗退原因說明]
                                    ,[訂單確認日期]
                                    ,[實體會員編號]
                                    ,[商品屬性]
                                    ,[商品贈品關聯代碼]
                                    ,[購買人]
                                    ,[購買人會員等級]
                                    ,[活動對象]
                                    ,[活動會員等級]
                                    ,[總成本]
                                    ,[是否為加價購品]
                                    ,[國碼]
                                    ,[收件國家]
                                    ,[取消原因]
                                    ,[購物車總額]
                                    ,'' [商品頁序號]
                                    ,[點數活動名稱]
                                    ,0 [折抵點數]
                                    ,0 [點數折扣金額]
                                    ,[已設定為不可退貨商品]
                                    ,[郵遞區號]
                                    ,[指定到貨日期]
                                    ,[指定到貨時段]
                                    ,[贈品券活動序號]
                                    ,[國家地區運費活動名稱]
                                    ,[運費折扣]
                                    ,[地區/州/省份]
                                    ,[城市]
                                    ,[鄉鎮市區]
                                    ,[街道]
                                    ,1 [實際出貨數量]
                                    ,[運費] [實際出貨金額]
                                    ,[配送商]
                                    ,0 [TS重量小計(g)]
                                    ,[運費券活動序號]
                                    ,[自訂活動代碼]
                                    ,[交期]
                                    ,[線上訂單建立類型]
                                    FROM [TKBUSINESS].[dbo].[TEMP91APPCOP]
                                    WHERE [折扣金額]<>0 

                                  
                                    GROUP  BY [購物車編號]
                                    ,[主單編號]
                                    ,[轉單日期時間]
                                    ,[預計出貨日期]
                                    ,[配送方式]
                                    ,[通路商]
                                    ,[溫層類別]
                                    ,[收件人]
                                    ,[收件人電話]
                                    ,[地址]
                                    ,[門市]
                                    ,[訂單來源]
                                    ,[運費]
                                    ,[配送編號]
                                    ,[狀態日期]
                                    ,[出貨單狀態]
                                    ,[訂單狀態]
                                    ,[活動代碼]
                                    ,[活動名稱]
                                    ,[付款方式]
                                    ,[折價券活動序號]
                                    ,[折價券活動名稱]
                                    ,[貨到物流中心日]
                                    ,[建議貨到期限]
                                    ,[會員編號]
                                    ,[商店備註]
                                    ,[訂購備註]
                                    ,[貨運單備註]
                                    ,[驗退原因說明]
                                    ,[訂單確認日期]
                                    ,[實體會員編號]
                                    ,[商品屬性]
                                    ,[商品贈品關聯代碼]
                                    ,[購買人]
                                    ,[購買人會員等級]
                                    ,[活動對象]
                                    ,[活動會員等級]
                                    ,[總成本]
                                    ,[是否為加價購品]
                                    ,[國碼]
                                    ,[收件國家]
                                    ,[取消原因]
                                    ,[購物車總額]
                                    ,[點數活動名稱]                                   
                                    ,[已設定為不可退貨商品]
                                    ,[郵遞區號]
                                    ,[指定到貨日期]
                                    ,[指定到貨時段]
                                    ,[贈品券活動序號]
                                    ,[國家地區運費活動名稱]
                                    ,[運費折扣]
                                    ,[地區/州/省份]
                                    ,[城市]
                                    ,[鄉鎮市區]
                                    ,[街道]
                                    ,[配送商]
                                    ,[運費券活動序號]
                                    ,[自訂活動代碼]
                                    ,[交期]
                                    ,[線上訂單建立類型]
                                    ORDER BY [購物車編號]

                                        ");


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

        public void ADDERPCOPTGCOPTH()
        {
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

                
                //COPTG --AND [購物車編號]  NOT IN (SELECT TG020 FROM [TK].dbo.COPTG WHERE ISNULL(TG020,'')<>'')
                //COPTH --AND [訂單編號]  NOT IN (SELECT TH074 FROM [TK].dbo.COPTH WHERE ISNULL(TH074,'')<>'')
                sbSql.AppendFormat(@" 
                                   
                                   
                                        --COPTG
                                        INSERT INTO [test0923].[dbo].[COPTG]
                                        (
                                         [COMPANY]
                                        ,[CREATOR]
                                        ,[USR_GROUP]
                                        ,[CREATE_DATE]
                                        ,[MODIFIER]
                                        ,[MODI_DATE]
                                        ,[FLAG]
                                        ,[CREATE_TIME]
                                        ,[MODI_TIME]
                                        ,[TRANS_TYPE]
                                        ,[TRANS_NAME]
                                        ,[sync_date]
                                        ,[sync_time]
                                        ,[sync_mark]
                                        ,[sync_count]
                                        ,[DataUser]
                                        ,[DataGroup]
                                        ,[TG001]
                                        ,[TG002]
                                        ,[TG003]
                                        ,[TG004]
                                        ,[TG005]
                                        ,[TG006]
                                        ,[TG007]
                                        ,[TG008]
                                        ,[TG009]
                                        ,[TG010]
                                        ,[TG011]
                                        ,[TG012]
                                        ,[TG013]
                                        ,[TG014]
                                        ,[TG015]
                                        ,[TG016]
                                        ,[TG017]
                                        ,[TG018]
                                        ,[TG019]
                                        ,[TG020]
                                        ,[TG021]
                                        ,[TG022]
                                        ,[TG023]
                                        ,[TG024]
                                        ,[TG025]
                                        ,[TG026]
                                        ,[TG027]
                                        ,[TG028]
                                        ,[TG029]
                                        ,[TG030]
                                        ,[TG031]
                                        ,[TG032]
                                        ,[TG033]
                                        ,[TG034]
                                        ,[TG035]
                                        ,[TG036]
                                        ,[TG037]
                                        ,[TG038]
                                        ,[TG039]
                                        ,[TG040]
                                        ,[TG041]
                                        ,[TG042]
                                        ,[TG043]
                                        ,[TG044]
                                        ,[TG045]
                                        ,[TG046]
                                        ,[TG047]
                                        ,[TG048]
                                        ,[TG049]
                                        ,[TG050]
                                        ,[TG051]
                                        ,[TG052]
                                        ,[TG053]
                                        ,[TG054]
                                        ,[TG055]
                                        ,[TG056]
                                        ,[TG057]
                                        ,[TG058]
                                        ,[TG059]
                                        ,[TG060]
                                        ,[TG061]
                                        ,[TG062]
                                        ,[TG063]
                                        ,[TG064]
                                        ,[TG065]
                                        ,[TG066]
                                        ,[TG067]
                                        ,[TG068]
                                        ,[TG069]
                                        ,[TG070]
                                        ,[TG071]
                                        ,[TG072]
                                        ,[TG073]
                                        ,[TG074]
                                        ,[TG075]
                                        ,[TG076]
                                        ,[TG077]
                                        ,[TG078]
                                        ,[TG079]
                                        ,[TG080]
                                        ,[TG081]
                                        ,[TG082]
                                        ,[TG083]
                                        ,[TG084]
                                        ,[TG085]
                                        ,[TG086]
                                        ,[TG087]
                                        ,[TG088]
                                        ,[TG089]
                                        ,[TG090]
                                        ,[TG091]
                                        ,[TG092]
                                        ,[TG093]
                                        ,[TG094]
                                        ,[TG095]
                                        ,[TG096]
                                        ,[TG097]
                                        ,[TG098]
                                        ,[TG099]
                                        ,[TG100]
                                        ,[TG101]
                                        ,[TG102]
                                        ,[TG103]
                                        ,[TG104]
                                        ,[TG105]
                                        ,[TG106]
                                        ,[TG107]
                                        ,[TG108]
                                        ,[TG109]
                                        ,[TG110]
                                        ,[TG111]
                                        ,[TG112]
                                        ,[TG113]
                                        ,[TG114]
                                        ,[TG115]
                                        ,[TG116]
                                        ,[TG117]
                                        ,[TG118]
                                        ,[TG119]
                                        ,[TG120]
                                        ,[TG121]
                                        ,[TG122]
                                        ,[TG123]
                                        ,[TG124]
                                        ,[TG125]
                                        ,[TG126]
                                        ,[TG127]
                                        ,[TG128]
                                        ,[TG129]
                                        ,[TG130]
                                        ,[TG131]
                                        ,[TG132]
                                        ,[TG133]
                                        ,[TG134]
                                        ,[TG135]
                                        ,[TG136]
                                        ,[TG137]
                                        ,[TG138]
                                        ,[TG139]
                                        ,[TG140]
                                        ,[TG141]
                                        ,[TG142]
                                        ,[TG143]
                                        ,[TG144]
                                        ,[TG145]
                                        ,[TG146]
                                        ,[TG147]
                                        ,[TG148]
                                        ,[TG149]
                                        ,[TG150]
                                        ,[TG151]
                                        ,[TG152]
                                        ,[TG153]
                                        ,[TG154]
                                        ,[TG155]
                                        ,[TG156]
                                        ,[TG157]
                                        ,[TG158]
                                        ,[TG159]
                                        ,[TG160]
                                        ,[TG161]
                                        ,[TG162]
                                        ,[TG163]
                                        ,[UDF01]
                                        ,[UDF02]
                                        ,[UDF03]
                                        ,[UDF04]
                                        ,[UDF05]
                                        ,[UDF06]
                                        ,[UDF07]
                                        ,[UDF08]
                                        ,[UDF09]
                                        ,[UDF10]
                                        )

                                        SELECT
                                         'TK' AS [COMPANY]
                                        ,'190083' AS [CREATOR]
                                        ,'102602' AS [USR_GROUP]
                                        ,CONVERT(NVARCHAR,GETDATE(),112) AS [CREATE_DATE]
                                        ,'' AS [MODIFIER]
                                        ,'' AS [MODI_DATE]
                                        ,1 AS [FLAG]
                                        ,CONVERT(NVARCHAR,GETDATE(),108) AS [CREATE_TIME]
                                        ,'' AS [MODI_TIME]
                                        ,'P001' AS [TRANS_TYPE]
                                        ,'COPMI08' AS [TRANS_NAME]
                                        ,'' AS [sync_date]
                                        ,'' AS [sync_time]
                                        ,'' AS [sync_mark]
                                        , 0 AS [sync_count]
                                        ,'' AS [DataUser]
                                        ,'102602' AS [DataGroup]
                                        ,[TEMP91APPCOP].TG001 AS [TG001]
                                        ,[TEMP91APPCOP].TG002 AS [TG002]
                                        ,CONVERT(NVARCHAR,GETDATE(),112) AS [TG003]
                                        ,MA001 AS [TG004]
                                        ,'117300' AS [TG005]
                                        ,'170007' AS [TG006]
                                        ,MA003 AS [TG007]
                                        ,(
                                        CASE WHEN [配送方式]='宅配' THEN [地址]+','+[收件人]+','+[收件人電話]
                                        WHEN [配送方式]='貨到付款' THEN [地址]+','+[收件人]+','+[收件人電話]
                                        WHEN [配送方式]='超商取貨付款' AND [通路商]='7-11(統一)' THEN '新北市樹林區佳園路2段70-1號(大智通),EC驗收組,02-26738186#101'
                                        WHEN [配送方式]='付款後超商取貨' AND [通路商]='7-11(統一)' THEN '新北市樹林區佳園路2段70-1號(大智通),EC驗收組,02-26738186#101'
                                        WHEN [配送方式]='超商取貨付款' AND [通路商]='全家' THEN '桃園市大溪區仁善里15鄰新光東路76巷22-2號(日翊文化),電子商務部,03-3075581'
                                        WHEN [配送方式]='付款後超商取貨' AND [通路商]='全家' THEN '桃園市大溪區仁善里15鄰新光東路76巷22-2號(日翊文化),電子商務部,03-3075581'
                                        WHEN [配送方式]='超商取貨付款' AND [通路商]='萊爾富' THEN '桃園市龍潭區中原路二段545巷146號(萊爾富),物流中心EC收件組,03-2865168#611'
                                        WHEN [配送方式]='付款後超商取貨' AND [通路商]='萊爾富' THEN '桃園市龍潭區中原路二段545巷146號(萊爾富),物流中心EC收件組,03-2865168#611'
                                        END
                                        ) AS [TG008]
                                        ,'' AS [TG009]
                                        ,'20' AS [TG010]
                                        ,'NTD' AS [TG011]
                                        ,1 AS [TG012]
                                        ,ROUND(CONVERT(INT,[購物車總額]) /1.05,0)AS [TG013]
                                        ,'' AS [TG014]
                                        ,MA010 AS [TG015]
                                        ,MA037 AS [TG016]
                                        ,'1' AS [TG017]
                                        ,MA025 AS [TG018]
                                        ,'' AS [TG019]
                                        ,[購物車編號] AS [TG020]
                                        ,'' AS [TG021]
                                        ,1 AS [TG022]
                                        ,'N' AS [TG023]
                                        ,'N' AS [TG024]
                                        ,(CONVERT(INT,[購物車總額])-ROUND(CONVERT(INT,[購物車總額]) /1.05,0)) AS [TG025]
                                        ,'170007' AS [TG026]
                                        ,'' AS [TG027]
                                        ,'' AS [TG028]
                                        ,'' AS [TG029]
                                        ,'N' AS [TG030]
                                        ,'1' AS [TG031]
                                        ,'0' AS [TG032]
                                        ,SUM(CONVERT(INT,[數量])) AS [TG033]
                                        ,'N' AS [TG034]
                                        ,'' AS [TG035]
                                        ,'N' AS [TG036]
                                        ,'N' AS [TG037]
                                        ,LEFT(CONVERT(NVARCHAR,GETDATE(),112),6) AS [TG038]
                                        ,'' AS [TG039]
                                        ,'' AS [TG040]
                                        ,0 AS [TG041]
                                        ,CONVERT(NVARCHAR,GETDATE(),112) AS [TG042]
                                        ,''  AS [TG043]
                                        ,0.0500 AS [TG044]
                                        ,ROUND(CONVERT(INT,[購物車總額]) /1.05,0) AS [TG045]
                                        ,(CONVERT(INT,[購物車總額])-ROUND(CONVERT(INT,[購物車總額]) /1.05,0)) AS [TG046]
                                        ,'204' AS [TG047]
                                        ,'' AS [TG048]
                                        ,'' AS [TG049]
                                        ,'' AS [TG050]
                                        ,'' AS [TG051]
                                        ,0 AS [TG052]
                                        ,0 AS [TG053]
                                        ,0 AS [TG054]
                                        ,'N' AS [TG055]
                                        ,'N' AS [TG056]
                                        ,'' AS [TG057]
                                        ,'' AS [TG058]
                                        ,'N' AS [TG059]
                                        ,'' AS [TG060]
                                        ,'N' AS [TG061]
                                        ,'N' AS [TG062]
                                        ,0 AS [TG063]
                                        ,'' AS [TG064]
                                        ,'' AS [TG065]
                                        ,MA003  AS [TG066]
                                        ,'' AS [TG067]
                                        ,'1' AS [TG068]
                                        ,0 AS [TG069]
                                        ,'N' AS [TG070]
                                        ,0 AS [TG071]
                                        ,'5' AS [TG072]
                                        ,'' AS [TG073]
                                        ,'' AS [TG074]
                                        ,[通路商]+'-'+[配送方式]+'-'+[收件人] AS [TG075]
                                        ,'' AS [TG076]
                                        ,'' AS [TG077]
                                        ,'' AS [TG078]
                                        ,'' AS [TG079]
                                        ,'' AS [TG080]
                                        ,'' AS [TG081]
                                        ,'' AS [TG082]
                                        ,'' AS [TG083]
                                        ,0 AS [TG084]
                                        ,0 AS [TG085]
                                        ,'' AS [TG086]
                                        ,'' AS [TG087]
                                        ,'' AS [TG088]
                                        ,'N' AS [TG089]
                                        ,'N' AS [TG090]
                                        ,0 AS [TG091]
                                        ,'N' AS [TG092]
                                        ,'' AS [TG093]
                                        ,'' AS [TG094]
                                        ,'' AS [TG095]
                                        ,'' AS [TG096]
                                        ,'N' AS [TG097]
                                        ,'' AS [TG098]
                                        ,0 AS [TG099]
                                        ,'N' AS [TG100]
                                        ,0 AS [TG101]
                                        ,'' AS [TG102]
                                        ,'' AS [TG103]
                                        ,'N' AS [TG104]
                                        ,'' AS [TG105]
                                        ,'' AS [TG106]
                                        ,'' AS [TG107]
                                        ,'' AS [TG108]
                                        ,'' AS [TG109]
                                        ,CONVERT(NVARCHAR,DATEADD(day, 1, GETDATE()),112) AS [TG110]
                                        ,'1' AS [TG111]
                                        ,'' AS [TG112]
                                        ,0 AS [TG113]
                                        ,0 AS [TG114]
                                        ,'N' AS [TG115]
                                        ,'N' AS [TG116]
                                        ,'' AS [TG117]
                                        ,'' AS [TG118]
                                        ,'' AS [TG119]
                                        ,'' AS [TG120]
                                        ,0 AS [TG121]
                                        ,'' AS [TG122]
                                        ,'' AS [TG123]
                                        ,'' AS [TG124]
                                        ,'' AS [TG125]
                                        ,'' AS [TG126]
                                        ,'' AS [TG127]
                                        ,'' AS [TG128]
                                        ,'' AS [TG129]
                                        ,'' AS [TG130]
                                        ,'' AS [TG131]
                                        ,0 AS [TG132]
                                        ,'' AS [TG133]
                                        ,'' AS [TG134]
                                        ,1 AS [TG135]
                                        ,0 AS [TG136]
                                        ,0 AS [TG137]
                                        ,0 AS [TG138]
                                        ,0 AS [TG139]
                                        ,0 AS [TG140]
                                        ,0 AS [TG141]
                                        ,'' AS [TG142]
                                        ,'' AS [TG143]
                                        ,'' AS [TG144]
                                        ,'N' AS [TG145]
                                        ,'' AS [TG146]
                                        ,'' AS [TG147]
                                        ,'' AS [TG148]
                                        ,'' AS [TG149]
                                        ,'' AS [TG150]
                                        ,CONVERT(NVARCHAR,GETDATE(),112) AS [TG151]
                                        ,0 AS [TG152]
                                        ,0 AS [TG153]
                                        ,0 AS [TG154]
                                        ,0 AS [TG155]
                                        ,'N' AS [TG156]
                                        ,'' AS [TG157]
                                        ,'COPMI08' AS [TG158]
                                        ,'' AS [TG159]
                                        ,'' AS [TG160]
                                        ,'' AS [TG161]
                                        ,'' AS [TG162]
                                        ,'' AS [TG163]
                                        ,'' AS [UDF01]
                                        ,[主單編號] AS [UDF02]
                                        ,'' AS [UDF03]
                                        ,'' AS [UDF04]
                                        ,'' AS [UDF05]
                                        ,0 AS [UDF06]
                                        ,0AS [UDF07]
                                        ,0 AS [UDF08]
                                        ,0 AS [UDF09]
                                        ,0 AS [UDF10]
                                        FROM  [TKBUSINESS].[dbo].[TEMP91APPCOP],[TK].dbo.COPMA
                                        WHERE 1=1
                                        AND MA001='11127673'

                                        AND [購物車編號]='TG220719S01169'
                                        GROUP BY TEMP91APPCOP.購物車編號,TEMP91APPCOP.TG001,TEMP91APPCOP.TG002,TEMP91APPCOP.配送方式,TEMP91APPCOP.地址,TEMP91APPCOP.收件人,TEMP91APPCOP.收件人電話,TEMP91APPCOP.購物車總額,TEMP91APPCOP.通路商,TEMP91APPCOP.主單編號
                                        ,MA001,MA002,MA003,MA010,MA037,MA025


                                        --COPTH
                                        INSERT INTO  [test0923].[dbo].[COPTH]
                                        (
                                        [COMPANY]
                                        ,[CREATOR]
                                        ,[USR_GROUP]
                                        ,[CREATE_DATE]
                                        ,[MODIFIER]
                                        ,[MODI_DATE]
                                        ,[FLAG]
                                        ,[CREATE_TIME]
                                        ,[MODI_TIME]
                                        ,[TRANS_TYPE]
                                        ,[TRANS_NAME]
                                        ,[sync_date]
                                        ,[sync_time]
                                        ,[sync_mark]
                                        ,[sync_count]
                                        ,[DataUser]
                                        ,[DataGroup]
                                        ,[TH001]
                                        ,[TH002]
                                        ,[TH003]
                                        ,[TH004]
                                        ,[TH005]
                                        ,[TH006]
                                        ,[TH007]
                                        ,[TH008]
                                        ,[TH009]
                                        ,[TH010]
                                        ,[TH011]
                                        ,[TH012]
                                        ,[TH013]
                                        ,[TH014]
                                        ,[TH015]
                                        ,[TH016]
                                        ,[TH017]
                                        ,[TH018]
                                        ,[TH019]
                                        ,[TH020]
                                        ,[TH021]
                                        ,[TH022]
                                        ,[TH023]
                                        ,[TH024]
                                        ,[TH025]
                                        ,[TH026]
                                        ,[TH027]
                                        ,[TH028]
                                        ,[TH029]
                                        ,[TH030]
                                        ,[TH031]
                                        ,[TH032]
                                        ,[TH033]
                                        ,[TH034]
                                        ,[TH035]
                                        ,[TH036]
                                        ,[TH037]
                                        ,[TH038]
                                        ,[TH039]
                                        ,[TH040]
                                        ,[TH041]
                                        ,[TH042]
                                        ,[TH043]
                                        ,[TH044]
                                        ,[TH045]
                                        ,[TH046]
                                        ,[TH047]
                                        ,[TH048]
                                        ,[TH049]
                                        ,[TH050]
                                        ,[TH051]
                                        ,[TH052]
                                        ,[TH053]
                                        ,[TH054]
                                        ,[TH055]
                                        ,[TH056]
                                        ,[TH057]
                                        ,[TH058]
                                        ,[TH059]
                                        ,[TH060]
                                        ,[TH061]
                                        ,[TH062]
                                        ,[TH063]
                                        ,[TH064]
                                        ,[TH065]
                                        ,[TH066]
                                        ,[TH067]
                                        ,[TH068]
                                        ,[TH069]
                                        ,[TH070]
                                        ,[TH071]
                                        ,[TH072]
                                        ,[TH073]
                                        ,[TH074]
                                        ,[TH075]
                                        ,[TH076]
                                        ,[TH077]
                                        ,[TH078]
                                        ,[TH079]
                                        ,[TH080]
                                        ,[TH081]
                                        ,[TH082]
                                        ,[TH083]
                                        ,[TH084]
                                        ,[TH085]
                                        ,[TH086]
                                        ,[TH087]
                                        ,[TH088]
                                        ,[TH089]
                                        ,[TH090]
                                        ,[TH091]
                                        ,[TH092]
                                        ,[TH093]
                                        ,[TH094]
                                        ,[TH095]
                                        ,[TH096]
                                        ,[TH097]
                                        ,[TH098]
                                        ,[TH099]
                                        ,[TH100]
                                        ,[TH101]
                                        ,[TH102]
                                        ,[TH103]
                                        ,[TH104]
                                        ,[TH105]
                                        ,[TH106]
                                        ,[TH107]
                                        ,[TH108]
                                        ,[TH109]
                                        ,[TH110]
                                        ,[TH111]
                                        ,[TH112]
                                        ,[TH113]
                                        ,[TH114]
                                        ,[TH115]
                                        ,[TH116]
                                        ,[TH117]
                                        ,[TH118]
                                        ,[TH119]
                                        ,[TH120]
                                        ,[TH121]
                                        ,[TH122]
                                        ,[TH123]
                                        ,[TH124]
                                        ,[TH125]
                                        ,[TH126]
                                        ,[TH127]
                                        ,[UDF01]
                                        ,[UDF02]
                                        ,[UDF03]
                                        ,[UDF04]
                                        ,[UDF05]
                                        ,[UDF06]
                                        ,[UDF07]
                                        ,[UDF08]
                                        ,[UDF09]
                                        ,[UDF10]
                                        )
                                        SELECT 
                                         'TK' AS [COMPANY]
                                        ,'190083' AS [CREATOR]
                                        ,'102602' AS [USR_GROUP]
                                        ,CONVERT(NVARCHAR,GETDATE(),112) AS [CREATE_DATE]
                                        ,'' AS [MODIFIER]
                                        ,'' AS [MODI_DATE]
                                        ,1 AS [FLAG]
                                        ,CONVERT(NVARCHAR,GETDATE(),108) AS [CREATE_TIME]
                                        ,'' AS [MODI_TIME]
                                        ,'P001' AS [TRANS_TYPE]
                                        ,'COPMI08' AS [TRANS_NAME]
                                        ,'' AS [sync_date]
                                        ,'' AS [sync_time]
                                        ,'' AS [sync_mark]
                                        , 0 AS [sync_count]
                                        ,'' AS [DataUser]
                                        ,'102602' AS [DataGroup]
                                        ,[TG001] AS [TH001]
                                        ,[TG002] AS [TH002]
                                        ,[TH003] AS [TH003]
                                        ,Replace([商品料號],'''','') AS [TH004]
                                        ,MB002 AS [TH005]
                                        ,MB003 AS [TH006]
                                        ,'20001' AS [TH007]
                                        ,CONVERT(INT,[數量]) AS [TH008]
                                        ,MB004 AS [TH009]
                                        ,0 AS [TH010]
                                        ,'' AS [TH011]
                                        ,CONVERT(INT,[商品單價]) AS [TH012]
                                        ,CONVERT(INT,[商品單價])*CONVERT(INT,[數量]) AS [TH013]
                                        ,'' AS [TH014]
                                        ,'' AS [TH015]
                                        ,'' AS [TH016]
                                        ,'******' AS [TH017]
                                        ,'' AS [TH018]
                                        ,'' AS [TH019]
                                        ,'N' AS [TH020]
                                        ,'N' AS [TH021]
                                        ,'' AS [TH022]
                                        ,'' AS [TH023]
                                        ,0 AS [TH024]
                                        ,1 AS [TH025]
                                        ,'N' AS [TH026]
                                        ,'' AS [TH027]
                                        ,'' AS [TH028]
                                        ,'' AS [TH029]
                                        ,'' AS [TH030]
                                        ,'1' AS [TH031]
                                        ,'' AS [TH032]
                                        ,'' AS [TH033]
                                        ,'' AS [TH034]
                                        ,ROUND((CONVERT(INT,[商品單價])*CONVERT(INT,[數量]))/1.05,0) AS [TH035]
                                        ,((CONVERT(INT,[商品單價])*CONVERT(INT,[數量]))-(ROUND((CONVERT(INT,[商品單價])*CONVERT(INT,[數量]))/1.05,0))) AS [TH036]
                                        ,ROUND((CONVERT(INT,[商品單價])*CONVERT(INT,[數量]))/1.05,0) AS [TH037]
                                        ,((CONVERT(INT,[商品單價])*CONVERT(INT,[數量]))-(ROUND((CONVERT(INT,[商品單價])*CONVERT(INT,[數量]))/1.05,0))) AS [TH038]
                                        ,0 AS [TH039]
                                        ,0 AS [TH040]
                                        ,'' AS [TH041]
                                        ,'N' AS [TH042]
                                        ,0 AS [TH043]
                                        ,0 AS [TH044]
                                        ,0 AS [TH045]
                                        ,0 AS [TH046]
                                        ,0 AS [TH047]
                                        ,0 AS [TH048]
                                        ,0 AS [TH049]
                                        ,0 AS [TH050]
                                        ,'' AS [TH051]
                                        ,'' AS [TH052]
                                        ,'' AS [TH053]
                                        ,'' AS [TH054]
                                        ,0 AS [TH055]
                                        ,0 AS [TH056]
                                        ,0 AS [TH057]
                                        ,'' AS [TH058]
                                        ,'' AS [TH059]
                                        ,'' AS [TH060]
                                        ,0 AS [TH061]
                                        ,'' AS [TH062]
                                        ,0 AS [TH063]
                                        ,0 AS [TH064]
                                        ,'' AS [TH065]
                                        ,'' AS [TH066]
                                        ,'' AS [TH067]
                                        ,'N' AS [TH068]
                                        ,0 AS [TH069]
                                        ,'' AS [TH070]
                                        ,0 AS [TH071]
                                        ,0 AS [TH072]
                                        ,0 AS [TH073]
                                        ,[訂單編號] AS [TH074]
                                        ,'' AS [TH075]
                                        ,'' AS [TH076]
                                        ,'' AS [TH077]
                                        ,'' AS [TH078]
                                        ,'' AS [TH079]
                                        ,'' AS [TH080]
                                        ,'' AS [TH081]
                                        ,'' AS [TH082]
                                        ,0 AS [TH083]
                                        ,'' AS [TH084]
                                        ,0 AS [TH085]
                                        ,0 AS [TH086]
                                        ,'' AS [TH087]
                                        ,'' AS [TH088]
                                        ,'' AS [TH089]
                                        ,'N' AS [TH090]
                                        ,'N' AS [TH091]
                                        ,'' AS [TH092]
                                        ,'' AS [TH093]
                                        ,0 AS [TH094]
                                        ,'' AS [TH095]
                                        ,'' AS [TH096]
                                        ,'' AS [TH097]
                                        ,0 AS [TH098]
                                        ,'' AS [TH099]
                                        ,'' AS [TH100]
                                        ,'3' AS [TH101]
                                        ,'0' AS [TH102]
                                        ,'0' AS [TH103]
                                        ,'1' AS [TH104]
                                        ,'' AS [TH105]
                                        ,'' AS [TH106]
                                        ,'N' AS [TH107]
                                        ,'' AS [TH108]
                                        ,0 AS [TH109]
                                        ,'' AS [TH110]
                                        ,0 AS [TH111]
                                        ,'' AS [TH112]
                                        ,'' AS [TH113]
                                        ,0 AS [TH114]
                                        ,0 AS [TH115]
                                        ,0 AS [TH116]
                                        ,'' AS [TH117]
                                        ,'' AS [TH118]
                                        ,'' AS [TH119]
                                        ,'' AS [TH120]
                                        ,'' AS [TH121]
                                        ,'' AS [TH122]
                                        ,0 AS [TH123]
                                        ,0 AS [TH124]
                                        ,0 AS [TH125]
                                        ,0 AS [TH126]
                                        ,0 AS [TH127]
                                        ,'' AS [UDF01]
                                        ,'' AS [UDF02]
                                        ,'' AS [UDF03]
                                        ,'' AS [UDF04]
                                        ,'' AS [UDF05]
                                        ,0 AS [UDF06]
                                        ,0 AS [UDF07]
                                        ,0 AS [UDF08]
                                        ,0 AS [UDF09]
                                        ,0 AS [UDF10]
                                        FROM  [TK].dbo.COPMA,[TKBUSINESS].[dbo].[TEMP91APPCOP]
                                        LEFT JOIN [TK].dbo.INVMB ON [商品料號]=''''+MB001
                                        WHERE 1=1
                                        AND MA001='11127673'
                                        
                                        AND [購物車編號]='TG220719S01169'

                                        ");


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

        #endregion

        #region BUTTON


        private void button1_Click(object sender, EventArgs e)
        {
            Search(dateTimePicker1.Value.ToString("yyyyMM"));
        }
        private void button2_Click(object sender, EventArgs e)
        {
            //找出有運費的購眷，新增運費的品號資料=599010000000
            ADDSENDINVMB();
            //找出[折扣金額]<>0，新增到折扣品號=599030000004
            //總折扣金額-總點數折扣金額
            ADDDISCOUNTINVMB();
            //找出點數折扣金額<>0，新增到折扣品號=599030000003
            ADDDISCOUNTINVMB2();

            //找出未轉入銷貨單的TG020<>購物車編號，新增TG001、TG002
            UPDATETEMP91APPCOPCOPTG001TG002();
            //依TG001、TG002，新增TH003
            UPDATETEMP91APPCOPCOPTH003();

            //新增到ERP的COPTG、COPTH
            ADDERPCOPTGCOPTH();

            MessageBox.Show("完成");
        }
        #endregion


    }
}
