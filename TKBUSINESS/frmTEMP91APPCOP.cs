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
        DataTable EXCEL = null;

        public frmTEMP91APPCOP()
        {
            InitializeComponent();

            SETTEXTBOX();
            SETDATES();

            comboBox1load();

        }


        #region FUNCTION
        public void comboBox1load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT [ID],[NAMES] FROM [TKBUSINESS].[dbo].[TEMP91APPCOPBASES] ORDER BY [ID] ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAMES", typeof(string));

            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "ID";
            comboBox1.DisplayMember = "NAMES";
            sqlConn.Close();


        }
        public void SETTEXTBOX()
        {
            textBox2.Text = "11127673";
            textBox1.Text = SERACHCOPMA(textBox2.Text.ToString().Trim());
        }

        public void SETDATES()
        {
            dateTimePicker2.Value = DateTime.Now;
            dateTimePicker3.Value = DateTime.Now;

            if (DateTime.Now.DayOfWeek== DayOfWeek.Friday)
            {
                dateTimePicker2.Value = DateTime.Now.AddDays(3);
                dateTimePicker3.Value = DateTime.Now.AddDays(3);
            }
            else if  (DateTime.Now.DayOfWeek == DayOfWeek.Saturday)
            {
                dateTimePicker2.Value = DateTime.Now.AddDays(2);
                dateTimePicker3.Value = DateTime.Now.AddDays(2);
            }         
            else
            {
                dateTimePicker2.Value = DateTime.Now.AddDays(1);
                dateTimePicker3.Value = DateTime.Now.AddDays(1);
            }
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


        public void Search(string YYYYMM,string STATUS)
        {
            DataSet ds = new DataSet();           
            StringBuilder sbSqlQUERY = new StringBuilder();
            YYYYMM = YYYYMM.Substring(YYYYMM.Length - 4, 4);

            try
            {
                sbSql.Clear();
                sbSqlQUERY.Clear();

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                if(STATUS.Equals("未出貨"))
                {
                    sbSqlQUERY.AppendFormat(@" 
                                            AND [TG001]+[TG002] NOT IN (SELECT [TG001]+[TG002] FROM [TK].dbo.COPTG)

                                            ");
                }
                else if (STATUS.Equals("已出貨"))
                {
                    sbSqlQUERY.AppendFormat(@" 
                                            AND [TG001]+[TG002]  IN (SELECT [TG001]+[TG002] FROM [TK].dbo.COPTG)

                                            ");
                }
                else if (STATUS.Equals("全部"))
                {
                    sbSqlQUERY.AppendFormat(@" 
                                        
                                            ");
                }

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
                                    {1}

                                    ORDER BY [購物車編號],[主單編號],[訂單編號]
                                         ", "TG"+ YYYYMM, sbSqlQUERY.ToString());

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

        public void Search2(string YYYYMMDD)
        {           

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
                                    WHERE [TG002] LIKE '{0}%'
                                    ORDER BY [購物車編號],[主單編號],[訂單編號]
                                         ", YYYYMMDD);

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

        public void UPDATETEMP91APPCOPCOPTG001TG002(string TG003)
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
                                     AND [購物車編號]  NOT IN (SELECT SUBSTRING(TG020,1,14) FROM [TK].dbo.COPTG WHERE ISNULL(TG020,'')<>'')
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
                    string TG002 = GETMAXTG002(TG001, TG003);

                    int serno = Convert.ToInt16(TG002.Substring(8, 3));
                    serno = serno - 1;

                    foreach (DataRow DR in ds.Tables["ds"].Rows)
                    {
                        string 購物車編號 = DR["購物車編號"].ToString();

                        //流水號+1
                        serno = serno + 1;
                        string temp = serno.ToString();
                        temp = temp.PadLeft(3, '0');
                        TG002 = TG003 + temp.ToString();

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
                        TC002 = SETTG002(ds4.Tables["TEMPds4"].Rows[0]["TG002"].ToString(), TG003);
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

        public string SETTG002(string TG002,string TG003)
        {
            if (TG002.Equals("00000000000"))
            {
                return TG003 + "001";
            }

            else
            {
                int serno = Convert.ToInt16(TG002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return TG003 + temp.ToString();
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
                                    ,'599010000000' [商品料號]
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
                                    , [轉單日期時間]
                                    ,'' [預計出貨日期]
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
                                    ,'599030000004' [商品料號]
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
                                    ,'' [交期]
                                    ,[線上訂單建立類型]
                                    FROM [TKBUSINESS].[dbo].[TEMP91APPCOP]
                                    WHERE [折扣金額]<>0 


                                    GROUP  BY [購物車編號]
                                    ,[主單編號]
                                    ,[轉單日期時間]
                             
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
                                   
                                    ,[線上訂單建立類型]

                                    HAVING (SUM(CONVERT(INT,[折扣金額]))-SUM(CONVERT(INT,[點數折扣金額])))<>0
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
                                    ,'' [預計出貨日期]
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
                                    ,'599030000003' [商品料號]
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
                                    ,'' [交期]
                                    ,[線上訂單建立類型]
                                    FROM [TKBUSINESS].[dbo].[TEMP91APPCOP]
                                    WHERE [折扣金額]<>0 

                                  
                                    GROUP  BY [購物車編號]
                                    ,[主單編號]
                                    ,[轉單日期時間]
                                   
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
                                   
                                    ,[線上訂單建立類型]

                                    HAVING (SUM(CONVERT(INT,[點數折扣金額]))*-1)<>0

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

        public void ADDERPCOPTGCOPTH(string TG003)
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

                //AND[購物車編號] = 'TG220719S01169'
                //COPTG --AND [購物車編號]  NOT IN (SELECT TG020 FROM [TK].dbo.COPTG WHERE ISNULL(TG020,'')<>'')
                //COPTH --AND [訂單編號]  NOT IN (SELECT TH074 FROM [TK].dbo.COPTH WHERE ISNULL(TH074,'')<>'')

                //INSERT INTO [test0923].[dbo].[COPTG]
                sbSql.AppendFormat(@" 
                                   
                                   
                                        --COPTG
                                        INSERT INTO [TK].[dbo].[COPTG]
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
                                        ,'{0}' AS [TG003]
                                        ,MA001 AS [TG004]
                                        ,MA015 AS [TG005]
                                        ,MA016 AS [TG006]
                                        ,MA003 AS [TG007]
                                        ,REPLACE(
                                            REPLACE(
                                            (CASE WHEN [配送方式]='宅配' THEN [地址]+','+[收件人]+','+[收件人電話]
                                            WHEN [配送方式]='貨到付款' THEN [地址]+','+[收件人]+','+[收件人電話]
                                            WHEN [配送方式]='超商取貨付款' AND [通路商]='7-11(統一)' THEN '新北市樹林區佳園路2段70-1號(大智通),EC驗收組,02-26738186#101'
                                            WHEN [配送方式]='付款後超商取貨' AND [通路商]='7-11(統一)' THEN '新北市樹林區佳園路2段70-1號(大智通),EC驗收組,02-26738186#101'
                                            WHEN [配送方式]='超商取貨付款' AND [通路商]='全家' THEN '桃園市大溪區仁善里15鄰新光東路76巷22-2號(日翊文化),電子商務部,03-3075581'
                                            WHEN [配送方式]='付款後超商取貨' AND [通路商]='全家' THEN '桃園市大溪區仁善里15鄰新光東路76巷22-2號(日翊文化),電子商務部,03-3075581'
                                            WHEN [配送方式]='超商取貨付款' AND [通路商]='萊爾富' THEN '桃園市龍潭區中原路二段545巷146號(萊爾富),物流中心EC收件組,03-2865168#611'
                                            WHEN [配送方式]='付款後超商取貨' AND [通路商]='萊爾富' THEN '桃園市龍潭區中原路二段545巷146號(萊爾富),物流中心EC收件組,03-2865168#611'
                                            END
                                            )
                                            ,'，',',')
                                            ,'Taiwan,','') AS [TG008]
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
                                        ,MA016 AS [TG026]
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
                                        ,LEFT('{0}',6) AS [TG038]
                                        ,'' AS [TG039]
                                        ,'' AS [TG040]
                                        ,0 AS [TG041]
                                        ,'{0}' AS [TG042]
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
                                        ,CONVERT(NVARCHAR(30),[通路商]+'-'+[配送方式]+'-'+[收件人]) AS [TG075]
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
                                        ,REPLACE([指定到貨日期],'/','') AS [TG110]
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
                                        ,'{0}' AS [TG151]
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
                                        ,(SELECT [主單編號] FROM [TKBUSINESS].[dbo].[TEMP91APPCOP] TEMP WHERE TEMP.購物車編號=[TEMP91APPCOP].購物車編號 GROUP BY [主單編號]  FOR XML PATH('')) AS [UDF02]
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
                                        AND [購物車編號]  NOT IN (SELECT SUBSTRING(TG020,1,14) FROM [TK].dbo.COPTG WHERE ISNULL(TG020,'')<>'')
                                        AND [TEMP91APPCOP].TG002 LIKE '{0}%'
                                        GROUP BY TEMP91APPCOP.購物車編號,TEMP91APPCOP.TG001,TEMP91APPCOP.TG002,TEMP91APPCOP.配送方式,TEMP91APPCOP.地址,TEMP91APPCOP.收件人,TEMP91APPCOP.收件人電話,TEMP91APPCOP.購物車總額,TEMP91APPCOP.通路商,TEMP91APPCOP.指定到貨日期
                                        ,MA001,MA002,MA003,MA010,MA037,MA025,MA015,MA016


                                        --COPTH
                                        INSERT INTO  [TK].[dbo].[COPTH]
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
                                        ,CASE WHEN (CONVERT(INT,[商品單價])*CONVERT(INT,[數量]))<0 THEN [折價券活動名稱] ELSE '' END AS [TH018]
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
                                        LEFT JOIN [TK].dbo.INVMB ON LTRIM(RTRIM(Replace([商品料號],'''','')))=MB001                                        
                                        WHERE 1=1

                                        AND MA001='11127673'
                                        AND ISNULL(MB002,'')<>''
                                        AND [訂單編號]  NOT IN (SELECT TH074 FROM [TK].dbo.COPTH WHERE ISNULL(TH074,'')<>'')
                                        AND TG002 LIKE '{0}%'

                                        ", TG003);


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

        public void CHECKADDDATA()
        {
            //IEnumerable<DataRow> tempExcept = null;

            DataTable DT1 = SEARCHTEMP91APPCOP();
            DataTable DT2 = IMPORTEXCEL();

            //找DataTable差集
            //要有相同的欄位名稱
            //找DataTable差集
            //如果兩個datatable中有部分欄位相同，可以使用Contains比較　　
            var tempExcept = from r in DT2.AsEnumerable()
                             where
                             !(from rr in DT1.AsEnumerable() select rr.Field<string>("訂單編號")).Contains(
                             r.Field<string>("訂單編號"))
                             select r;


            //var tempExcept = DT2.AsEnumerable();

            if (tempExcept.Count() > 0)
            {
                //差集集合
                DataTable dt3 = tempExcept.CopyToDataTable();

                INSERTINTOTEMP91APPCOP(dt3);
            }
        }

        public DataTable SEARCHTEMP91APPCOP()
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();
          
            //THISYEARS = "21";

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();



                //核準過TASK_RESULT='0'
                //AND DOC_NBR  LIKE 'QC1002{0}%'

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [訂單編號]
                                    FROM [TKBUSINESS].[dbo].[TEMP91APPCOP]

                                    ");


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

        public DataTable IMPORTEXCEL()
        {
            //記錄選到的檔案路徑
            _path = null;

            OpenFileDialog od = new OpenFileDialog();
            od.Filter = "Excell|*.xls;*.xlsx;";

            DialogResult dr = od.ShowDialog();
            if (dr == DialogResult.Abort)
            {
                return null;
            }
            if (dr == DialogResult.Cancel)
            {
                return null;
            }
           

            textBox3.Text = od.FileName.ToString();
            _path = od.FileName.ToString();

            try
            {
                //  ExcelConn(_path);
                //找出不同excel的格式，設定連接字串
                //xls跟非xls
                string constr = null;
                string CHECKEXCELFORMAT = _path.Substring(_path.Length - 4, 4);

                if (CHECKEXCELFORMAT.CompareTo(".xls") == 0)
                {
                    constr = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _path + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
                }
                else
                {
                    constr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _path + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
                }

                //找出excel的第1張分頁名稱，用query中                
                OleDbConnection Econ = new OleDbConnection(constr);
                Econ.Open();



                DataTable excelShema = Econ.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string firstSheetName = excelShema.Rows[0]["TABLE_NAME"].ToString();

                string Query = string.Format("Select * FROM [{0}]", firstSheetName);
                OleDbCommand Ecom = new OleDbCommand(Query, Econ);


                DataTable dtExcelData = new DataTable();

                OleDbDataAdapter oda = new OleDbDataAdapter(Query, Econ);
                Econ.Close();
                oda.Fill(dtExcelData);
                DataTable Exceldt = dtExcelData;

                //把第一列的欄位名移除
                //Exceldt.Rows[0].Delete();

                if(Exceldt.Rows.Count>0)
                {
                    return Exceldt;
                }
                else
                {
                    return null;
                }
                

            }
            catch (Exception ex)
            {
                return null;
                //MessageBox.Show(string.Format("錯誤:{0}", ex.Message), "Not Imported", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
        }

        public void INSERTINTOTEMP91APPCOP(DataTable DT )
        {
            string 購物車編號 = null;
            string 主單編號 = null;
            string 訂單編號 = null;
            string 轉單日期時間 = null;
            string 預計出貨日期 = null;
            string 配送方式 = null;
            string 通路商 = null;
            string 溫層類別 = null;
            string 收件人 = null;
            string 收件人電話 = null;
            string 地址 = null;
            string 門市 = null;
            string 訂單來源 = null;
            string 商品名稱 = null;
            string 商品選項 = null;
            string 商品料號 = null;
            string 數量 = null;
            string 商品單價 = null;
            string 運費 = null;
            string 配送編號 = null;
            string 狀態日期 = null;
            string 出貨單狀態 = null;
            string 訂單狀態 = null;
            string 活動代碼 = null;
            string 活動名稱 = null;
            string 折扣金額 = null;
            string 銷售金額折扣後 = null;
            string 付款方式 = null;
            string 活動折扣金額 = null;
            string 折價券活動序號 = null;
            string 折價券活動名稱 = null;
            string 折價券折扣金額 = null;
            string 貨到物流中心日 = null;
            string 建議貨到期限 = null;
            string 會員編號 = null;
            string 商店備註 = null;
            string 訂購備註 = null;
            string 貨運單備註 = null;
            string 驗退原因說明 = null;
            string 訂單確認日期 = null;
            string 實體會員編號 = null;
            string 商品屬性 = null;
            string 商品贈品關聯代碼 = null;
            string 購買人 = null;
            string 購買人會員等級 = null;
            string 活動對象 = null;
            string 活動會員等級 = null;
            string 總成本 = null;
            string 是否為加價購品 = null;
            string 國碼 = null;
            string 收件國家 = null;
            string 取消原因 = null;
            string 購物車總額 = null;
            string 商品頁序號 = null;
            string 點數活動名稱= null;
            string 折抵點數 = null;
            string 點數折扣金額 = null;
            string 已設定為不可退貨商品 = null;
            string 郵遞區號 = null;
            string 指定到貨日期 = null;
            string 指定到貨時段 = null;
            string 贈品券活動序號 = null;
            string 國家地區運費活動名稱 = null;
            string 運費折扣 = null;
            string 地區州省份= null;
            string 城市 = null;
            string 鄉鎮市區 = null;
            string 街道 = null;
            string 實際出貨數量 = null;
            string 實際出貨金額 = null;
            string 配送商 = null;
            string TS重量小計g = null;
            string 運費券活動序號 = null;
            string 自訂活動代碼 = null;
            string 交期 = null;
            string 線上訂單建立類型 = null;
            string TG001 = null;
            string TG002 = null;
            string TH003 = null;
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
                
                foreach (DataRow DR in DT.Rows)
                {
                   if(!string.IsNullOrEmpty(DR["訂單編號"].ToString()))
                    {
                        //
                        try
                        {
                            購物車編號 = DR["購物車編號"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            主單編號 = DR["主單編號"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            訂單編號 = DR["訂單編號"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {

                            轉單日期時間 = DR["轉單日期時間"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            string dateTimeString = DR["預計出貨日期"].ToString().Replace("'", "").Replace("上午", "").Replace("下午", "");
                            DateTime SET_dateTime = Convert.ToDateTime(dateTimeString);
                            SET_dateTime = SET_dateTime.AddDays(1);

                            if (SET_dateTime.DayOfWeek == DayOfWeek.Saturday)
                            {
                                SET_dateTime = SET_dateTime.AddDays(2);
                            }
                            else if (SET_dateTime.DayOfWeek == DayOfWeek.Sunday)
                            {
                                SET_dateTime = SET_dateTime.AddDays(1);
                            }


                            //預計出貨日期 = DR["預計出貨日期"].ToString().Replace("'", "").Replace(" ", "");
                            預計出貨日期 = SET_dateTime.ToString("yyyy/MM/dd");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            配送方式 = DR["配送方式"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            通路商 = DR["通路商"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            溫層類別 = DR["溫層類別"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            收件人 = DR["收件人"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            收件人電話 = DR["收件人電話"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            地址 = DR["地址"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            門市 = DR["門市"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            訂單來源 = DR["訂單來源"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            商品名稱 = DR["商品名稱"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            商品選項 = DR["商品選項"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            商品料號 = DR["商品料號"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            數量 = DR["數量"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            商品單價 = DR["商品單價"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            運費 = DR["運費"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            配送編號 = DR["配送編號"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            狀態日期 = DR["狀態日期"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            出貨單狀態 = DR["出貨單狀態"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            訂單狀態 = DR["訂單狀態"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            活動代碼 = DR["活動代碼"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            活動名稱 = DR["活動名稱"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            折扣金額 = DR["折扣金額"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            銷售金額折扣後 = DR["銷售金額(折扣後)"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            付款方式 = DR["付款方式"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            活動折扣金額 = DR["活動折扣金額"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            折價券活動序號 = DR["折價券活動序號"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            折價券活動名稱 = DR["折價券活動名稱"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            折價券折扣金額 = DR["折價券折扣金額"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            貨到物流中心日 = DR["貨到物流中心日"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            建議貨到期限 = DR["建議貨到期限"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            會員編號 = DR["會員編號"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            商店備註 = DR["商店備註"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            訂購備註 = DR["訂購備註"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            貨運單備註 = DR["貨運單備註"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            驗退原因說明 = DR["驗退原因說明"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            訂單確認日期 = DR["訂單確認日期"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            實體會員編號 = DR["實體會員編號"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            商品屬性 = DR["商品屬性"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            商品贈品關聯代碼 = DR["商品贈品關聯代碼"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            購買人 = DR["購買人"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            購買人會員等級 = DR["購買人會員等級"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            活動對象 = DR["活動對象"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            活動會員等級 = DR["活動會員等級"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            總成本 = DR["總成本"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            是否為加價購品 = DR["是否為加價購品"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            國碼 = DR["國碼"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            收件國家 = DR["收件國家"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            取消原因 = DR["取消原因"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            購物車總額 = DR["購物車總額"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            商品頁序號 = DR["商品頁序號"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            點數活動名稱 = DR["點數活動名稱"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            折抵點數 = DR["折抵點數"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            點數折扣金額 = DR["點數折扣金額"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            已設定為不可退貨商品 = DR["已設定為不可退貨商品"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            郵遞區號 = DR["郵遞區號"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            if(!string.IsNullOrEmpty(DR["指定到貨日期"].ToString()))
                            {
                                string dateTimeString = DR["指定到貨日期"].ToString().Replace("'", "").Replace("上午", "").Replace("下午", "");
                                DateTime SET_dateTime = Convert.ToDateTime(dateTimeString);

                                //SET_dateTime = SET_dateTime.AddDays(1);

                                //if (SET_dateTime.DayOfWeek == DayOfWeek.Saturday)
                                //{
                                //    SET_dateTime = SET_dateTime.AddDays(2);
                                //}
                                //else if (SET_dateTime.DayOfWeek == DayOfWeek.Sunday)
                                //{
                                //    SET_dateTime = SET_dateTime.AddDays(1);
                                //}

                                //指定到貨日期 = DR["指定到貨日期"].ToString().Replace("'", "").Replace(" ", "");
                                指定到貨日期 = SET_dateTime.ToString("yyyy/MM/dd");
                            }
                            else
                            {
                                指定到貨日期 = "";
                            }

                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            指定到貨時段 = DR["指定到貨時段"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            贈品券活動序號 = DR["贈品券活動序號"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            國家地區運費活動名稱 = DR["國家地區運費活動名稱"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            運費折扣 = DR["運費折扣"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            地區州省份 = DR["地區/州/省份"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            城市 = DR["城市"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            鄉鎮市區 = DR["鄉鎮市區"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            街道 = DR["街道"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            實際出貨數量 = DR["實際出貨數量"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            實際出貨金額 = DR["實際出貨金額"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            配送商 = DR["配送商"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            TS重量小計g = DR["TS重量小計(g)"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            運費券活動序號 = DR["運費券活動序號"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            自訂活動代碼 = DR["自訂活動代碼"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            交期 = DR["交期"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            線上訂單建立類型 = DR["線上訂單建立類型"].ToString().Replace("'", "").Replace(" ", "");
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            TG001 = "";
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            TG002 = "";
                        }
                        catch
                        {

                        }
                        //
                        try
                        {
                            TH003 = "";
                        }
                        catch
                        {

                        }




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
                                        ,[TG001]
                                        ,[TG002]
                                        ,[TH003]
                                        )
                                        VALUES
                                        (
                                        '{0}'
                                        ,'{1}'
                                        ,'{2}'
                                        ,'{3}'
                                        ,'{4}'
                                        ,'{5}'
                                        ,'{6}'
                                        ,'{7}'
                                        ,'{8}'
                                        ,'{9}'
                                        ,'{10}'
                                        ,'{11}'
                                        ,'{12}'
                                        ,'{13}'
                                        ,'{14}'
                                        ,'{15}'
                                        ,'{16}'
                                        ,'{17}'
                                        ,'{18}'
                                        ,'{19}'
                                        ,'{20}'
                                        ,'{21}'
                                        ,'{22}'
                                        ,'{23}'
                                        ,'{24}'
                                        ,'{25}'
                                        ,'{26}'
                                        ,'{27}'
                                        ,'{28}'
                                        ,'{29}'
                                        ,'{30}'
                                        ,'{31}'
                                        ,'{32}'
                                        ,'{33}'
                                        ,'{34}'
                                        ,'{35}'
                                        ,'{36}'
                                        ,'{37}'
                                        ,'{38}'
                                        ,'{39}'
                                        ,'{40}'
                                        ,'{41}'
                                        ,'{42}'
                                        ,'{43}'
                                        ,'{44}'
                                        ,'{45}'
                                        ,'{46}'
                                        ,'{47}'
                                        ,'{48}'
                                        ,'{49}'
                                        ,'{50}'
                                        ,'{51}'
                                        ,'{52}'
                                        ,'{53}'
                                        ,'{54}'
                                        ,'{55}'
                                        ,'{56}'
                                        ,'{57}'
                                        ,'{58}'
                                        ,'{59}'
                                        ,'{60}'
                                        ,'{61}'
                                        ,'{62}'
                                        ,'{63}'
                                        ,'{64}'
                                        ,'{65}'
                                        ,'{66}'
                                        ,'{67}'
                                        ,'{68}'
                                        ,'{69}'
                                        ,'{70}'
                                        ,'{71}'
                                        ,'{72}'
                                        ,'{73}'
                                        ,'{74}'
                                        ,'{75}'
                                        ,'{76}'
                                        ,'{77}'
                                        ,'{78}'

                                        )
                                           
                                         ", 購物車編號
                                                , 主單編號
                                                , 訂單編號
                                                , 轉單日期時間
                                                , 預計出貨日期
                                                , 配送方式
                                                , 通路商
                                                , 溫層類別
                                                , 收件人
                                                , 收件人電話
                                                , 地址
                                                , 門市
                                                , 訂單來源
                                                , 商品名稱
                                                , 商品選項
                                                , 商品料號
                                                , 數量
                                                , 商品單價
                                                , 運費
                                                , 配送編號
                                                , 狀態日期
                                                , 出貨單狀態
                                                , 訂單狀態
                                                , 活動代碼
                                                , 活動名稱
                                                , 折扣金額
                                                , 銷售金額折扣後
                                                , 付款方式
                                                , 活動折扣金額
                                                , 折價券活動序號
                                                , 折價券活動名稱
                                                , 折價券折扣金額
                                                , 貨到物流中心日
                                                , 建議貨到期限
                                                , 會員編號
                                                , 商店備註
                                                , 訂購備註
                                                , 貨運單備註
                                                , 驗退原因說明
                                                , 訂單確認日期
                                                , 實體會員編號
                                                , 商品屬性
                                                , 商品贈品關聯代碼
                                                , 購買人
                                                , 購買人會員等級
                                                , 活動對象
                                                , 活動會員等級
                                                , 總成本
                                                , 是否為加價購品
                                                , 國碼
                                                , 收件國家
                                                , 取消原因
                                                , 購物車總額
                                                , 商品頁序號
                                                , 點數活動名稱
                                                , 折抵點數
                                                , 點數折扣金額
                                                , 已設定為不可退貨商品
                                                , 郵遞區號
                                                , 指定到貨日期
                                                , 指定到貨時段
                                                , 贈品券活動序號
                                                , 國家地區運費活動名稱
                                                , 運費折扣
                                                , 地區州省份
                                                , 城市
                                                , 鄉鎮市區
                                                , 街道
                                                , 實際出貨數量
                                                , 實際出貨金額
                                                , 配送商
                                                , TS重量小計g
                                                , 運費券活動序號
                                                , 自訂活動代碼
                                                , 交期
                                                , 線上訂單建立類型
                                                , TG001
                                                , TG002
                                                , TH003

                                            );
                    }

                }



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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox6.Text = "";

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox6.Text = row.Cells["購物車編號"].Value.ToString();
                }
                else
                {
                    

                }
            }
        }

        public void DELETETEMP91APPCOP(string ID)
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
                                    DELETE [TKBUSINESS].[dbo].[TEMP91APPCOP]
                                    WHERE [購物車編號]='{0}'
                                       
                                        ", ID);



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

        public void CHECKID()
        {

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
                                    SELECT [購物車編號],COUNT([購物車編號]) AS NUM
                                    FROM(
                                    SELECT DISTINCT
                                    [購物車編號]
                                    ,[主單編號]
                                    FROM [TKBUSINESS].[dbo].[TEMP91APPCOP]
                                    GROUP BY [購物車編號],[主單編號]
                                    ) AS TEMP
                                    GROUP BY [購物車編號]
                                    HAVING COUNT([購物車編號])>=2
                                    ");

                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    

                }
                else
                {
                    StringBuilder ID = new StringBuilder();

                    ID.AppendFormat(@"以下是購物車編號跟主單編號有重複 ");
                    foreach(DataRow DR in ds.Tables["ds"].Rows)
                    {
                        ID.AppendFormat(@" 
                                        購物車編號='{0}'
                                        ", DR["購物車編號"].ToString());

                        ID.AppendLine();
                    }

                    MessageBox.Show(ID.ToString());
                    

                }



            }
            catch
            {

            }
            finally
            {

            }
        }

        #endregion

        #region BUTTON


        private void button1_Click(object sender, EventArgs e)
        {
            Search(dateTimePicker1.Value.ToString("yyyyMM"), comboBox1.Text.ToString());

            MessageBox.Show("完成");

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
            UPDATETEMP91APPCOPCOPTG001TG002(dateTimePicker2.Value.ToString("yyyyMMdd"));
            //依TG001、TG002，新增TH003
            UPDATETEMP91APPCOPCOPTH003();

           


            Search(dateTimePicker1.Value.ToString("yyyyMM"), comboBox1.Text.ToString());

            MessageBox.Show("完成");
        }
        private void button4_Click(object sender, EventArgs e)
        {
            CHECKADDDATA();

            Search(dateTimePicker1.Value.ToString("yyyyMM"), comboBox1.Text.ToString());
            MessageBox.Show("完成");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETETEMP91APPCOP(textBox6.Text.ToString());

                Search(dateTimePicker1.Value.ToString("yyyyMM"),comboBox1.Text.ToString());
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
          
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Search2(dateTimePicker3.Value.ToString("yyyyMMdd"));
        }
        private void button6_Click(object sender, EventArgs e)
        {
            //新增到ERP的COPTG、COPTH
            ADDERPCOPTGCOPTH(dateTimePicker2.Value.ToString("yyyyMMdd"));
        }

        private void button7_Click(object sender, EventArgs e)
        {
            CHECKID();
        }

        #endregion


    }
}
