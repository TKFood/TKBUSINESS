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
                                    ,[TH001]
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

        #endregion

        #region BUTTON


        private void button1_Click(object sender, EventArgs e)
        {
            Search(dateTimePicker1.Value.ToString("yyyyMM"));
        }

        #endregion

      
    }
}
