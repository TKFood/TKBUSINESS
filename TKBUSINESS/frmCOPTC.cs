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
    public partial class frmCOPTC : Form
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

        public frmCOPTC()
        {
            InitializeComponent();

        }
        private void frmCOPTC_Load(object sender, EventArgs e)
        {
            combobox1load();
            combobox1load2();
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

            String Sequel = "SELECT  [ID],[KINDS],[NAMES],[VALUE] FROM [TKBUSINESS].[dbo].[TBPARA] WHERE [KINDS]='frmCOPTC'";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("NAMES", typeof(string));    
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "NAMES";
            comboBox1.DisplayMember = "NAMES";
            sqlConn.Close();
        }
        public void combobox1load2()
        {

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            String Sequel = "SELECT  [ID],[KINDS],[NAMES],[VALUE] FROM [TKBUSINESS].[dbo].[TBPARA] WHERE [KINDS]='frmCOPTCTD016'";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("NAMES", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "NAMES";
            comboBox2.DisplayMember = "NAMES";
            sqlConn.Close();
        }

        public void Search()
        {
            try
            {
                sbSql.Clear();
                sbSql = SETsbSql();

                if (!string.IsNullOrEmpty(sbSql.ToString()))
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);


                    adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, tablename);
                    sqlConn.Close();


                    if (ds.Tables[tablename].Rows.Count == 0)
                    {
                        dataGridView1.DataSource = null;

                        textBox1.Text = null;
                        textBox2.Text = null;
                        textBox3.Text = null;
                    }
                    else
                    {
                        dataGridView1.DataSource = ds.Tables[tablename];
                        dataGridView1.AutoResizeColumns();
                        //rownum = ds.Tables[talbename].Rows.Count - 1;                       

                        //dataGridView1.CurrentCell = dataGridView1[0, 2];

                    }
                }
                else
                {

                }



            }
            catch
            {

            }
            finally
            {

            }
        }

        public StringBuilder SETsbSql()
        {
            StringBuilder STR = new StringBuilder();


          
            STR.AppendFormat(@"  
                                SELECT TC001 AS '單別',TC002 AS '單號',TC003 AS '日期',TC004 AS '客戶',TC053 AS '名稱',TC012 AS '客戶單號' ,TC042 AS '付款條件' 
                                FROM [TK].dbo.COPTC
                                WHERE TC001='{0}'
                                AND TC003>='{1}' AND TC003<='{2}'
                                ORDER BY TC001,TC002 
                                ", comboBox1.Text.ToString(), dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));

            tablename = "TEMPds1";

            return STR;
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    textBox2.Text = row.Cells["單別"].Value.ToString();
                    textBox3.Text = row.Cells["單號"].Value.ToString();
                    textBox1.Text = row.Cells["客戶單號"].Value.ToString();
                    textBox4.Text = row.Cells["付款條件"].Value.ToString();


                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;
                    textBox4.Text = null;
                }
            }
        }


        public void Search_DG2(string TC002)
        {
            try
            {
                sbSql.Clear();
                sbSql.AppendFormat(@"
                                    SELECT TC001 AS '單別',TC002 AS '單號',TC003 AS '日期',TC004 AS '客戶',TC053 AS '名稱'
                                    ,TD003 AS '序號'
                                    ,TD004 AS '品號'
                                    ,TD005 AS '品名'
                                    ,TD016 AS '結案碼'
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC002 LIKE '%{0}%'
                                    ORDER BY TC001,TC002,TD003 
                                    ", TC002);
                if (!string.IsNullOrEmpty(sbSql.ToString()))
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlDataAdapter da = new SqlDataAdapter(sbSql.ToString(), sqlConn))
                    {
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridView2.DataSource = dt;
                        dataGridView2.AutoResizeColumns();
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

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            textBox6.Text = null;
            textBox7.Text = null;
            textBox8.Text = null;

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    textBox6.Text = row.Cells["單別"].Value.ToString();
                    textBox7.Text = row.Cells["單號"].Value.ToString();
                    textBox8.Text = row.Cells["序號"].Value.ToString();
                    comboBox2.Text = row.Cells["結案碼"].Value.ToString();
                }
                else
                {
                    textBox6.Text = null;
                    textBox7.Text = null;
                    textBox8.Text = null;
                }
            }
        }

        public void SETSTATUS()
        {
            textBox1.ReadOnly = false;
            textBox4.ReadOnly = false;
        }

        public void SETSTATUS2()
        {
            textBox1.ReadOnly = true;
            textBox4.ReadOnly = true;
        }
        public void UPDATECOPTC()
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
                                    UPDATE [TK].dbo.COPTC
                                    SET TC012='{0}',TC042='{1}'
                                    WHERE TC001='{2}' AND TC002='{3}'
                                    ", textBox1.Text, textBox4.Text, textBox2.Text, textBox3.Text);

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


        public void UPDATECOPTD(string TD001,string TD002,string TD003,string TD016)
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

                using (SqlCommand cmd = new SqlCommand())
                {
                    sbSql.AppendFormat(@" 
                                    UPDATE [TK].dbo.COPTD
                                    SET TD016='{0}'
                                    WHERE TD001='{1}' AND TD002='{2}' AND TD003='{3}'
                                    ", TD016, TD001, TD002, TD003);
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
            Search();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SETSTATUS();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            UPDATECOPTC();
            SETSTATUS2();

            Search();
        }


        private void button4_Click(object sender, EventArgs e)
        {
            string TC005 = textBox5.Text.Trim();
            Search_DG2(TC005);
        }


        private void button5_Click(object sender, EventArgs e)
        {
            string TD001 = textBox6.Text.Trim();
            string TD002 = textBox7.Text.Trim();    
            string TD003 = textBox8.Text.Trim();
            string TD016 = comboBox2.Text;

            UPDATECOPTD(TD001, TD002, TD003, TD016);

            string TC005 = textBox5.Text.Trim();
            Search_DG2(TC005);
        }
        #endregion

    }
}
