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
    public partial class frmCOPMD : Form
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
   
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;
        int result;

        string MD001 = "";

        public frmCOPMD()
        {
            InitializeComponent();
        }

        private void frmCOPMD_Load(object sender, EventArgs e)
        {
            Search(textBox1.Text.Trim());
        }
        #region FUNCTION
        public void Search(string MB001)
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
                                    SELECT 
                                    MA001 AS '客代'
                                    ,MA002  AS '客戶'
                                    FROM [TK].dbo.COPMA
                                    WHERE (MA001 LIKE '%{0}%' OR MA002 LIKE '%{0}%')
                                         ", MB001);

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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            MD001 = "";

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    MD001 = row.Cells["客代"].Value.ToString();
                }
                else
                {


                }
            }
        }

        public void Search_COPMD(string MD001,string MD003)
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
                                   SELECT          
                                    MD002 AS '地址代號'
                                    ,MD003 AS '地址一'
                                    ,MD004 AS '地址二'
                                    ,MD005 AS '備註'
                                    ,MD006 AS '全名'
                                    ,MD007 AS '連絡人'
                                    ,MD008 AS '統一編號'
                                    ,MD009 AS 'TEL_NO'
                                    ,MD010 AS 'FAX_NO'
                                    ,MD011 AS '收貨部門'
                                    ,MD012 AS '收貨人'
                                    ,MD001 AS '客戶代號'
                                    FROM              [TK].dbo.COPMD
                                    WHERE          (MD001 = '{0}')
                                    AND (MD002 LIKE '%{1}%' OR MD003 LIKE '%{1}%' OR MD006 LIKE '%{1}%'  OR MD012 LIKE '%{1}%')
                                    ORDER BY MD002
                                         ", MD001, MD003);

                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;

                }
                else
                {
                    dataGridView2.DataSource = ds.Tables["ds"];
                    dataGridView2.AutoResizeColumns();
                    dataGridView2.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                    dataGridView2.DefaultCellStyle.Font = new Font("Tahoma", 10);
                    //dataGridView1.Columns["序號"].Width = 30;
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
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    textBox3.Text = row.Cells["地址代號"].Value.ToString().Trim();
                    textBox4.Text = row.Cells["地址一"].Value.ToString().Trim();
                    textBox5.Text = row.Cells["地址二"].Value.ToString().Trim();
                    textBox6.Text = row.Cells["備註"].Value.ToString().Trim();
                    textBox7.Text = row.Cells["全名"].Value.ToString().Trim();
                    textBox8.Text = row.Cells["連絡人"].Value.ToString().Trim();
                    textBox9.Text = row.Cells["統一編號"].Value.ToString().Trim();
                    textBox10.Text = row.Cells["TEL_NO"].Value.ToString().Trim();
                    textBox11.Text = row.Cells["FAX_NO"].Value.ToString().Trim();
                    textBox12.Text = row.Cells["收貨部門"].Value.ToString().Trim();
                    textBox13.Text = row.Cells["收貨人"].Value.ToString().Trim();
                    textBox14.Text = row.Cells["客戶代號"].Value.ToString().Trim();

                    textBox26.Text = row.Cells["客戶代號"].Value.ToString().Trim();
                }
                else
                {
                    textBox3.Text = "";
                    textBox4.Text = "";
                    textBox5.Text = "";
                    textBox6.Text = "";
                    textBox7.Text = "";
                    textBox8.Text = "";
                    textBox9.Text = "";
                    textBox10.Text = "";
                    textBox11.Text = "";
                    textBox12.Text = "";
                    textBox13.Text = "";
                    textBox14.Text = "";

                    textBox26.Text = "";

                }
            }
        }

        public void DELETE_COPMD(string MD001,string MD002)
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
                                    DELETE [TK].dbo.COPMD
                                    WHERE MD001='{0}' AND MD002='{1}'
                                       
                                        ", MD001,  MD002);



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

        public void UPDATE_COPMD(string MD001
                                , string MD002
                                , string MD003
                                , string MD004
                                , string MD005
                                , string MD006
                                , string MD007
                                , string MD008
                                , string MD009
                                , string MD010
                                , string MD011
                                , string MD012
                                )
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
                                    UPDATE [TK].dbo.COPMD
                                    SET MD003='{2}'
                                    ,MD004='{3}'
                                    ,MD005='{4}'
                                    ,MD006='{5}'
                                    ,MD007='{6}'
                                    ,MD008 ='{7}'
                                    ,MD009 ='{8}'
                                    ,MD010='{9}'
                                    ,MD011='{10}'
                                    ,MD012='{11}'
                                    WHERE MD001='{0}' AND MD002='{1}'
                                       
                                        ", MD001, MD002, MD003, MD004, MD005, MD006, MD007, MD008, MD009, MD010, MD011, MD012);



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

        public void ADD_COPMD(
                                 string MD001
                                , string MD002
                                , string MD003
                                , string MD004
                                , string MD005
                                , string MD006
                                , string MD007
                                , string MD008
                                , string MD009
                                , string MD010
                                , string MD011
                                , string MD012
                                )
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
                                   INSERT INTO [TK].dbo.COPMD
                                    (
                                    MD001
                                    ,MD002
                                    ,MD003
                                    ,MD004
                                    ,MD005
                                    ,MD006 
                                    ,MD007
                                    ,MD008 
                                    ,MD009 
                                    ,MD010
                                    ,MD011
                                    ,MD012
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
                                    )
                                       
                                        ", MD001, MD002, MD003, MD004, MD005, MD006, MD007, MD008, MD009, MD010, MD011, MD012);



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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search(textBox1.Text.Trim());
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Search_COPMD(MD001, textBox2.Text);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            UPDATE_COPMD(textBox14.Text.Trim()
                                , textBox3.Text.Trim()
                                , textBox4.Text.Trim()
                                , textBox5.Text.Trim()
                                , textBox6.Text.Trim()
                                , textBox7.Text.Trim()
                                , textBox8.Text.Trim()
                                , textBox9.Text.Trim()
                                , textBox10.Text.Trim()
                                , textBox11.Text.Trim()
                                , textBox12.Text.Trim()
                                , textBox13.Text.Trim()
                                );

            Search_COPMD(MD001, textBox2.Text);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETE_COPMD(textBox14.Text.Trim(), textBox3.Text.Trim());

                Search_COPMD(MD001, textBox2.Text);

            }
            else if (dialogResult == DialogResult.No)
            {

            }


             
        }
        private void button5_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox15.Text))
            {
                ADD_COPMD(
                    textBox26.Text.Trim()
                  , textBox15.Text.Trim()
                  , textBox16.Text.Trim()
                  , textBox17.Text.Trim()
                  , textBox18.Text.Trim()
                  , textBox19.Text.Trim()
                  , textBox20.Text.Trim()
                  , textBox21.Text.Trim()
                  , textBox22.Text.Trim()
                  , textBox23.Text.Trim()
                  , textBox24.Text.Trim()
                  , textBox25.Text.Trim()
                  );

                Search_COPMD(MD001, textBox2.Text);
            }
            else
            {
                MessageBox.Show("地址代號 未填不能新增");
            }
           
        }

        #endregion

      
    }
}
