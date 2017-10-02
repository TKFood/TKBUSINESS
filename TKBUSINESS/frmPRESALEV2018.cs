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

namespace TKBUSINESS
{
    public partial class frmPRESALEV2018 : Form
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
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;
        string SALSESID = null;
        int result;
        DataGridViewRow drPRESLAES = new DataGridViewRow();


        public frmPRESALEV2018()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SearchPRESALE2018()
        {
            try
            {
                sbSql.Clear();
                sbSql = SETsbSql();

                if (!string.IsNullOrEmpty(sbSql.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);
                    adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, tablename);
                    sqlConn.Close();

                    
                    if (ds.Tables[tablename].Rows.Count == 0)
                    {
                        dataGridView1.DataSource = null;
                    }
                    else
                    {
                        dataGridView1.DataSource = ds.Tables[tablename];
                        dataGridView1.AutoResizeColumns();
                        //rownum = ds.Tables[talbename].Rows.Count - 1;
                        dataGridView1.CurrentCell = dataGridView1.Rows[rownum].Cells[8];

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
            StringBuilder STRQUERY = new StringBuilder();

            if (!string.IsNullOrEmpty(textBox1.Text.ToString()))
            {
                STRQUERY.AppendFormat(@" AND [CUSTOMERID] LIKE '{0}%'", textBox1.Text.ToString());
            }
            if (!string.IsNullOrEmpty(textBox3.Text.ToString()))
            {
                STRQUERY.AppendFormat(@" AND [SALESID] LIKE '{0}%'", textBox3.Text.ToString());
            }


            STR.AppendFormat(@"  SELECT");
            STR.AppendFormat(@"  [YEARS] AS '年',[MONTHS] AS '月',[SALESID] AS '業務',[SALESNAME] AS '業務名',[CUSTOMERID] AS '客戶',[CUSTOMERNAME] AS '客戶名'");
            STR.AppendFormat(@"  ,[MB001] AS '品號',[MB002] AS '品名',[PRICES] AS '單價',[NUM] AS '數量',[TMONEY] AS '金額'");
            STR.AppendFormat(@"  ,[ID]");
            STR.AppendFormat(@"  FROM [TKBUSINESS].[dbo].[PRESALE2018]");
            STR.AppendFormat(@"  WHERE [YEARS]='{0}' AND [MONTHS]>='{1}' AND [MONTHS]<='{2}'",numericUpDown1.Value.ToString(), numericUpDown2.Value.ToString(), numericUpDown3.Value.ToString());
            STR.AppendFormat(@"  {0}", STRQUERY.ToString());
            STR.AppendFormat(@"  ORDER BY  [YEARS],[MONTHS],[CUSTOMERID],[MB001]");
            STR.AppendFormat(@"  ");

            tablename = "TEMPds1";

            return STR;
        }

        public void SETNULL()
        {
            textBox7.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
            textBox11.Text = null;
            textBox12.Text = null;
            textBox13.Text = null;
            textBox14.Text = null;
            textBox15.Text = null;
            textBox20.Text = null;


        }

        public void SETREADONLY(string TYPE)
        {
            if(TYPE.Equals("1"))
            {
                textBox7.ReadOnly = true;
                textBox8.ReadOnly = true;
                textBox9.ReadOnly = true;
                textBox10.ReadOnly = true;
                textBox11.ReadOnly = true;
                textBox12.ReadOnly = true;
                textBox13.ReadOnly = true;
                textBox14.ReadOnly = true;
                
            }
            else
            {
                textBox7.ReadOnly = false;
                textBox8.ReadOnly = false;
                textBox9.ReadOnly = false;
                textBox10.ReadOnly = false;
                textBox11.ReadOnly = false;
                textBox12.ReadOnly = false;
                textBox13.ReadOnly = false;
                textBox14.ReadOnly = false;
            }
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            CALMONEY();
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            CALMONEY();
        }

        public void CALMONEY()
        {
            if (!string.IsNullOrEmpty(textBox13.Text.ToString()) && !string.IsNullOrEmpty(textBox14.Text.ToString()))
            {
                if ((Convert.ToDouble(textBox13.Text.ToString()) > 0) && (Convert.ToDouble(textBox14.Text.ToString()) > 0))
                {
                    textBox15.Text = (Convert.ToDouble(textBox13.Text.ToString()) * Convert.ToDouble(textBox14.Text.ToString())).ToString();
                }
            }
            else
            {
                textBox15.Text = null;
            }

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            SEARCHEMP();
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            SEARCHCOPMA();
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            SEARCHPRODUCTNAME();
        }
        public void SEARCHEMP()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@" SELECT [Code],[CnName] FROM [HRMDB].[dbo].[Employee]   WHERE [Code]='{0}'", textBox7.Text.ToString());
            Sequel.AppendFormat(@"  ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);

            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("Code", typeof(string));
            dt.Columns.Add("CnName", typeof(string));
            da.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                textBox8.Text = dt.Rows[0]["CnName"].ToString();
            }
            else
            {
                textBox8.Text = null;
            }

            sqlConn.Close();
        }

        public void SEARCHCOPMA()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@" SELECT MA001,MA002  FROM [TK].dbo.COPMA   WHERE MA001='{0}'", textBox9.Text.ToString());
            Sequel.AppendFormat(@"  ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);

            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MA001", typeof(string));
            dt.Columns.Add("MA002", typeof(string));
            da.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                textBox10.Text = dt.Rows[0]["MA002"].ToString();
            }
            else
            {
                textBox10.Text = null;
            }

            sqlConn.Close();
        }

        public void SEARCHPRODUCTNAME()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@" SELECT MB001,MB002  FROM [TK].dbo.INVMB  WHERE MB001='{0}'", textBox11.Text.ToString());
            Sequel.AppendFormat(@"  ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);

            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MB001", typeof(string));
            dt.Columns.Add("MB002", typeof(string));
            da.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                textBox12.Text = dt.Rows[0]["MB002"].ToString();
            }
            else
            {
                textBox12.Text = null;
            }

            sqlConn.Close();
        }

        public void SAVEPRESALE2018()
        {
            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" INSERT INTO  [TKBUSINESS].[dbo].[PRESALE2018]");
                sbSql.AppendFormat(" ([ID],[YEARS],[MONTHS],[SALESID],[SALESNAME],[CUSTOMERID],[CUSTOMERNAME],[MB001],[MB002],[PRICES],[NUM],[TMONEY])");
                sbSql.AppendFormat(" VALUES({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}')", "NEWID()", textBox6.Text, numericUpDown4.Value.ToString(),textBox7.Text, textBox8.Text, textBox9.Text, textBox10.Text, textBox11.Text, textBox12.Text, textBox13.Text, textBox14.Text, textBox15.Text, textBox20.Text);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

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
            SearchPRESALE2018();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SETNULL();
            SETREADONLY("0");
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SAVEPRESALE2018();
            SearchPRESALE2018();
            SETREADONLY("1");
        }
        private void button4_Click(object sender, EventArgs e)
        {
            SETREADONLY("0");
        }








        #endregion

       
    }
}
