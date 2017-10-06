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
            STR.AppendFormat(@"  [YEARS] AS '年',[MONTHS] AS '月',[SALESNAME] AS '業務名',[CUSTOMERNAME] AS '客戶名'");
            STR.AppendFormat(@"  ,[MB001] AS '品號',[MB002] AS '品名',[PRICES] AS '單價',[NUM] AS '數量',[TMONEY] AS '金額'");
            STR.AppendFormat(@"  ,[SALESID] AS '業務',[CUSTOMERID] AS '客戶'");
            STR.AppendFormat(@"  ,[ID]");
            STR.AppendFormat(@"  FROM [TKBUSINESS].[dbo].[PRESALE2018]");
            STR.AppendFormat(@"  WHERE [YEARS]={0} AND [MONTHS]>={1} AND [MONTHS]<={2}",numericUpDown1.Value.ToString(), numericUpDown2.Value.ToString(), numericUpDown3.Value.ToString());
            STR.AppendFormat(@"  {0}", STRQUERY.ToString());
            STR.AppendFormat(@"  ORDER BY  [YEARS],CONVERT(INT,[MONTHS]),[SALESID],[CUSTOMERID],[MB001]");
            STR.AppendFormat(@"  ");

            tablename = "TEMPds1";

            return STR;
        }

        public void SETNULL()
        {
            textBox7.Text = null;
            textBox8.Text = null;
            //textBox9.Text = null;
            //textBox10.Text = null;
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
                //textBox8.ReadOnly = true;
                textBox9.ReadOnly = true;
                //textBox10.ReadOnly = true;
                textBox11.ReadOnly = true;
                //textBox12.ReadOnly = true;
                textBox13.ReadOnly = true;
                textBox14.ReadOnly = true;
                
            }
            else
            {
                textBox7.ReadOnly = false;
                //textBox8.ReadOnly = false;
                textBox9.ReadOnly = false;
                //textBox10.ReadOnly = false;
                textBox11.ReadOnly = false;
                //textBox12.ReadOnly = false;
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
        public void UPDATESALE2018()
        {
            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" UPDATE [TKBUSINESS].[dbo].[PRESALE2018]");
                sbSql.AppendFormat(" SET [YEARS]='{0}',[MONTHS]='{1}',[SALESID]='{2}',[SALESNAME]='{3}',[CUSTOMERID]='{4}',[CUSTOMERNAME]='{5}',[MB001]='{6}',[MB002]='{7}',[PRICES]='{8}',[NUM]='{9}',[TMONEY]='{10}' ", textBox6.Text, numericUpDown4.Value.ToString(), textBox7.Text, textBox8.Text, textBox9.Text, textBox10.Text, textBox11.Text, textBox12.Text, textBox13.Text, textBox14.Text, textBox15.Text);
                sbSql.AppendFormat(" WHERE [ID]='{0}'",textBox20.Text);
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

        public void DELPRESALE2018()
        {
            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" DELETE [TKBUSINESS].[dbo].[PRESALE2018]");
                sbSql.AppendFormat(" WHERE [ID]='{0}'", textBox20.Text);
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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;



                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox6.Text = row.Cells["年"].Value.ToString();
                    textBox7.Text = row.Cells["業務"].Value.ToString();
                    textBox8.Text = row.Cells["業務名"].Value.ToString();
                    textBox9.Text = row.Cells["客戶"].Value.ToString();
                    textBox10.Text = row.Cells["客戶名"].Value.ToString();
                    textBox11.Text = row.Cells["品號"].Value.ToString();
                    textBox12.Text = row.Cells["品名"].Value.ToString();
                    textBox13.Text = row.Cells["單價"].Value.ToString();
                    textBox14.Text = row.Cells["數量"].Value.ToString();
                    textBox15.Text = row.Cells["金額"].Value.ToString();
                    textBox20.Text = row.Cells["ID"].Value.ToString();

                    numericUpDown4.Value = Convert.ToInt32(row.Cells["月"].Value.ToString());

                   

                }
                else
                {
                    SETNULL();

                }
            }
        }

        public void SETFASTREPORT()
        {
            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\銷售預估.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL();
            Table.SelectCommand = SQL;
            report1.Preview = previewControl1;
            report1.Show();


        }

        public string SETFASETSQL()
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            if (!string.IsNullOrEmpty(textBox201.Text.ToString()))
            {
                STRQUERY.AppendFormat(@" AND [CUSTOMERID] LIKE '{0}%'", textBox201.Text.ToString());
            }
            if (!string.IsNullOrEmpty(textBox200.Text.ToString()))
            {
                STRQUERY.AppendFormat(@" AND [SALESID] LIKE '{0}%'", textBox200.Text.ToString());
            }

            FASTSQL.AppendFormat(@"  SELECT");
            FASTSQL.AppendFormat(@"  [YEARS] AS '年',[MONTHS] AS '月',[SALESNAME] AS '業務名',[CUSTOMERNAME] AS '客戶名'");
            FASTSQL.AppendFormat(@"  ,[MB001] AS '品號',[MB002] AS '品名',[PRICES] AS '單價',[NUM] AS '數量',[TMONEY] AS '金額'");
            FASTSQL.AppendFormat(@"  ,[SALESID] AS '業務',[CUSTOMERID] AS '客戶'");
            FASTSQL.AppendFormat(@"  ,[ID]");
            FASTSQL.AppendFormat(@"  FROM [TKBUSINESS].[dbo].[PRESALE2018]");
            FASTSQL.AppendFormat(@"  WHERE [YEARS]={0} AND [MONTHS]>={1} AND [MONTHS]<={2}", numericUpDown5.Value.ToString(), numericUpDown6.Value.ToString(), numericUpDown7.Value.ToString());
            FASTSQL.AppendFormat(@"  {0}", STRQUERY.ToString());
            FASTSQL.AppendFormat(@"  ORDER BY  [YEARS],CONVERT(INT,[MONTHS]),[SALESID],[CUSTOMERID],[MB001]");
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
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
            if(string.IsNullOrEmpty(textBox20.Text))
            {
                SAVEPRESALE2018();
                SearchPRESALE2018();
                SETREADONLY("1");
            }
            else
            {
                UPDATESALE2018();
                SearchPRESALE2018();
                SETREADONLY("1");
            }
           
        }
        private void button4_Click(object sender, EventArgs e)
        {
            rownum = dataGridView1.CurrentCell.RowIndex;
            SETREADONLY("0");
        }




        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELPRESALE2018();
                SearchPRESALE2018();               
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }



        private void button8_Click(object sender, EventArgs e)
        {
            frmPRESALEV2018COPMA SUBfrmPRESALEV2018COPMA = new frmPRESALEV2018COPMA();
            SUBfrmPRESALEV2018COPMA.ShowDialog();
            textBox9.Text = SUBfrmPRESALEV2018COPMA.TextBoxMsg;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            frmPRESALEV2018INVMB SUBfrmPRESALEV2018INVMB = new frmPRESALEV2018INVMB();
            SUBfrmPRESALEV2018INVMB.ShowDialog();
            textBox11.Text = SUBfrmPRESALEV2018INVMB.TextBoxMsg;

        }

        private void button9_Click(object sender, EventArgs e)
        {
            frmPRESALEV2018COPY SUBfrmPRESALEV2018COPY = new frmPRESALEV2018COPY();
            SUBfrmPRESALEV2018COPY.ShowDialog();

            SearchPRESALE2018();
            MessageBox.Show("完成");

        }


        #endregion


    }
}
