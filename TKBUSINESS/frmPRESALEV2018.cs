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
        SqlConnection sqlConn2 = new SqlConnection();
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

        string DELID;


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

        public void SearchPRESALE2018V2()
        {
            try
            {
                sbSql.Clear();
                sbSql = SETsbSqlV2();

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
                        dataGridView3.DataSource = ds.Tables[tablename];
                        dataGridView3.AutoResizeColumns();
                        //rownum = ds.Tables[talbename].Rows.Count - 1;
                        dataGridView3.CurrentCell = dataGridView1.Rows[rownum].Cells[8];

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
            STR.AppendFormat(@"  ,[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[PRICES] AS '單價',[NUM] AS '數量',[TMONEY] AS '金額'");
            STR.AppendFormat(@"  ,[SALESID] AS '業務',[CUSTOMERID] AS '客戶'");
            STR.AppendFormat(@"  ,[ID]");
            STR.AppendFormat(@"  FROM [TKBUSINESS].[dbo].[PRESALE2018]");
            STR.AppendFormat(@"  WHERE [YEARS]={0} AND [MONTHS]>={1} AND [MONTHS]<={2}", numericUpDown1.Value.ToString(), numericUpDown2.Value.ToString(), numericUpDown3.Value.ToString());
            STR.AppendFormat(@"  {0}", STRQUERY.ToString());
            STR.AppendFormat(@"  ORDER BY  [YEARS],CONVERT(INT,[MONTHS]),[SALESID],[CUSTOMERID],[MB001]");
            STR.AppendFormat(@"  ");
            tablename = "TEMPds1";

            return STR;
        }

        public StringBuilder SETsbSqlV2()
        {
            StringBuilder STR = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            if (!string.IsNullOrEmpty(textBox41.Text.ToString()))
            {
                STRQUERY.AppendFormat(@" AND [CUSTOMERID] LIKE '{0}%'", textBox41.Text.ToString());
            }
            if (!string.IsNullOrEmpty(textBox40.Text.ToString()))
            {
                STRQUERY.AppendFormat(@" AND [SALESID] LIKE '{0}%'", textBox40.Text.ToString());
            }



            STR.AppendFormat(@"  SELECT");
            STR.AppendFormat(@"   [YEARS] AS '年',[MONTHS] AS '月',[SALESID] AS '業務',[SALESNAME] AS '業務名',[CUSTOMERID] AS '客戶',[CUSTOMERNAME] AS '客戶名' ");
            STR.AppendFormat(@"  ,[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[PRICES] AS '單價',[NUM] AS '數量',[TMONEY] AS '金額'");
            STR.AppendFormat(@"  ,[ID]");
            STR.AppendFormat(@"  FROM [TKBUSINESS].[dbo].[PRESALE2018]");
            STR.AppendFormat(@"  WHERE [YEARS]={0} AND [MONTHS]>={1} AND [MONTHS]<={2}", numericUpDown8.Value.ToString(), numericUpDown9.Value.ToString(), numericUpDown10.Value.ToString());
            STR.AppendFormat(@"  {0}", STRQUERY.ToString());
            STR.AppendFormat(@"  ORDER BY  [YEARS],CONVERT(INT,[MONTHS]),[SALESID],[CUSTOMERID],[MB001],SERNO");
            STR.AppendFormat(@"  ");

            tablename = "TEMPds2";

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
            textBox21.Text = null;

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
            Sequel.AppendFormat(@" SELECT MB001,MB002,MB003  FROM [TK].dbo.INVMB  WHERE MB001='{0}'", textBox11.Text.ToString());
            Sequel.AppendFormat(@"  ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);

            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MB001", typeof(string));
            dt.Columns.Add("MB002", typeof(string));
            dt.Columns.Add("MB003", typeof(string));
            da.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                textBox12.Text = dt.Rows[0]["MB002"].ToString();
                textBox21.Text = dt.Rows[0]["MB003"].ToString();
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
                sbSql.AppendFormat(" ([ID],[YEARS],[MONTHS],[SALESID],[SALESNAME],[CUSTOMERID],[CUSTOMERNAME],[MB001],[MB002],[PRICES],[NUM],[TMONEY],[MB003])");
                sbSql.AppendFormat(" VALUES({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}')", "NEWID()", textBox6.Text, numericUpDown4.Value.ToString(),textBox7.Text, textBox8.Text, textBox9.Text, textBox10.Text, textBox11.Text, textBox12.Text, textBox13.Text, textBox14.Text, textBox15.Text,  textBox21.Text);
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
                sbSql.AppendFormat(" SET [YEARS]='{0}',[MONTHS]='{1}',[SALESID]='{2}',[SALESNAME]='{3}',[CUSTOMERID]='{4}',[CUSTOMERNAME]='{5}',[MB001]='{6}',[MB002]='{7}',[PRICES]='{8}',[NUM]='{9}',[TMONEY]='{10}',[MB003]='{11}' ", textBox6.Text, numericUpDown4.Value.ToString(), textBox7.Text, textBox8.Text, textBox9.Text, textBox10.Text, textBox11.Text, textBox12.Text, textBox13.Text, textBox14.Text, textBox15.Text, textBox21.Text);
                sbSql.AppendFormat(" WHERE [ID]='{0}'",textBox20.Text);
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.CommandText = sbSql.ToString();
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
                    textBox21.Text = row.Cells["規格"].Value.ToString();

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
            FASTSQL.AppendFormat(@"  ,[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[PRICES] AS '單價',[NUM] AS '數量',[TMONEY] AS '金額'");
            FASTSQL.AppendFormat(@"  ,[SALESID] AS '業務',[CUSTOMERID] AS '客戶'");
            FASTSQL.AppendFormat(@"  ,[ID]");
            FASTSQL.AppendFormat(@"  FROM [TKBUSINESS].[dbo].[PRESALE2018]");
            FASTSQL.AppendFormat(@"  WHERE [YEARS]={0} AND [MONTHS]>={1} AND [MONTHS]<={2}", numericUpDown5.Value.ToString(), numericUpDown6.Value.ToString(), numericUpDown7.Value.ToString());
            FASTSQL.AppendFormat(@"  {0}", STRQUERY.ToString());
            FASTSQL.AppendFormat(@"  ORDER BY  [YEARS],CONVERT(INT,[MONTHS]),[SALESID],[CUSTOMERID],[MB001]");
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }


        public void INSERTTABLE(string MB001,string MB002)
        {
            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" INSERT INTO [TKBUSINESS].[dbo].[TEMP]");
                sbSql.AppendFormat(" ([ID],[MB001],[MB002])");
                sbSql.AppendFormat(" VALUES({0},'{1}','{2}')", "NEWID()", MB001, MB002);
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

        public void UPDATETABLE(string ID,string MB001,string MB002)
        {
            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" UPDATE [TKBUSINESS].[dbo].[TEMP]");
                sbSql.AppendFormat(" SET [MB001]='{0}',[MB002]='{1}'",MB001,MB002);
                sbSql.AppendFormat(" WHERE [ID]='{0}'", ID);
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

        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                MessageBox.Show(dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
            }
        }
        private void dataGridView2_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            
        }
        private void dataGridView2_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["SERNO"].Value = "99";
        }

        private void dataGridView2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if(e.ColumnIndex==2)
            {
                frmPRESALEV2018COPMA SUBfrmPRESALEV2018COPMA = new frmPRESALEV2018COPMA();
                SUBfrmPRESALEV2018COPMA.ShowDialog();
                dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = SUBfrmPRESALEV2018COPMA.TextBoxMsg;
            }
        }

        private void dataGridView3_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["年"].Value = numericUpDown8.Value.ToString();
            e.Row.Cells["單價"].Value = 0;
            e.Row.Cells["數量"].Value = 0;
            //e.Row.Cells["SERNO"].Value = "0";
        }

        private void dataGridView3_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 4)
                {
                    frmPRESALEV2018COPMA SUBfrmPRESALEV2018COPMA = new frmPRESALEV2018COPMA();
                    SUBfrmPRESALEV2018COPMA.ShowDialog();
                    dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = SUBfrmPRESALEV2018COPMA.TextBoxMsg.Trim();
                }
                if (e.ColumnIndex == 6)
                {
                    frmPRESALEV2018INVMB SUBfrmPRESALEV2018INVMB = new frmPRESALEV2018INVMB();
                    SUBfrmPRESALEV2018INVMB.ShowDialog();
                    dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = SUBfrmPRESALEV2018INVMB.TextBoxMsg.Trim();
                }
            }
            catch
            {

            }
          
        }
        private void dataGridView3_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //MessageBox.Show("aa");
           

            //
           
           

            
            
        }

        private void dataGridView1_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void dataGridView3_Validated(object sender, EventArgs e)
        {
           
        }
        private void dataGridView3_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            if (!string.IsNullOrEmpty(dataGridView3.Rows[e.RowIndex].Cells[2].Value.ToString().Trim()))
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                StringBuilder Sequel = new StringBuilder();
                Sequel.AppendFormat(@" SELECT [Code],[CnName] FROM [HRMDB].[dbo].[Employee]   WHERE [Code]='{0}'", dataGridView3.Rows[e.RowIndex].Cells[2].Value.ToString().Trim());
                Sequel.AppendFormat(@"  ");
                SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);

                DataTable dt = new DataTable();
                sqlConn.Open();

                dt.Columns.Add("Code", typeof(string));
                dt.Columns.Add("CnName", typeof(string));
                da.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    dataGridView3.Rows[e.RowIndex].Cells[3].Value = dt.Rows[0]["CnName"].ToString().Trim();
                }
                else
                {
                    dataGridView3.Rows[e.RowIndex].Cells[3].Value = null;
                }

                sqlConn.Close();
            }
            if (!string.IsNullOrEmpty(dataGridView3.Rows[e.RowIndex].Cells[4].Value.ToString().Trim()))
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                StringBuilder Sequel = new StringBuilder();
                Sequel.AppendFormat(@" SELECT MA001,MA002  FROM [TK].dbo.COPMA   WHERE MA001='{0}'", dataGridView3.Rows[e.RowIndex].Cells[4].Value.ToString().Trim());
                Sequel.AppendFormat(@"  ");
                SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);

                DataTable dt = new DataTable();
                sqlConn.Open();

                dt.Columns.Add("MA001", typeof(string));
                dt.Columns.Add("MA002", typeof(string));
                da.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    dataGridView3.Rows[e.RowIndex].Cells[5].Value = dt.Rows[0]["MA002"].ToString().Trim();
                }
                else
                {
                    dataGridView3.Rows[e.RowIndex].Cells[5].Value = null;
                }

                sqlConn.Close();
            }

            if (!string.IsNullOrEmpty(dataGridView3.Rows[e.RowIndex].Cells[6].Value.ToString().Trim()))
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn2 = new SqlConnection(connectionString);
                StringBuilder Sequel2 = new StringBuilder();

                Sequel2.AppendFormat(@" SELECT MB001,MB002,MB003  FROM [TK].dbo.INVMB  WHERE MB001='{0}'", dataGridView3.Rows[e.RowIndex].Cells[6].Value.ToString().Trim());
                Sequel2.AppendFormat(@"  ");
                SqlDataAdapter da2 = new SqlDataAdapter(Sequel2.ToString(), sqlConn);

                DataTable dt2 = new DataTable();
                sqlConn2.Open();

                dt2.Columns.Add("MB001", typeof(string));
                dt2.Columns.Add("MB002", typeof(string));
                dt2.Columns.Add("MB003", typeof(string));
                da2.Fill(dt2);

                if (dt2.Rows.Count > 0)
                {
                    dataGridView3.Rows[e.RowIndex].Cells[7].Value = dt2.Rows[0]["MB002"].ToString().Trim();
                    dataGridView3.Rows[e.RowIndex].Cells[8].Value = dt2.Rows[0]["MB003"].ToString().Trim();
                }
                else
                {
                    dataGridView3.Rows[e.RowIndex].Cells[7].Value = null;
                    dataGridView3.Rows[e.RowIndex].Cells[8].Value = null;
                }

                sqlConn.Close();
            }

            if (!string.IsNullOrEmpty(dataGridView3.Rows[e.RowIndex].Cells[9].Value.ToString().Trim()) && !string.IsNullOrEmpty(dataGridView3.Rows[e.RowIndex].Cells[10].Value.ToString().Trim()))
            {
                dataGridView3.Rows[e.RowIndex].Cells[11].Value = Convert.ToDecimal(dataGridView3.Rows[e.RowIndex].Cells[9].Value.ToString()) * Convert.ToDecimal(dataGridView3.Rows[e.RowIndex].Cells[10].Value.ToString());
            }
        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.CurrentRow != null)
            {

                int rowindex = dataGridView3.CurrentRow.Index;
                DataGridViewRow row = dataGridView3.Rows[rowindex];

                if (rowindex >= 0)
                {
                    DELID = row.Cells["ID"].Value.ToString();
                }
                else
                {
                    DELID = null;

                }
            }
            else
            {
                DELID = null;

            }
                  
        }

        public void INSERTPRESALE2018(string YEARS, string MONTHS, string SALESID, string SALESNAME, string CUSTOMERID, string CUSTOMERNAME, string MB001, string MB002, string MB003, string PRICES, string NUM, string TMONEY)
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
                sbSql.AppendFormat(" ([ID],[YEARS],[MONTHS],[SALESID],[SALESNAME],[CUSTOMERID],[CUSTOMERNAME],[MB001],[MB002],[MB003],[PRICES],[NUM],[TMONEY])");
                sbSql.AppendFormat(" VALUES({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',{10},{11},{12})", "NEWID()", YEARS, MONTHS, SALESID, SALESNAME, CUSTOMERID, CUSTOMERNAME, MB001, MB002, MB003, PRICES, NUM, TMONEY);
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


        public void UPDATEPRESALE2018(string ID, string YEARS, string MONTHS, string SALESID, string SALESNAME, string CUSTOMERID, string CUSTOMERNAME, string MB001, string MB002, string MB003, string PRICES, string NUM, string TMONEY)
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
                sbSql.AppendFormat(" SET [YEARS]='{0}',[MONTHS]='{1}',[SALESID]='{2}',[SALESNAME]='{3}',[CUSTOMERID]='{4}',[CUSTOMERNAME]='{5}',[MB001]='{6}',[MB002]='{7}',[MB003]='{8}',[PRICES]={9},[NUM]={10},[TMONEY]={11} ", YEARS, MONTHS, SALESID, SALESNAME, CUSTOMERID, CUSTOMERNAME, MB001, MB002, MB003, PRICES, NUM, TMONEY);
                sbSql.AppendFormat(" WHERE [ID]='{0}'", ID);
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

        public void DELETEPRESALE2018()
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
                sbSql.AppendFormat(" WHERE [ID]='{0}'", DELID);
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

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection sqlConn2 = new SqlConnection();
                SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
                SqlDataAdapter adapter2 = new SqlDataAdapter();
                DataSet ds2=new DataSet();
                
                string sbSql2 = "SELECT [ID],[SERNO],[MB001],[MB002] FROM [TKBUSINESS].[dbo].[TEMP] ORDER BY [SERNO]";

                if (!string.IsNullOrEmpty(sbSql2.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn2 = new SqlConnection(connectionString);
                    adapter2 = new SqlDataAdapter(sbSql2.ToString(), sqlConn2);
                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);

                    sqlConn2.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "TEMP");
                    sqlConn2.Close();


                    if (ds2.Tables["TEMP"].Rows.Count == 0)
                    {
                        dataGridView2.DataSource = null;
                    }
                    else
                    {
                        dataGridView2.DataSource = ds2.Tables["TEMP"];
                        dataGridView2.AutoResizeColumns();
                        dataGridView2.Columns["ID"].ReadOnly = true;
                        dataGridView2.Columns["SERNO"].ReadOnly = true;


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

        private void button10_Click(object sender, EventArgs e)
        {
            dataGridView2.EndEdit();

            int rows = 1;

            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if(dataGridView2.Rows.Count-1 >= rows)
                {
                    if (string.IsNullOrEmpty(row.Cells["ID"].Value.ToString()))
                    {
                        INSERTTABLE(row.Cells["MB001"].Value.ToString(), row.Cells["MB002"].Value.ToString());
                    }
                    else if (!string.IsNullOrEmpty(row.Cells["ID"].Value.ToString()))
                    {
                        UPDATETABLE(row.Cells["ID"].Value.ToString(), row.Cells["MB001"].Value.ToString(), row.Cells["MB002"].Value.ToString());
                    }

                    rows++;
                }
              
            }
            MessageBox.Show("Records");

            button11.PerformClick();
        }
        private void button12_Click(object sender, EventArgs e)
        {
            SearchPRESALE2018V2();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            dataGridView3.ReadOnly = false;
        }
        private void button14_Click(object sender, EventArgs e)
        {
            dataGridView3.EndEdit();
            dataGridView3.ReadOnly = true;           

            int rows = 1;

            foreach (DataGridViewRow row in dataGridView3.Rows)
            {
                if (dataGridView3.Rows.Count - 1 >= rows)
                {
                    if (string.IsNullOrEmpty(row.Cells["ID"].Value.ToString()))
                    {
                        INSERTPRESALE2018(row.Cells["年"].Value.ToString(), row.Cells["月"].Value.ToString(), row.Cells["業務"].Value.ToString(), row.Cells["業務名"].Value.ToString(), row.Cells["客戶"].Value.ToString(), row.Cells["客戶名"].Value.ToString(), row.Cells["品號"].Value.ToString(), row.Cells["品名"].Value.ToString(), row.Cells["規格"].Value.ToString(), row.Cells["單價"].Value.ToString(), row.Cells["數量"].Value.ToString(), row.Cells["金額"].Value.ToString());
                    }
                    else if (!string.IsNullOrEmpty(row.Cells["ID"].Value.ToString()))
                    {
                        UPDATEPRESALE2018(row.Cells["ID"].Value.ToString(),row.Cells["年"].Value.ToString(), row.Cells["月"].Value.ToString(), row.Cells["業務"].Value.ToString(), row.Cells["業務名"].Value.ToString(), row.Cells["客戶"].Value.ToString(), row.Cells["客戶名"].Value.ToString(), row.Cells["品號"].Value.ToString(), row.Cells["品名"].Value.ToString(), row.Cells["規格"].Value.ToString(), row.Cells["單價"].Value.ToString(), row.Cells["數量"].Value.ToString(), row.Cells["金額"].Value.ToString());
                    }

                    rows++;
                }

            }
            MessageBox.Show("存檔完成");

            button12.PerformClick();
        }
        private void button15_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETEPRESALE2018();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
            button12.PerformClick();
        }


        private void button16_Click(object sender, EventArgs e)
        {
            frmPRESALEV2018COPY SUBfrmPRESALEV2018COPY = new frmPRESALEV2018COPY();
            SUBfrmPRESALEV2018COPY.ShowDialog();

            SearchPRESALE2018V2();
            MessageBox.Show("完成");
        }

        private void button17_Click(object sender, EventArgs e)
        {
            frmPRESALEV2018COPYCOPMA SUBfrmPRESALEV2018COPYCOPMA = new frmPRESALEV2018COPYCOPMA();
            SUBfrmPRESALEV2018COPYCOPMA.ShowDialog();

            SearchPRESALE2018V2();
            MessageBox.Show("完成");

        }

        #endregion


    }
}
