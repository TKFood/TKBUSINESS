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

        #endregion




        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SearchPRESALE2018();
        }
        private void button2_Click(object sender, EventArgs e)
        {

        }
        private void button3_Click(object sender, EventArgs e)
        {

        }
        private void button4_Click(object sender, EventArgs e)
        {

        }

        #endregion






    }
}
