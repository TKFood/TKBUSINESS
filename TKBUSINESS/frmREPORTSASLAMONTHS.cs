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
using FastReport;
using FastReport.Data;
using TKITDLL;

namespace TKBUSINESS
{
    public partial class frmREPORTSASLAMONTHS : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        Report report1 = new Report();

        string tablename = null;
        int rownum = 0;

        public frmREPORTSASLAMONTHS()
        {
            InitializeComponent();

            SETDATE();
        }

        #region FUNCTION
        public void SETDATE()
        {
            dateTimePicker1.Value = Convert.ToDateTime(DateTime.Now.Year + "/01/01");
            dateTimePicker2.Value = DateTime.Now.AddMonths(0).AddDays(-DateTime.Now.AddMonths(1).Day);
        }
        public void Search_INVMB_DG1(string MB001)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            try
            {
                sbSql.Clear();
                sbSql.AppendFormat(@"
                                    SELECT RTRIM(LTRIM(MB001)) AS '品號',RTRIM(LTRIM(MB002))  AS '品名'
                                    FROM [TK].dbo.INVMB
                                    WHERE (MB001 LIKE '4%' OR MB001 LIKE '5%' )
                                    AND (MB001 LIKE '%{0}%' OR MB002 LIKE '%{0}%')
                                    ORDER BY MB001
                                    ", MB001);


                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                    dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                    dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 10);
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
            textBox2.Text = null;
            textBox3.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    textBox2.Text = row.Cells["品號"].Value.ToString().Trim();
                    textBox3.Text = row.Cells["品名"].Value.ToString().Trim();
                }
            }
        }

        public void Search_CMSME_DG2(string ME001)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            try
            {
                sbSql.Clear();
                sbSql.AppendFormat(@"
                                    SELECT 
                                    RTRIM(LTRIM(ME001)) AS '部門代號'
                                    ,RTRIM(LTRIM(ME002)) AS '部門'
                                    FROM [TK].dbo.CMSME
                                    WHERE( ME001 LIKE '%{0}%' OR ME002 LIKE '%{0}%')
                                    AND ME002 NOT LIKE '%停用%'
                                    ORDER BY ME001
                                    ", ME001);


                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

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

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];

                    textBox6.Text = row.Cells["部門代號"].Value.ToString().Trim();
                    textBox7.Text = row.Cells["部門"].Value.ToString().Trim();
                }
            }
        }

        public void SETFASTREPORT_SASLA(string SDATES, string EDATES, string LA005, string LA007)
        {
            string P1 = SDATES;
            string P2 = EDATES;

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);


            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            Report report1 = new Report();

            SQL1 = SETSQL_SETFASTREPORT_SASLA(SDATES, EDATES, LA005, LA007);
            report1.Load(@"REPORT\銷貨月報v4.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();




            report1.SetParameterValue("P1", P1);
            report1.SetParameterValue("P2", P2);

            report1.Preview = previewControl4;
            report1.Show();
        }

        public StringBuilder SETSQL_SETFASTREPORT_SASLA(string SDATES, string EDATES, string LA005, string LA007)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SBQUERY = new StringBuilder();
            StringBuilder SBQUERY2 = new StringBuilder();

            if (!string.IsNullOrEmpty(LA007)&&!LA007.Equals("''"))
            {
                SBQUERY.AppendFormat(@"   AND LA007 IN ({0}) ", LA007);
            }
            else
            {
                SBQUERY.AppendFormat(@"  ");
            }

            SB.AppendFormat(@" 
                            SELECT 
                            YEAR(LA015) AS 'YEARS',MONTH(LA015) AS 'MONTHS',LA005 AS '品號',MB002 AS '品名',MB003 AS '規格'
                            ,SUM(LA016-LA019+LA025) AS '銷售淨量',SUM(LA017-LA020-LA022-LA023) AS '銷貨淨額',SUM(LA024) AS '成本'
                            ,(SUM(LA017-LA020-LA022-LA023)-SUM(LA024)) AS '毛利'
                            ,(SUM(LA017-LA020-LA022-LA023)-SUM(LA024))/SUM(LA017-LA020-LA022-LA023) AS '毛利率'
                            ,CONVERT(NVARCHAR,YEAR(LA015))+'/'+CONVERT(NVARCHAR,MONTH(LA015)) AS 'YMS'

                            FROM [TK].dbo.SASLA
                            LEFT JOIN [TK].dbo.INVMB ON MB001=LA005
                            WHERE LA005 IN 
                            ( 
                            {2}
                            )
                            {3}
                            AND CONVERT(NVARCHAR,LA015,112)>='{0}' AND  CONVERT(NVARCHAR,LA015,112)<='{1}'
                            GROUP BY YEAR(LA015),MONTH(LA015),LA005,MB002,MB003
                            ORDER BY YEAR(LA015),MONTH(LA015),LA005

                            ", SDATES, EDATES, LA005, SBQUERY.ToString());

            return SB;

        }
        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            string MB001 = textBox4.Text.Trim() + "''";
            string ME001= textBox8.Text.Trim() + "''";
             
            SETFASTREPORT_SASLA(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), MB001, ME001);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Search_INVMB_DG1(textBox1.Text.Trim());
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox2.Text) && !string.IsNullOrEmpty(textBox3.Text))
            {
                textBox4.Text = textBox4.Text + "'" + textBox2.Text.Trim() + "','" + textBox3.Text.Trim() + "'," + Environment.NewLine;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox4.Text = null;
        }
        private void button5_Click(object sender, EventArgs e)
        {
            Search_CMSME_DG2(textBox5.Text.Trim());
        }


        private void button6_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox6.Text) && !string.IsNullOrEmpty(textBox7.Text))
            {
                textBox8.Text = textBox8.Text + "'" + textBox6.Text.Trim() + "','" + textBox7.Text.Trim() + "'," + Environment.NewLine;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            textBox8.Text = null;
        }
        #endregion


    }
}
