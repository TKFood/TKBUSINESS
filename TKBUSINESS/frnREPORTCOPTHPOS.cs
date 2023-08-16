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
    public partial class frnREPORTCOPTHPOS : Form
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
     
        string tablename = null;
        int rownum = 0;
        public frnREPORTCOPTHPOS()
        {
            InitializeComponent();

            textBox5.Text = DateTime.Now.Year.ToString();
        }


        #region FUNCTION
        public void SETFASTREPORT(string SDATES, string EDATES,string MB001)
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


            SQL1 = SETSQL(SDATES, EDATES, MB001); 
            Report report1 = new Report();
            report1.Load(@"REPORT\銷貨單業績.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();


            report1.SetParameterValue("P1", P1);
            report1.SetParameterValue("P2", P2);

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL(string SDATES, string EDATES,string MB001)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT 
                            MV002 AS '業務員'

                            ,TH004 AS '品號'
                            ,TH005 AS '品名'
                            ,SUM(TH008) TH008
                            ,SUM(TH037) AS '未稅金額'
                            ,TH025 AS 折扣率
                            ,MD003
                            ,MD004
                            ,(CASE WHEN ISNULL(MD004,0)<>0 THEN SUM(TH008)*MD004/MD003 ELSE SUM(TH008) END ) AS  '銷售數量'
                            FROM [TK].dbo.COPTG
                            LEFT JOIN [TK].dbo.CMSMV ON MV001=TG006
                            ,[TK].dbo.COPTH
                            LEFT JOIN [TK].dbo.INVMD ON MD001=TH004 AND MD002=TH009
                            WHERE 1=1
                            AND TG001=TH001 AND TG002=TH002
                            AND TH020='Y'
                            AND TG003>='{0}' AND TG003<='{1}'
                            AND TH004 IN (
                            {2}
                            )
                            GROUP BY TG006,MV002,TH004,TH005,TH025,MD003,MD004
                            ORDER BY MV002,TG006,TH004,TH005,TH025,MD003,MD004

                            ", SDATES, EDATES, MB001);

            return SB;

        }

        public void SETFASTREPORT_POSTB(string TB036)
        {
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


            SQL1 = SETSQL_POSTB(TB036);
            Report report1 = new Report();
            report1.Load(@"REPORT\POS活動業績.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", P1);
            //report1.SetParameterValue("P2", P2);

            report1.Preview = previewControl2;
            report1.Show();
        }

        public StringBuilder SETSQL_POSTB(string TB036)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT 
                            TB002 AS '門市代'
                            ,MA002 AS '門市'
                            ,TB010 AS '品號'
                            ,MB002 AS '品名'
                            ,SUM(TB019) AS '銷售數量'
                            ,SUM(TB031) AS '未稅金額'
                            FROM [TK].dbo.POSTB
                            LEFT JOIN [TK].dbo.INVMB ON MB001=TB010
                            LEFT JOIN [TK].dbo.WSCMA ON MA001=TB002 
                            WHERE TB036='{0}'
                            GROUP BY  TB002,MA002,TB010,MB002
                            ORDER BY  TB002,MA002,TB010,MB002

                            ", TB036);

            return SB;

        }

        public void SETFASTREPORT_POSTB_V2(string SDATES,string EDATES,string MB001)
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


            SQL1 = SETSQL_POSTB_V2(SDATES, EDATES, MB001);
            Report report1 = new Report();
            report1.Load(@"REPORT\POS業績.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.SetParameterValue("P1", P1);
            report1.SetParameterValue("P2", P2);

            report1.Preview = previewControl3;
            report1.Show();
        }

        public StringBuilder SETSQL_POSTB_V2(string SDATES, string EDATES, string MB001)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"                           
                            SELECT 
                            TB002 AS '門市代'
                            ,MA002 AS '門市'
                            ,TB010 AS '品號'
                            ,MB002 AS '品名'
                            ,SUM(TB019) AS '銷售數量'
                            ,SUM(TB031) AS '未稅金額'
                            FROM [TK].dbo.POSTB
                            LEFT JOIN [TK].dbo.INVMB ON MB001=TB010
                            LEFT JOIN [TK].dbo.WSCMA ON MA001=TB002 
                            WHERE 1=1
                            AND TB001>='{0}' AND TB001<='{1}'
                            AND TB010 IN 
                            (
                            {2}
                            )
                            GROUP BY  TB002,MA002,TB010,MB002
                            ORDER BY  TB002,MA002,TB010,MB002

                            ", SDATES, EDATES, MB001);

            return SB;

        }
        public void Search_INVMB(string MB001)
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
                                    SELECT MB001 AS '品號',MB002  AS '品名'
                                    FROM [TK].dbo.INVMB
                                    WHERE (MB001 LIKE '%{0}%' OR MB002 LIKE '%{0}%')
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

        public void Search_INVMB_V2(string MB001)
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
                                    SELECT MB001 AS '品號',MB002  AS '品名'
                                    FROM [TK].dbo.INVMB
                                    WHERE (MB001 LIKE '%{0}%' OR MB002 LIKE '%{0}%')
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
                    dataGridView3.DataSource = null;
                }
                else
                {
                    dataGridView3.DataSource = ds.Tables["ds"];
                    dataGridView3.AutoResizeColumns();
                    dataGridView3.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                    dataGridView3.DefaultCellStyle.Font = new Font("Tahoma", 10);
                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        public void Search_POSMA(string MA001,string MA002)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            try
            {
                sbSqlQuery.Clear();
                if(!string.IsNullOrEmpty(MA002))
                {
                    sbSqlQuery.AppendFormat(@"  AND (特價代號 LIKE '%{0}%' OR 特價名稱 LIKE '%{0}%' )", MA002);
                }
                else
                {
                    sbSqlQuery.AppendFormat(@" ");
                }

                sbSql.Clear();
                sbSql.AppendFormat(@"
                                   SELECT *
                                    FROM
                                    (
                                    SELECT MA001 AS '活動代號',MA002 AS '活動名稱',MB003 AS  '特價代號',MB004 AS '特價名稱'
                                    FROM [TK].dbo.POSMA
                                    LEFT JOIN [TK].dbo.POSMB ON MB001=MA001
                                    UNION ALL
                                    SELECT MA001 AS '活動代號',MA002 AS '活動名稱',MI003 AS  '特價代號',MI004 AS '特價名稱'
                                    FROM [TK].dbo.POSMA
                                    LEFT JOIN [TK].dbo.POSMI ON MI001=MA001
                                    UNION ALL
                                    SELECT MA001 AS '活動代號',MA002 AS '活動名稱',MM003 AS  '特價代號',MM004 AS '特價名稱'
                                    FROM [TK].dbo.POSMA
                                    LEFT JOIN [TK].dbo.POSMM ON MM001=MA001
                                    UNION ALL
                                    SELECT MA001 AS '活動代號',MA002 AS '活動名稱',MO003 AS  '特價代號',MO004 AS '特價名稱'
                                    FROM [TK].dbo.POSMA
                                    LEFT JOIN [TK].dbo.POSMO ON MO001=MA001
                                    ) AS TEMP
                                    WHERE 活動代號 LIKE '{0}%' 
                                    {1}
                                    ORDER BY 活動代號,特價代號

                                    ", MA001, sbSqlQuery.ToString());


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


        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox3.Text = null;
            textBox4.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                  
                    textBox3.Text = row.Cells["品號"].Value.ToString().Trim();
                    textBox4.Text = row.Cells["品名"].Value.ToString().Trim();
                }
            }

        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            textBox6.Text = null;
            textBox8.Text = null;

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];

                    textBox6.Text = row.Cells["特價代號"].Value.ToString().Trim();
                    textBox8.Text = row.Cells["特價名稱"].Value.ToString().Trim();
                }
            }
        }
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            textBox10.Text = null;
            textBox11.Text = null;

            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];

                    textBox10.Text = row.Cells["品號"].Value.ToString().Trim();
                    textBox11.Text = row.Cells["品名"].Value.ToString().Trim();
                }
            }
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            string MB001 = textBox2.Text.Trim() + "''" ;
            SETFASTREPORT(dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"), MB001);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Search_INVMB(textBox1.Text.Trim());
        }
        private void button3_Click(object sender, EventArgs e)
        {          
            if (!string.IsNullOrEmpty(textBox3.Text)&& !string.IsNullOrEmpty(textBox4.Text))
            {
                textBox2.Text = textBox2.Text + "'" + textBox3.Text.Trim() + "','" + textBox4.Text.Trim()+"',"+Environment.NewLine;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox2.Text = null;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Search_POSMA(textBox5.Text.Trim(), textBox7.Text.Trim());
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SETFASTREPORT_POSTB(textBox6.Text.Trim());
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Search_INVMB_V2(textBox9.Text.Trim());
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox10.Text) && !string.IsNullOrEmpty(textBox11.Text))
            {
                textBox12.Text = textBox12.Text + "'" + textBox10.Text.Trim() + "','" + textBox11.Text.Trim() + "'," + Environment.NewLine;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            textBox12.Text = null;
        }
        private void button10_Click(object sender, EventArgs e)
        {
            string MB001 = textBox12.Text.Trim() + "''";
            SETFASTREPORT_POSTB_V2(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), MB001);
        }

        #endregion


    }
}
