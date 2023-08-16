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

        #endregion

       
    }
}
