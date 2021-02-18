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
    public partial class frmREPORTCOPTH : Form
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



        public frmREPORTCOPTH()
        {
            InitializeComponent();
        }


        #region FUNCTION
        private void frmREPORTCOPTH_Load(object sender, EventArgs e)
        {
            //dataGridView1
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 50;   //設定寬度
            cbCol.HeaderText = "選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView1.Columns.Insert(0, cbCol);
        }
        public void Search(string MB001)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();
                

               
                sbSql.AppendFormat(@"  
                                    SELECT MB001 AS '品號',MB002 AS '品名'
                                    FROM [TK].dbo.INVMB
                                    WHERE (MB001 LIKE '{0}%' OR MB002 LIKE '{0}%')
                                    ORDER BY MB001
                                    ", MB001);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();

                      
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

        public void ADDTOTEXTBOX()
        {
            string MB001;

            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                DataGridViewCheckBoxCell cbx = (DataGridViewCheckBoxCell)dr.Cells[0];

                if ((bool)cbx.FormattedValue)
                {
                    MB001 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString().Trim();
                    

                    //MessageBox.Show(TA001 + "-"+ TA002);
                    if (!string.IsNullOrEmpty(MB001))
                    {
                        textBox2.AppendText(MB001 + Environment.NewLine);
                    }
                }
                else
                {
                    
                }
            }
        }

        public void CLEARTEXTBOX()
        {
            textBox2.Text = null;
        }

        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();
            report1.Load(@"REPORT\銷售商品-客戶表.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", textBox1.Text);
            //report1.SetParameterValue("P2", textBox2.Text);
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            StringBuilder SB = new StringBuilder();
            string strLineData = FINDTEXTBOX();
            //MessageBox.Show(strLineData);

            SB.AppendFormat(@" 
                            SELECT TG004,TG007,TH004,TH005,SUM(LA011) LA011,SUM(TH037)  TH037
                            FROM [TK].dbo.COPTG,[TK].dbo.COPTH,[TK].dbo.INVLA
                            WHERE TG001=TH001 AND TG002=TH002
                            AND TH001=LA006 AND TH002=LA007 AND TH003=LA008
                            AND TG004 NOT LIKE'1%'
                            AND TG023='Y'
                            AND TG003>='{0}' AND TG003<='{1}'
                            AND TH004 IN ({2})
                            GROUP BY TG004,TG007,TH004,TH005
                            ORDER BY TG004,TG007,TH004,TH005
                            ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), strLineData);

            return SB;

        }

        public string FINDTEXTBOX()
        {
            string strLineData = null;
            StringBuilder ReturnMB001 = new StringBuilder();

            for (int i = 0; i < textBox2.Lines.Length; i++)
            {
                ReturnMB001.AppendFormat("'" + textBox2.Lines[i] + "'" + ",");
            }
           

            ReturnMB001.AppendFormat("'9'");
            return ReturnMB001.ToString(); 
        }

        #endregion




        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            if(!String.IsNullOrEmpty(textBox1.Text.Trim()))
            {
                Search(textBox1.Text.Trim());
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ADDTOTEXTBOX();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            CLEARTEXTBOX();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        #endregion


    }
}
