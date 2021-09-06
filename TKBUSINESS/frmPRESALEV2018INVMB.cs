using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Reflection;
using System.Threading;
using System.Globalization;
using TKITDLL;

namespace TKBUSINESS
{
    public partial class frmPRESALEV2018INVMB : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();



        DataTable dt = new DataTable();
        string tablename = null;
        int result;

        string MB001;


        public frmPRESALEV2018INVMB()
        {
            InitializeComponent();
            this.textBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox1_KeyDown);
        }

        #region FUNCTION
        public void SEARCHINVMB()
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

                sbSql.Clear();
                sbSqlQuery.Clear();

                if (string.IsNullOrEmpty(textBox1.Text))
                {
                    sbSql.AppendFormat(@"SELECT MB001,MB002,MB003 FROM [TK].dbo.INVMB ORDER BY MB001  ");
                }
                else
                {
                    sbSql.AppendFormat(@"SELECT MB001,MB002,MB003 FROM [TK].dbo.INVMB WHERE MB001 LIKE '{0}%' OR MB002 LIKE '{0}%' ORDER BY MB001", textBox1.Text);
                }

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {

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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    MB001 = row.Cells["MB001"].Value.ToString();

                }
                else
                {
                    MB001 = null;

                }
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (Keys.Enter == e.KeyCode)
            {
                SEARCHINVMB();
            }
        }

        public string TextBoxMsg
        {
            set
            {

            }
            get
            {
                return MB001;
            }
        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHINVMB();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        #endregion

      
    }
}
