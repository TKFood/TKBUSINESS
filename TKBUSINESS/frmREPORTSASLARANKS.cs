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
    public partial class frmREPORTSASLARANKS : Form
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
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;

        public frmREPORTSASLARANKS()
        {
            InitializeComponent();

            SETDATES();
            comboBox1load();
        }

        #region FUNCTION
        public void comboBox1load()
        {
            DataTable dt = new DataTable();
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);


            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"
                                SELECT [NAMES]
                                FROM [TKBUSINESS].[dbo].[TBPARA]
                                WHERE [KINDS]='SASLALA007'
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
           
            sqlConn.Open();
           
            dt.Columns.Add("NAMES", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.DisplayMember = "NAMES";
            sqlConn.Close();


        }
        public void SETDATES()
        {
            DateTime currentDate = DateTime.Now;
            DateTime firstDayOfLastMonth = new DateTime(currentDate.Year, currentDate.Month - 1, 1);
            dateTimePicker1.Value = firstDayOfLastMonth;
            dateTimePicker2.Value = firstDayOfLastMonth;
        }
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            CHECKDATE_SDAYS(dateTimePicker1.Value);
        }
        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            CHECKDATE_EDAYS(dateTimePicker2.Value);
        }
        public void CHECKDATE_SDAYS(DateTime SDAYS)
        {
            DateTime currentDate = DateTime.Now;
            DateTime firstDayOfMonth = new DateTime(currentDate.Year, currentDate.Month, 1);
            DateTime firstDayOfLastMonth = new DateTime(currentDate.Year, currentDate.Month-1, 1);

            DateTime yourDateTime = SDAYS;// 設定您想要檢查的 DateTime 物件

            if (yourDateTime > firstDayOfMonth)
            {
                // 您的 DateTime 大於或等於本月第一天
                dateTimePicker1.Value = firstDayOfLastMonth;
                MessageBox.Show("日期只能在上個月之前");
            }
            else
            {
               
            }
        }
        public void CHECKDATE_EDAYS(DateTime EDAYS)
        {
            DateTime currentDate = DateTime.Now;
            DateTime firstDayOfMonth = new DateTime(currentDate.Year, currentDate.Month, 1);
            DateTime firstDayOfLastMonth = new DateTime(currentDate.Year, currentDate.Month - 1, 1);

            DateTime yourDateTime = EDAYS;// 設定您想要檢查的 DateTime 物件

            if (yourDateTime > firstDayOfMonth)
            {
                // 您的 DateTime 大於或等於本月第一天
                dateTimePicker2.Value = firstDayOfLastMonth;
                MessageBox.Show("日期只能在上個月之前");
            }
            else
            {

            }
        }

        #endregion
        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {

        }

        #endregion

       
    }
}
