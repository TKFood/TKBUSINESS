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
    public partial class frmREPORTCOPTGHGSET : Form
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

        public frmREPORTCOPTGHGSET()
        {
            InitializeComponent();

            comboBox1load();
            comboBox2load();

            comboBox1.Text = "全部";
        }

        #region FUNCTION
        public void LoadComboBoxData(ComboBox comboBox, string query, string valueMember, string displayMember)
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            using (SqlConnection connection = new SqlConnection(sqlsb.ConnectionString))
            {
                SqlCommand command = new SqlCommand(query, connection);
                connection.Open();

                SqlDataAdapter adapter = new SqlDataAdapter(command);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                comboBox.DataSource = dataTable;
                comboBox.ValueMember = valueMember;
                comboBox.DisplayMember = displayMember;
            }
        }
        public void comboBox1load()
        {
            LoadComboBoxData(comboBox1, "SELECT [KINDS],[NAMES],[VALUE] FROM [TKBUSINESS].[dbo].[TBPARA] WHERE [KINDS]='TH020' ORDER BY ID", "NAMES", "NAMES");
        }
        public void comboBox2load()
        {
            LoadComboBoxData(comboBox2, "SELECT [KINDS],[NAMES],[VALUE] FROM [TKBUSINESS].[dbo].[TBPARA] WHERE [KINDS]='REPORT1' ORDER BY ID", "NAMES", "NAMES");
        }

        public void SETFASTREPORT(string REPORTSKIND,string SDATES, string EDATES, string TG023)
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

            if (REPORTSKIND.Equals("官網銷貨備貨"))
            {
                SQL1 = SETSQL1(SDATES, EDATES, TG023);

                report1.Load(@"REPORT\銷貨備貨V1.frx");
            }
            else if (REPORTSKIND.Equals("官網銷貨備貨明細"))
            {
                SQL1 = SETSQL2(SDATES, EDATES, TG023);

                report1.Load(@"REPORT\銷貨備貨明細.frx");
            }          

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();


            report1.SetParameterValue("P1", P1);
            report1.SetParameterValue("P2", P2);

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1(string SDATES, string EDATES, string TG023)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SBQUERY = new StringBuilder();
            StringBuilder SBQUERY2 = new StringBuilder();

            if (!string.IsNullOrEmpty(TG023) && TG023.Equals("Y"))
            {
                SBQUERY.AppendFormat(@" 
                                       AND TG023='Y' 
                                        ");
            }
            else if (!string.IsNullOrEmpty(TG023) && TG023.Equals("N"))
            {
                SBQUERY.AppendFormat(@" 
                                       AND TG023='N' 
                                        ");
            }
            else if (!string.IsNullOrEmpty(TG023) && TG023.Equals("全部"))
            {
                SBQUERY.AppendFormat(@" 
                                         AND TG023<>'V' 
                                        ");
            }


            SB.AppendFormat(@" 
                           SELECT TH004 AS '品號',TH005 AS '品名',TH009 AS '單位',SUM(TH008+TH024) AS '數量'
                            FROM [TK].dbo.COPTG,[TK].dbo.COPTH
                            WHERE 1=1
                            AND TG001=TH001 AND TG002=TH002
                            AND TG001='A233'
                            AND TG003>='{0}' AND TG003<='{1}'
                            {2}

                            GROUP BY TH004,TH005,TH009
                            ORDER BY TH004


                            ",SDATES,EDATES, SBQUERY.ToString());

            return SB;
        }

        public StringBuilder SETSQL2(string SDATES, string EDATES, string TG023)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SBQUERY = new StringBuilder();
            StringBuilder SBQUERY2 = new StringBuilder();

            if (!string.IsNullOrEmpty(TG023) && TG023.Equals("Y"))
            {
                SBQUERY.AppendFormat(@" 
                                       AND TG023='Y' 
                                        ");
            }
            else if (!string.IsNullOrEmpty(TG023) && TG023.Equals("N"))
            {
                SBQUERY.AppendFormat(@" 
                                       AND TG023='N' 
                                        ");
            }
            else if (!string.IsNullOrEmpty(TG023) && TG023.Equals("全部"))
            {
                SBQUERY.AppendFormat(@" 
                                    AND TG023<>'V' 
                                        ");
            }


            SB.AppendFormat(@" 
                            SELECT TG001 AS '銷貨單',TG002 AS '銷貨單號',TH004 AS '品號',TH005 AS '品名',TH009 AS '單位',SUM(TH008+TH024) AS '數量'
                            FROM [TK].dbo.COPTG,[TK].dbo.COPTH
                            WHERE 1=1
                            AND TG001=TH001 AND TG002=TH002
                            AND TG001='A233'
                            AND TG003>='{0}' AND TG003<='{1}'
                            {2}

                            GROUP BY  TG001,TG002,TH004,TH005,TH009
                            ORDER BY  TG001,TG002,TH004


                            ", SDATES, EDATES, SBQUERY.ToString());

            return SB;
        }

        #endregion

        #region BUTTON


        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(comboBox2.Text.ToString(),dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"), comboBox1.Text.ToString());
        }

        #endregion
    }
}
