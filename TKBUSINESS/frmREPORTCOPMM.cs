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

namespace TKBUSINESS
{
    public partial class frmREPORTCOPMM : Form
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

        public frmREPORTCOPMM()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();
            
            Report report1 = new Report();

            if(comboBox1.Text.Trim().Equals("業務年月"))
            {
                report1.Load(@"REPORT\業務年月.frx");
                SQL1 = SETSQL();
            }
            else if (comboBox1.Text.Trim().Equals("業務年月客戶"))
            {
                report1.Load(@"REPORT\業務年月客戶.frx");
                SQL1 = SETSQL2();
            }
            else if (comboBox1.Text.Trim().Equals("業務年月客戶商品"))
            {
                report1.Load(@"REPORT\業務年月客戶商品.frx");
                SQL1 = SETSQL3();
            }
            else if (comboBox1.Text.Trim().Equals("業務年月客戶商品明細"))
            {
                report1.Load(@"REPORT\業務年月客戶商品明細.frx");
                SQL1 = SETSQL4();
            }
            

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            StringBuilder SB = new StringBuilder();
      
            SB.AppendFormat(@" 
            SELECT MM001,MM011,MV002,MN003,SUM(MN004) MN004,SUM(MN005)  MN005
            FROM [TK].dbo.[COPMM],[TK].dbo.[COPMN],[TK].dbo.COPMA,[TK].dbo.INVMB,[TK].dbo.CMSMV
            WHERE MM001=MN001 AND MM002=MN002
            AND MA001=MM003
            AND MB001=MM017
            AND MV001=MM011
            AND MM001='{0}'
            AND MM011='{1}'
            GROUP BY MM001,MM011,MV002,MN003
            ",dateTimePicker1.Value.ToString("yyyy"),textBox1.Text);

            return SB;

        }

        public StringBuilder SETSQL2()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
            SELECT MM001,MM011,MV002,MN003,MM003,MA002,SUM(MN004) MN004,SUM(MN005)  MN005
            FROM [TK].dbo.[COPMM],[TK].dbo.[COPMN],[TK].dbo.COPMA,[TK].dbo.INVMB,[TK].dbo.CMSMV
            WHERE MM001=MN001 AND MM002=MN002
            AND MA001=MM003
            AND MB001=MM017
            AND MV001=MM011
            AND MM001='{0}'
            AND MM011='{1}'
            GROUP BY MM001,MM011,MV002,MN003,MM003,MA002
            ", dateTimePicker1.Value.ToString("yyyy"), textBox1.Text);

            return SB;

        }

        public StringBuilder SETSQL3()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
            SELECT MM001,MM011,MV002,MN003,MM003,MA002,MM017,MB002,SUM(MN004) MN004,SUM(MN005)  MN005
            FROM [TK].dbo.[COPMM],[TK].dbo.[COPMN],[TK].dbo.COPMA,[TK].dbo.INVMB,[TK].dbo.CMSMV
            WHERE MM001=MN001 AND MM002=MN002
            AND MA001=MM003
            AND MB001=MM017
            AND MV001=MM011
            AND MM001='{0}'
            AND MM011='{1}'
            GROUP BY MM001,MM011,MV002,MN003,MM003,MA002,MM017,MB002
            ", dateTimePicker1.Value.ToString("yyyy"), textBox1.Text);

            return SB;

        }

        public StringBuilder SETSQL4()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
            SELECT MM001,MM011,MV002,MN003,MM003,MA002,MM017,MB002,SUM(MN004) MN004,SUM(MN005)  MN005
            FROM [TK].dbo.[COPMM],[TK].dbo.[COPMN],[TK].dbo.COPMA,[TK].dbo.INVMB,[TK].dbo.CMSMV
            WHERE MM001=MN001 AND MM002=MN002
            AND MA001=MM003
            AND MB001=MM017
            AND MV001=MM011
            AND MM001='{0}'
            AND MM011='{1}'
            AND MM003='{2}'
            GROUP BY MM001,MM011,MV002,MN003,MM003,MA002,MM017,MB002
            ", dateTimePicker1.Value.ToString("yyyy"), textBox1.Text, textBox2.Text);

            return SB;

        }


        #endregion

        #region BUTTON
        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        #endregion
    }
}
