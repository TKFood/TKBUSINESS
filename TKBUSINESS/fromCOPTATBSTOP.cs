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
using TKITDLL;

namespace TKBUSINESS
{
    public partial class fromCOPTATBSTOP : Form
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

    

        public fromCOPTATBSTOP()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search(string TA002)
        {
            DataSet ds = new DataSet();
            StringBuilder sbSqlQUERY = new StringBuilder();            

            try
            {
                sbSql.Clear();
                sbSqlQUERY.Clear();

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.AppendFormat(@"
                                   SELECT TA001 AS '報價單別',TA002 AS '報價單號',TA006 AS '客戶',MV002 AS '業務員',TB004 AS '品號',TB005 AS '品名',TB007 AS '報價數量',TB008 AS '報價單位',TB009  AS '報價單價'
                                    ,(CASE WHEN TA022 IN ('1') THEN '內含'  WHEN TA022 IN ('2') THEN '外加'  WHEN TA022 IN ('3') THEN '零稅率' WHEN TA022 IN ('4') THEN '免稅' WHEN TA022 IN ('5') THEN '不計稅'  END  ) AS '稅別'
                                    ,TB016 AS '生效日期',TB017 AS '失效日期'

                                    ,TB006 AS '規格',TA004 AS '客代',TA005 AS '業務'
                                    FROM [TK].dbo.COPTB,[TK].dbo.COPTA
                                    LEFT JOIN [TK].dbo.CMSMV ON MV001=TA005
                                    WHERE 1=1
                                    AND TA001=TB001 AND TA002=TB002
                                    AND TA002 LIKE '{0}%'
                                    ORDER BY TA001,TA002



                                         ", TA002);

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
                    //rownum = ds.Tables[talbename].Rows.Count - 1;                       

                    dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                    dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 10);
                    dataGridView1.Columns["報價單別"].Width = 60;
                    dataGridView1.Columns["報價單號"].Width = 100;
                    dataGridView1.Columns["客戶"].Width = 100;
                    dataGridView1.Columns["業務員"].Width = 60;
                    dataGridView1.Columns["品號"].Width = 100;
                    dataGridView1.Columns["品名"].Width = 100;
                    dataGridView1.Columns["報價數量"].Width = 60;
                    dataGridView1.Columns["報價單位"].Width = 60;
                    dataGridView1.Columns["報價單價"].Width = 100;
                    dataGridView1.Columns["稅別"].Width = 60;
                    dataGridView1.Columns["生效日期"].Width = 100;
                    dataGridView1.Columns["失效日期"].Width = 100;
                    dataGridView1.Columns["規格"].Width = 100;
                    dataGridView1.Columns["客代"].Width = 100;
                    dataGridView1.Columns["業務"].Width = 100;

                }



            }
            catch
            {

            }
            finally
            {

            }
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox1.Text))
            {
                Search(textBox1.Text);
            }
           
        }

        #endregion
    }
}
