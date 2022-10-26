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
        private void fromCOPTATBSTOP_Load(object sender, EventArgs e)
        {
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 50;   //設定寬度
            cbCol.HeaderText = "　選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView1.Columns.Insert(0, cbCol);

            #region 建立全选 CheckBox

            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            Rectangle rect = dataGridView1.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 4 - 18;
            rect.Y = rect.Location.Y + (rect.Height / 2 - 9);

            CheckBox cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(18, 18);
            cbHeader.Location = rect.Location;

            //全选要设定的事件
            cbHeader.CheckedChanged += new EventHandler(cbHeader_CheckedChanged);

            //将 CheckBox 加入到 dataGridView
            dataGridView1.Controls.Add(cbHeader);


            #endregion
        }

        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView1.Controls.Find("checkboxHeader", true)[0]).Checked;

            }



        }

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
                                    ,TA007 AS '幣別'

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
                    dataGridView1.Columns["幣別"].Width = 100;

                }



            }
            catch
            {

            }
            finally
            {

            }
        }

        public void SETTB017MB018()
        {
            string TA001TA002TB004TB016 = null;
            string MB001MB002MB003MB004MB017 = null;

            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {


                    TA001TA002TB004TB016 = TA001TA002TB004TB016 + "'" + dr.Cells["報價單別"].Value.ToString()+ dr.Cells["報價單號"].Value.ToString()+ dr.Cells["品號"].Value.ToString()+ dr.Cells["生效日期"].Value.ToString() + "'";

                    TA001TA002TB004TB016 = TA001TA002TB004TB016 + ",";

                }

            }

            TA001TA002TB004TB016 = TA001TA002TB004TB016 + "''";

            UPDATECOPTBTB017(TA001TA002TB004TB016, dateTimePicker1.Value.ToString("yyyyMMdd"));

            //MessageBox.Show(TA001TA002TB004);
        }
        public void SETTB017MB018NULL()
        {
            string TA001TA002TB004TB016 = null;
            string MB001MB002MB003MB004MB017 = null;

            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {


                    TA001TA002TB004TB016 = TA001TA002TB004TB016 + "'" + dr.Cells["報價單別"].Value.ToString() + dr.Cells["報價單號"].Value.ToString() + dr.Cells["品號"].Value.ToString() + dr.Cells["生效日期"].Value.ToString() + "'";

                    TA001TA002TB004TB016 = TA001TA002TB004TB016 + ",";

                }

            }

            TA001TA002TB004TB016 = TA001TA002TB004TB016 + "''";

            UPDATECOPTBTB017(TA001TA002TB004TB016,"");

            //MessageBox.Show(TA001TA002TB004);
        }

        public void UPDATECOPTBTB017(string TA001TA002TB004TB016, string TB017)
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


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                //AND[購物車編號] = 'TG220719S01169'
                //COPTG --AND [購物車編號]  NOT IN (SELECT TG020 FROM [TK].dbo.COPTG WHERE ISNULL(TG020,'')<>'')
                //COPTH --AND [訂單編號]  NOT IN (SELECT TH074 FROM [TK].dbo.COPTH WHERE ISNULL(TH074,'')<>'')

                //INSERT INTO [test0923].[dbo].[COPTG]
                sbSql.AppendFormat(@" 
                                        UPDATE [TK].dbo.COPTB
                                        SET TB017='{1}'
                                        WHERE LTRIM(RTRIM(TB001))+LTRIM(RTRIM(TB002))+LTRIM(RTRIM(TB004))+LTRIM(RTRIM(TB016))
                                        IN
                                        (
                                        {0}
                                        )
                                        ", TA001TA002TB004TB016, TB017);


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易                      
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

        public void UPDATECOMMBMB018()
        {

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
        private void button2_Click(object sender, EventArgs e)
        {
            SETTB017MB018();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SETTB017MB018NULL();
        }

        #endregion


    }
}
