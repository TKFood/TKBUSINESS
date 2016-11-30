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

namespace TKBUSINESS
{
    public partial class frmPRESALE : Form
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
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;
        string SALSESID=null;

        public frmPRESALE()
        {
            InitializeComponent();
            //tableLayoutPanel2.AutoScroll = true;
            //tableLayoutPanel2.AutoScrollMinSize = new Size(1000, 600);
            combobox1load();            
            combobox3load();
            combobox2load();

        }

        #region FUNCTION
        public void combobox1load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@" SELECT MV001,MV002  ");
            Sequel.AppendFormat(@" FROM CMSMV ");
            Sequel.AppendFormat(@" WHERE MV001   IN ('160092','070005','090002','140020','140049','140051','140078','150012','160155') ");
            Sequel.AppendFormat(@"  ");

            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);

            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MV001", typeof(string));
            dt.Columns.Add("MV002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MV001";
            comboBox1.DisplayMember = "MV002";
            sqlConn.Close();

        }
        public void combobox2load()
        {
            if(!string.IsNullOrEmpty(SALSESID))
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder Sequel = new StringBuilder();
                Sequel.AppendFormat(@" SELECT ME001,ME002 FROM CMSME ");
                Sequel.AppendFormat(@" WHERE ME002 NOT LIKE '%停用%' ");
                Sequel.AppendFormat(@"  AND ME001 IN (SELECT MV004 FROM CMSMV WHERE  MV001='{0}')", SALSESID.ToString());
                //Sequel.AppendFormat(@" AND EXISTS (SELECT MV004 FROM COPMA,CMSMV WHERE MA016=MV001 AND ISNULL(MA016,'')<>'' AND MV004=ME001  AND MV001='{0}')", SALSESID.ToString());
                Sequel.AppendFormat(@" ");

                SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
                DataTable dt = new DataTable();
                sqlConn.Open();

                dt.Columns.Add("ME001", typeof(string));
                dt.Columns.Add("ME002", typeof(string));
                 da.Fill(dt);
                comboBox2.DataSource = dt.DefaultView;
                comboBox2.ValueMember = "ME001";
                comboBox2.DisplayMember = "ME002";
                sqlConn.Close();
                if(dt.Rows.Count>0)
                {
                    textBox1.Text = dt.Rows[0]["ME001"].ToString();
                }
                
            }
           

        }

        public void combobox3load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@" SELECT MV001,MV002  ");
            Sequel.AppendFormat(@" FROM CMSMV ");
            Sequel.AppendFormat(@" WHERE MV001   IN ('160092','070005','090002','140020','140049','140051','140078','150012','160155') ");
            Sequel.AppendFormat(@"  ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);

            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MV001", typeof(string));
            dt.Columns.Add("MV002", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "MV001";
            comboBox3.DisplayMember = "MV002";
            textBox2.Text = dt.Rows[0]["MV001"].ToString(); 
            sqlConn.Close();

            SALSESID = textBox2.Text.ToString();
            combobox2load();

        }
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            SALSESID = comboBox3.SelectedValue.ToString();
            textBox2.Text = SALSESID;
            combobox2load();
        }
       
        public void Search()
        {
            try
            {
                sbSql.Clear();
                sbSql = SETsbSql();

                if (!string.IsNullOrEmpty(sbSql.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);
                    adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, tablename);
                    sqlConn.Close();

                    labelget.Text = "資料筆數:" + ds.Tables[tablename].Rows.Count.ToString();

                    if (ds.Tables[tablename].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        dataGridView1.DataSource = ds.Tables[tablename];
                        dataGridView1.AutoResizeColumns();
                        //rownum = ds.Tables[talbename].Rows.Count - 1;
                        dataGridView1.CurrentCell = dataGridView1.Rows[rownum].Cells[0];

                        //dataGridView1.CurrentCell = dataGridView1[0, 2];

                    }
                }
                else
                {

                }



            }
            catch
            {

            }
            finally
            {

            }
        }

        public StringBuilder SETsbSql()
        {
            StringBuilder STR = new StringBuilder();

            STR.AppendFormat(@"  SELECT [YEARS] AS '年度',[MONTHS] AS '月份',[DEPID] AS '部門代號' ,[DEPNAME] AS '部門名'");
            STR.AppendFormat(@"  ,[SALESID] AS '業務員代號',[SALESNAME] AS '業務名',[CUSTOMERID] AS '客戶代號',[CUSTOMERNAME] AS '客戶名' ");
            STR.AppendFormat(@"  ,[PRODUCTID] AS '商品代號',[PRODUCTNAME] AS '商品名'");
            STR.AppendFormat(@"  ,[PRICES] AS '單價',[NUM] AS '數量',[TMONEY] AS '金額',[ID]");
            STR.AppendFormat(@"   FROM [TKBUSINESS].[dbo].[PRESALE]");
            STR.AppendFormat(@"  ");
            STR.AppendFormat(@"  ");
            tablename = "TEMPds1";

            return STR;
        }

        public void ExcelExport()
        {
            Search();

            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;


            dt = ds.Tables[tablename];
            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }


            int j = 0;
            if (tablename.Equals("TEMPds1"))
            {
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString()));
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString()));
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString()));

                    j++;
                }
            }


            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\查詢{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }


        }
        #endregion

        #region BUTTON
       
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }

        #endregion

        
    }
}
