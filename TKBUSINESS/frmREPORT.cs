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
    public partial class frmREPORT : Form
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
        string mdate;

        public frmREPORT()
        {
            InitializeComponent();
        }

        #region FUNCTION
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

            if (comboBox1.Text.ToString().Equals("客戶收入預估"))
            {      
                STR.AppendFormat(@" SELECT [CUSTOMERNAME] AS '客戶名稱',[YEARS] AS '年度',CONVERT(INT,[MONTHS])  AS '月份', SUM([TMONEY])  AS '金額'  ");
                STR.AppendFormat(@"  FROM [TKBUSINESS].[dbo].[PRESALE] ");
                STR.AppendFormat(@"  WHERE   [YEARS]>='{0}'AND CONVERT(INT,[MONTHS])>='{1}'", dateTimePicker1.Value.ToString("yyyy"), Convert.ToInt16(dateTimePicker1.Value.ToString("MM")));
                STR.AppendFormat(@"  AND [YEARS]<='{0}'AND CONVERT(INT,[MONTHS])<='{1}'", dateTimePicker2.Value.ToString("yyyy"), Convert.ToInt16(dateTimePicker2.Value.ToString("MM")));
                STR.AppendFormat(@"  GROUP BY  [CUSTOMERNAME],[YEARS],CONVERT(INT,[MONTHS])");
                STR.AppendFormat(@"  ORDER BY [CUSTOMERNAME],[YEARS],CONVERT(INT,[MONTHS])");
                STR.AppendFormat(@"  ");
                tablename = "TEMPds1";
            }
            else if (comboBox1.Text.ToString().Equals("電子商務明細"))
            {
                STR.AppendFormat(@"  SELECT [YEARS] AS '年度',CONVERT(INT,[MONTHS]) AS '月份' ,[PRODUCTID] AS '商品代號',[PRODUCTNAME] AS '商品名' ,SUM([PRICES]) AS '單價',SUM([NUM]) AS '數量',SUM([TMONEY]) AS '金額'  ");
                STR.AppendFormat(@"  FROM [TKBUSINESS].[dbo].[PRESALE]   ");
                STR.AppendFormat(@"  WHERE   [YEARS]>='2017'AND CONVERT(INT,[MONTHS])>='{1}'", dateTimePicker1.Value.ToString("yyyy"), Convert.ToInt16(dateTimePicker1.Value.ToString("MM"))); ;
                STR.AppendFormat(@"  AND [YEARS]<='2017'AND CONVERT(INT,[MONTHS])<='{1}'", dateTimePicker2.Value.ToString("yyyy"), Convert.ToInt16(dateTimePicker2.Value.ToString("MM")));
                STR.AppendFormat(@"  AND [CUSTOMERID]='1ZZZZZZZ'");
                STR.AppendFormat(@"  GROUP BY [YEARS],CONVERT(INT,[MONTHS]),[CUSTOMERID],[PRODUCTID],[PRODUCTNAME] ");
                STR.AppendFormat(@"  ORDER BY [YEARS],CONVERT(INT,[MONTHS]),[CUSTOMERID]");
                STR.AppendFormat(@"  ");
                tablename = "TEMPds2";
            }
            else if (comboBox1.Text.ToString().Equals("消費者及員購明細"))
            {
                STR.AppendFormat(@"  SELECT [YEARS] AS '年度',CONVERT(INT,[MONTHS]) AS '月份' ,[PRODUCTID] AS '商品代號',[PRODUCTNAME] AS '商品名' ,SUM([PRICES]) AS '單價',SUM([NUM]) AS '數量',SUM([TMONEY]) AS '金額'  ");
                STR.AppendFormat(@"  FROM [TKBUSINESS].[dbo].[PRESALE]   ");
                STR.AppendFormat(@"  WHERE   [YEARS]>='2017'AND CONVERT(INT,[MONTHS])>='{1}'", dateTimePicker1.Value.ToString("yyyy"), Convert.ToInt16(dateTimePicker1.Value.ToString("MM"))); ;
                STR.AppendFormat(@"  AND [YEARS]<='2017'AND CONVERT(INT,[MONTHS])<='{1}'", dateTimePicker2.Value.ToString("yyyy"), Convert.ToInt16(dateTimePicker2.Value.ToString("MM")));
                STR.AppendFormat(@"  AND [CUSTOMERID]='1ZZZZZZA'");
                STR.AppendFormat(@"  GROUP BY [YEARS],CONVERT(INT,[MONTHS]),[CUSTOMERID],[PRODUCTID],[PRODUCTNAME] ");
                STR.AppendFormat(@"  ORDER BY [YEARS],CONVERT(INT,[MONTHS]),[CUSTOMERID]");
                STR.AppendFormat(@"  ");
                tablename = "TEMPds2";
            }

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
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                   
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
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }
        #endregion



    }
}
