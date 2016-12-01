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
        int result;
        DataGridViewRow drPRESLAES = new DataGridViewRow();

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
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

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
                        
                        textBox3.Text = null;
                        textBox4.Text = null;
                        textBox5.Text = null;
                        textBox6.Text = null;
                        textBox7.Text = null;                        
                        textBox9.Text = null;
                        textBox10.Text = null;
                        textBoxID.Text = null;

                    }
                    else
                    {
                        dataGridView1.DataSource = ds.Tables[tablename];
                        dataGridView1.AutoResizeColumns();
                        //rownum = ds.Tables[talbename].Rows.Count - 1;
                        dataGridView1.CurrentCell = dataGridView1.Rows[rownum].Cells[8];

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
            StringBuilder STRQUERY = new StringBuilder();
            if (!string.IsNullOrEmpty(textBox8.Text.ToString()))
            {
                STRQUERY.AppendFormat(@" AND [CUSTOMERID] LIKE '{0}%'", textBox8.Text.ToString());
            }

            STR.AppendFormat(@"  SELECT [YEARS] AS '年度',[MONTHS] AS '月份',[DEPID] AS '部門代號' ,[DEPNAME] AS '部門名'");
            STR.AppendFormat(@"  ,[SALESID] AS '業務員代號',[SALESNAME] AS '業務名',[CUSTOMERID] AS '客戶代號',[CUSTOMERNAME] AS '客戶名' ");
            STR.AppendFormat(@"  ,[PRODUCTID] AS '商品代號',[PRODUCTNAME] AS '商品名'");
            STR.AppendFormat(@"  ,[PRICES] AS '單價',[NUM] AS '數量',[TMONEY] AS '金額',[ID]");
            STR.AppendFormat(@"   FROM [TKBUSINESS].[dbo].[PRESALE]");
            STR.AppendFormat(@"   WHERE [SALESID]='{0}'",comboBox1.SelectedValue.ToString());
            STR.AppendFormat(@"   AND  [YEARS]>='{0}' AND CONVERT(INT,[MONTHS])>='{1}'", dateTimePicker1.Value.ToString("yyyy"), Convert.ToInt16(dateTimePicker1.Value.ToString("MM")));
            STR.AppendFormat(@"   AND  [YEARS]<='{0}' AND CONVERT(INT,[MONTHS])<='{1}'", dateTimePicker2.Value.ToString("yyyy"), Convert.ToInt16(dateTimePicker2.Value.ToString("MM")));
            STR.AppendFormat(@"  {0}", STRQUERY.ToString());
            STR.AppendFormat(@"   ORDER BY [YEARS],CONVERT(INT,[MONTHS]),[CUSTOMERID]");
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
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString());
                    ws.GetRow(j + 1).CreateCell(8).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString());
                    ws.GetRow(j + 1).CreateCell(9).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString());
                    ws.GetRow(j + 1).CreateCell(10).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[10].ToString()));
                    ws.GetRow(j + 1).CreateCell(11).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[11].ToString()));
                    ws.GetRow(j + 1).CreateCell(12).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[12].ToString()));

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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count >= 1)
            {
                drPRESLAES = dataGridView1.Rows[dataGridView1.SelectedCells[0].RowIndex];

                comboBox2.Text= drPRESLAES.Cells["部門名"].Value.ToString();
                comboBox3.Text = drPRESLAES.Cells["業務名"].Value.ToString();
                numericUpDown3.Value = Convert.ToDecimal(drPRESLAES.Cells["年度"].Value.ToString());
                numericUpDown4.Value = Convert.ToDecimal(drPRESLAES.Cells["月份"].Value.ToString());
                textBox1.Text = drPRESLAES.Cells["部門代號"].Value.ToString();
                textBox2.Text = drPRESLAES.Cells["業務員代號"].Value.ToString();
                textBox3.Text = drPRESLAES.Cells["商品代號"].Value.ToString();
                textBox4.Text = drPRESLAES.Cells["商品名"].Value.ToString();
                textBox5.Text = drPRESLAES.Cells["單價"].Value.ToString();
                textBox6.Text = drPRESLAES.Cells["數量"].Value.ToString();
                textBox7.Text = drPRESLAES.Cells["金額"].Value.ToString();
                textBoxID.Text = drPRESLAES.Cells["ID"].Value.ToString();
                textBox9.Text = drPRESLAES.Cells["客戶代號"].Value.ToString();
                textBox10.Text = drPRESLAES.Cells["客戶名"].Value.ToString();


            }
            else
            {
                
            }
        }
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            SEARCHPRODUCTNAME();
        }

        public void SEARCHPRODUCTNAME()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@" SELECT MB001,MB002  ");
            Sequel.AppendFormat(@" FROM INVMB ");
            Sequel.AppendFormat(@" WHERE MB001='{0}'",textBox3.Text.ToString());
            Sequel.AppendFormat(@"  ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);

            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MB001", typeof(string));
            dt.Columns.Add("MB002", typeof(string));
            da.Fill(dt);

            if (dt.Rows.Count > 0)
            {
               textBox4.Text = dt.Rows[0]["MB002"].ToString();
            }
           
            sqlConn.Close();
        }
        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            CALMONEY();
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            CALMONEY();
        }

        public void CALMONEY()
        {
            if(!string.IsNullOrEmpty(textBox5.Text.ToString()) && !string.IsNullOrEmpty(textBox6.Text.ToString()))
            {
                if ((Convert.ToDouble(textBox5.Text.ToString()) > 0) && (Convert.ToDouble(textBox6.Text.ToString()) > 0))
                {
                    textBox7.Text = (Convert.ToDouble(textBox5.Text.ToString()) * Convert.ToDouble(textBox6.Text.ToString())).ToString();
                }
            }
          
        }
        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            SERACHCOPMA();
        }

        public void SERACHCOPMA()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@" SELECT MA001,MA002  ");
            Sequel.AppendFormat(@" FROM COPMA ");
            Sequel.AppendFormat(@" WHERE MA001='{0}'", textBox9.Text.ToString());
            Sequel.AppendFormat(@"  ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);

            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MA001", typeof(string));
            dt.Columns.Add("MA002", typeof(string));
            da.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                textBox10.Text = dt.Rows[0]["MA002"].ToString();
            }

            sqlConn.Close();
        }
        public void SETADD()
        {
            comboBox2.Enabled = true;
            comboBox3.Enabled = true;
            numericUpDown3.Enabled = true;
            numericUpDown4.Enabled = true;
            //textBox1.ReadOnly = false;
            //textBox2.ReadOnly = false;
            textBox3.ReadOnly = false;
            textBox4.ReadOnly = false;
            textBox5.ReadOnly = false;
            textBox6.ReadOnly = false;
            textBox7.ReadOnly = false;
            textBoxID.Text = null;
            textBox9.ReadOnly = false;
            //textBox10.ReadOnly = false;
        }
        public void SETADDNEW()
        {
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;
            textBoxID.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
        }

        public void SETUPDATE()
        {
            comboBox2.Enabled = true;
            comboBox3.Enabled = true;
            numericUpDown3.Enabled = true;
            numericUpDown4.Enabled = true;
            //textBox1.ReadOnly = false;
            //textBox2.ReadOnly = false;
            textBox3.ReadOnly = false;
            textBox4.ReadOnly = false;
            textBox5.ReadOnly = false;
            textBox6.ReadOnly = false;
            textBox7.ReadOnly = false;
            textBox9.ReadOnly = false;
            //textBox10.ReadOnly = false;
        }
        public void SETFINISH()
        {
            comboBox2.Enabled = false;
            comboBox3.Enabled = false;
            numericUpDown3.Enabled = false;
            numericUpDown4.Enabled = false;
            //textBox1.ReadOnly = true;
            //textBox2.ReadOnly = true;
            textBox3.ReadOnly = true;
            textBox4.ReadOnly = true;
            textBox5.ReadOnly = true;
            textBox6.ReadOnly = true;
            textBox7.ReadOnly = true;          
            textBox9.ReadOnly = true;
            //textBox10.ReadOnly = true;
        }

        public void UPDATE()
        {
            try
            {
               
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.Append(" UPDATE [PRESALE]  ");
                sbSql.AppendFormat("  SET [YEARS]='{1}',[MONTHS]='{2}',[DEPID]='{3}',[DEPNAME]='{4}',[SALESID]='{5}',[SALESNAME]='{6}',[CUSTOMERID]='{7}',[CUSTOMERNAME]='{8}',[PRODUCTID]='{9}',[PRODUCTNAME]='{10}',[PRICES]='{11}',[NUM]='{12}',[TMONEY]='{13}'  WHERE [ID]='{0}' ", textBoxID.Text.ToString(), numericUpDown3.Value.ToString(),  numericUpDown4.Value.ToString() , textBox1.Text.ToString(), comboBox2.Text.ToString(), textBox2.Text.ToString(), comboBox3.Text.ToString(), textBox9.Text.ToString(), textBox10.Text.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString());
                sbSql.Append("   ");

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

        public void ADD()
        {
            try
            {
                
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.Append(" INSERT INTO [PRESALE] ");
                sbSql.Append(" ([ID],[YEARS],[MONTHS],[DEPID],[DEPNAME],[SALESID],[SALESNAME],[CUSTOMERID],[CUSTOMERNAME],[PRODUCTID],[PRODUCTNAME],[PRICES],[NUM],[TMONEY]) ");
                sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}') ", Guid.NewGuid(),numericUpDown3.Value.ToString(),  numericUpDown4.Value.ToString().ToString(),textBox1.Text.ToString(),comboBox2.Text.ToString(),textBox2.Text.ToString(),comboBox3.Text.ToString(),textBox9.Text.ToString(), textBox10.Text.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString());

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

        public void DELETE()
        {
            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.Append(" DELETE [PRESALE]  ");
                sbSql.AppendFormat("  WHERE [ID]='{0}' ", textBoxID.Text.ToString());
                sbSql.Append("   ");

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
        #endregion

        #region BUTTON

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETADD();
            SETADDNEW();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SETUPDATE();
        }

        private void button4_Click(object sender, EventArgs e)
        {            
            DialogResult dialogResult = MessageBox.Show("確認要刪除嗎?", "yes/no", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETE();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
            Search();
            SETFINISH();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBoxID.Text.ToString()))
            {
                UPDATE();
            }
            else
            {
                ADD();
            }
            if (ds.Tables["TEMPds1"].Rows.Count >= 1)
            {
                rownum = dataGridView1.CurrentCell.RowIndex;
            }

            Search();
            SETFINISH();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }



        #endregion


    }
}
