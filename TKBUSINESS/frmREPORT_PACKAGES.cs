using FastReport;
using FastReport.Data;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.Extractor;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.POIFS;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using TKITDLL;

namespace TKBUSINESS
{
    public partial class frmREPORT_PACKAGES : Form
    {
        // 1. 定義表單狀態
        private enum EditState
        {
            Browse, // 瀏覽狀態
            Add,    // 新增狀態
            Edit    // 修改狀態
        }

        private EditState currentState = EditState.Browse;
        private DataTable dataTable = new DataTable();

        public frmREPORT_PACKAGES()
        {
            InitializeComponent();

            SwitchState(EditState.Browse); // 預設進入瀏覽模式
        }

        #region FUNCTION

        public void SEARCH_DG1(string SDATE)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [ID]
                                    ,[SHIPDATES]
                                    ,[SHIPFROM]
                                    ,[SHIPTO]
                                    ,[ITEM]
                                    ,[PRODUCTDESCRIPTION]
                                    ,[CARRIER]
                                    ,[PO]
                                    ,[SUPPLIER]
                                    ,[EXP]
                                    ,[LOT]
                                    ,[COO]
                                    ,[WEIGHT]
                                    ,[CUBE]
                                    ,[SELLUNIT]
                                    ,[ORDERUNITS]
                                    ,[LABEL]
                                    ,[BATCH]
                                    FROM [TKBUSINESS].[dbo].[REPORT_PACKAGES]
                                    WHERE [SHIPDATES]='{0}'
                                    ORDER BY  [ID]


                                    ", SDATE);
                SqlDataAdapter da = new SqlDataAdapter(@"" + sbSql, sqlConn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;


            }
            catch
            {

            }
            finally
            {

            }
        }

        private void SwitchState(EditState newState)
        {
            currentState = newState;

            switch (currentState)
            {
                case EditState.Browse:
                    // 瀏覽模式：允許點選與觸發編輯動作；禁止存檔取消
                    toolStripButton1.Enabled = true;
                    toolStripButton2.Enabled = dataGridView1.CurrentRow != null;
                    toolStripButton3.Enabled = dataGridView1.CurrentRow != null;
                    toolStripButton4.Enabled = false;
                   
                    dataGridView1.Enabled = true;
                    SetInputsReadOnly(true);                    

                    break;

                case EditState.Add:
                    // 新增模式：主鍵 (ID) 開放輸入
                    toolStripButton1.Enabled = false;
                    toolStripButton2.Enabled = false;
                    toolStripButton3.Enabled = false;
                    toolStripButton4.Enabled = true;
                   

                    dataGridView1.Enabled = false; // 鎖定 Grid，避免切換資料列
                    ClearInputs();
                    SetInputsReadOnly(false);
                    textBoxID.ReadOnly = false; // 主鍵可輸入
                    textBoxID.Focus();
                    break;

                case EditState.Edit:
                    // 修改模式：主鍵 (ID) 鎖定唯讀
                    toolStripButton1.Enabled = false;
                    toolStripButton2.Enabled = false;
                    toolStripButton3.Enabled = false;
                    toolStripButton4.Enabled = true;                   

                    dataGridView1.Enabled = false;
                    SetInputsReadOnly(false);
                    textBoxID.ReadOnly = true; // 主鍵不可修改
                    
                    break;
            }


            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");
            SEARCH_DG1(SDATES);
        }

        // 設定控制項唯讀狀態
        private void SetInputsReadOnly(bool readOnly)
        {
            textBox1.ReadOnly = readOnly;
            textBox2.ReadOnly = readOnly;
            textBox3.ReadOnly = readOnly;
            textBox4.ReadOnly = readOnly;
            textBox5.ReadOnly = readOnly;
            textBox6.ReadOnly = readOnly;
            textBox7.ReadOnly = readOnly;
            textBox8.ReadOnly = readOnly;
            textBox9.ReadOnly = readOnly;
            textBox10.ReadOnly = readOnly;
            textBox11.ReadOnly = readOnly;
            textBox12.ReadOnly = readOnly;
            textBox13.ReadOnly = readOnly;
            textBox14.ReadOnly = readOnly;
            textBox15.ReadOnly = readOnly;
            textBox16.ReadOnly = readOnly;

        }

        // 清空輸入欄位
        private void ClearInputs()
        {
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
            textBox7.Clear();
            textBox8.Clear();
            textBox9.Clear();
            textBox10.Clear();
            textBox11.Clear();
            textBox12.Clear();
            textBox13.Clear();
            textBox14.Clear();
            textBox15.Clear();
            textBox16.Clear();
           
        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");           
            SEARCH_DG1(SDATES);
        }

        private void btnTSNew_Click(object sender, EventArgs e)
        {
            SwitchState(EditState.Add);
        }

        // 【修改】按鈕
        private void btnTSEdit_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow == null) return;
            SwitchState(EditState.Edit);
        }

        // 【刪除】按鈕（即時二次確認後刪除）
        private void btnTSDelete_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow == null) return;

            string id = textBoxID.Text.Trim();

            if (MessageBox.Show($"確定要刪除代碼為 [{id}] 的資料嗎？", "刪除確認",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                // TODO: 執行 SQL DELETE ... 
                // ExecuteDeleteSQL(id);

                // 模擬資料表更新
                //DataRow[] targetRows = dataTable.Select($"ID = '{id}'");
                //foreach (var row in targetRows) row.Delete();

                MessageBox.Show("刪除成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                SwitchState(EditState.Browse);
            }
        }

        // 【存檔】按鈕（依狀態執行 INSERT 或 UPDATE）
        private void btnTSSave_Click(object sender, EventArgs e)
        {
            string id = textBoxID.Text.Trim();
            

            // 基礎資料驗證
            if (string.IsNullOrEmpty(id) )
            {
                MessageBox.Show("請填寫所有必要欄位！", "驗證錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                if (currentState == EditState.Add)
                {
                    // TODO: 執行 SQL INSERT ...
                    // ExecuteInsertSQL(id, name);

                    //dataTable.Rows.Add(id, name);
                    MessageBox.Show("新增成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else if (currentState == EditState.Edit)
                {
                    // TODO: 執行 SQL UPDATE ...
                    // ExecuteUpdateSQL(id, name);

                    //DataRow[] rows = dataTable.Select($"ID = '{id}'");
                    //if (rows.Length > 0) rows[0]["Name"] = name;

                    MessageBox.Show("修改成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                // 存檔成功後，恢復為瀏覽模式
                SwitchState(EditState.Browse);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"儲存發生錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        #endregion
    }
}
