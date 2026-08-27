using FastReport;
using FastReport.Data;
using FastReport.Preview;
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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            ClearInputs();

            if (dataGridView1.CurrentRow != null)
            {
                DataGridViewRow selectedRow = dataGridView1.CurrentRow;
                textBoxID.Text = selectedRow.Cells["ID"].Value?.ToString();
                dateTimePicker2.Value = DateTime.TryParseExact(selectedRow.Cells["SHIPDATES"].Value?.ToString(), "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out DateTime shipDate) ? shipDate : DateTime.Now;
                textBox1.Text = selectedRow.Cells["SHIPFROM"].Value?.ToString();
                textBox2.Text = selectedRow.Cells["SHIPTO"].Value?.ToString();
                textBox3.Text = selectedRow.Cells["ITEM"].Value?.ToString();
                textBox4.Text = selectedRow.Cells["PRODUCTDESCRIPTION"].Value?.ToString();
                textBox5.Text = selectedRow.Cells["CARRIER"].Value?.ToString();
                textBox6.Text = selectedRow.Cells["PO"].Value?.ToString();
                textBox7.Text = selectedRow.Cells["SUPPLIER"].Value?.ToString();
                textBox8.Text = selectedRow.Cells["EXP"].Value?.ToString();
                textBox9.Text = selectedRow.Cells["LOT"].Value?.ToString();
                textBox10.Text = selectedRow.Cells["COO"].Value?.ToString();
                textBox11.Text = selectedRow.Cells["WEIGHT"].Value?.ToString();
                textBox12.Text = selectedRow.Cells["CUBE"].Value?.ToString();
                textBox13.Text = selectedRow.Cells["SELLUNIT"].Value?.ToString();
                textBox14.Text = selectedRow.Cells["ORDERUNITS"].Value?.ToString();
                textBox15.Text = selectedRow.Cells["LABEL"].Value?.ToString();
                textBox16.Text = selectedRow.Cells["BATCH"].Value?.ToString();
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
                    toolStripButton2.Enabled = true;
                    toolStripButton3.Enabled = true;
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

        public void ADD_TOREPORT_PACKAGES(
            string SHIPDATES,
            string SHIPFROM,
            string SHIPTO,
            string ITEM,
            string PRODUCTDESCRIPTION,
            string CARRIER,
            string PO,
            string SUPPLIER,
            string EXP,
            string LOT,
            string COO,
            string WEIGHT,
            string CUBE,
            string SELLUNIT,
            string ORDERUNITS,
            string LABEL,
            string BATCH
            )
        {
            // 1. 處理連線字串與解密
            Class1 tkId = new Class1();
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
            sqlsb.Password = tkId.Decryption(sqlsb.Password);
            sqlsb.UserID = tkId.Decryption(sqlsb.UserID);

            // 2. 使用 using 確保資源自動釋放
            using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
            {
                conn.Open();
                // 開啟交易
                using (SqlTransaction trans = conn.BeginTransaction())
                {
                    try
                    {
                        string sql = @"
                                    INSERT INTO [TKBUSINESS].[dbo].[REPORT_PACKAGES]
                                    (
                                    [SHIPDATES]
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
                                    )
                                    VALUES
                                    (
                                    @SHIPDATES
                                    ,@SHIPFROM
                                    ,@SHIPTO
                                    ,@ITEM
                                    ,@PRODUCTDESCRIPTION
                                    ,@CARRIER
                                    ,@PO
                                    ,@SUPPLIER
                                    ,@EXP
                                    ,@LOT
                                    ,@COO
                                    ,@WEIGHT
                                    ,@CUBE
                                    ,@SELLUNIT
                                    ,@ORDERUNITS
                                    ,@LABEL
                                    ,@BATCH
                                    )
                                       
                                        ";

                        using (SqlCommand command = new SqlCommand(sql, conn, trans))
                        {
                            command.CommandTimeout = 60;
                            // 使用參數化查詢，避免 SQL Injection
                            command.Parameters.AddWithValue("@SHIPDATES", SHIPDATES ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@SHIPFROM", SHIPFROM ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@SHIPTO", SHIPTO ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@ITEM", ITEM ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@PRODUCTDESCRIPTION", PRODUCTDESCRIPTION ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@CARRIER", CARRIER ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@PO", PO ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@SUPPLIER", SUPPLIER ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@EXP", EXP ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@LOT", LOT ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@COO", COO ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@WEIGHT", WEIGHT ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@CUBE", CUBE ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@SELLUNIT", SELLUNIT ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@ORDERUNITS", ORDERUNITS ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@LABEL", LABEL ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@BATCH", BATCH ?? (object)DBNull.Value);
                            

                            int rowsAffected = command.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                trans.Commit();
                            }
                            else
                            {
                                // 若更新筆數為 0，通常代表 ID 不存在
                                trans.Rollback();
                                MessageBox.Show("更新失敗。");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // 發生錯誤時回滾
                        if (trans.Connection != null) trans.Rollback();
                        MessageBox.Show("系統錯誤：" + ex.Message);
                    }
                }
            }

        }

        public void UPDATE_TOREPORT_PACKAGES(
           string ID,
           string SHIPDATES,
           string SHIPFROM,
           string SHIPTO,
           string ITEM,
           string PRODUCTDESCRIPTION,
           string CARRIER,
           string PO,
           string SUPPLIER,
           string EXP,
           string LOT,
           string COO,
           string WEIGHT,
           string CUBE,
           string SELLUNIT,
           string ORDERUNITS,
           string LABEL,
           string BATCH
           )
        {
            // 1. 處理連線字串與解密
            Class1 tkId = new Class1();
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
            sqlsb.Password = tkId.Decryption(sqlsb.Password);
            sqlsb.UserID = tkId.Decryption(sqlsb.UserID);

            // 2. 使用 using 確保資源自動釋放
            using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
            {
                conn.Open();
                // 開啟交易
                using (SqlTransaction trans = conn.BeginTransaction())
                {
                    try
                    {
                        string sql = @"                                    
                                    UPDATE  [TKBUSINESS].[dbo].[REPORT_PACKAGES]
                                    SET
                                    [SHIPDATES]=@SHIPDATES
                                    ,[SHIPFROM]=@SHIPFROM
                                    ,[SHIPTO]=@SHIPTO
                                    ,[ITEM]=@ITEM
                                    ,[PRODUCTDESCRIPTION]=@PRODUCTDESCRIPTION
                                    ,[CARRIER]=@CARRIER
                                    ,[PO]=@PO
                                    ,[SUPPLIER]=@SUPPLIER
                                    ,[EXP]=@EXP
                                    ,[LOT]=@LOT
                                    ,[COO]=@COO
                                    ,[WEIGHT]=@WEIGHT
                                    ,[CUBE]=@CUBE
                                    ,[SELLUNIT]=@SELLUNIT
                                    ,[ORDERUNITS]=@ORDERUNITS
                                    ,[LABEL]=@LABEL
                                    ,[BATCH]=@BATCH
                                    WHERE [ID]=@ID                                       
                                        ";

                        using (SqlCommand command = new SqlCommand(sql, conn, trans))
                        {
                            command.CommandTimeout = 60;
                            // 使用參數化查詢，避免 SQL Injection
                            command.Parameters.AddWithValue("@ID", ID ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@SHIPDATES", SHIPDATES ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@SHIPFROM", SHIPFROM ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@SHIPTO", SHIPTO ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@ITEM", ITEM ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@PRODUCTDESCRIPTION", PRODUCTDESCRIPTION ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@CARRIER", CARRIER ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@PO", PO ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@SUPPLIER", SUPPLIER ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@EXP", EXP ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@LOT", LOT ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@COO", COO ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@WEIGHT", WEIGHT ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@CUBE", CUBE ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@SELLUNIT", SELLUNIT ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@ORDERUNITS", ORDERUNITS ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@LABEL", LABEL ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@BATCH", BATCH ?? (object)DBNull.Value);


                            int rowsAffected = command.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                trans.Commit();
                            }
                            else
                            {
                                // 若更新筆數為 0，通常代表 ID 不存在
                                trans.Rollback();
                                MessageBox.Show("更新失敗。");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // 發生錯誤時回滾
                        if (trans.Connection != null) trans.Rollback();
                        MessageBox.Show("系統錯誤：" + ex.Message);
                    }
                }
            }

        }

        public void DELETE_TOREPORT_PACKAGES(
          string ID
          )
        {
            // 1. 處理連線字串與解密
            Class1 tkId = new Class1();
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
            sqlsb.Password = tkId.Decryption(sqlsb.Password);
            sqlsb.UserID = tkId.Decryption(sqlsb.UserID);

            // 2. 使用 using 確保資源自動釋放
            using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
            {
                conn.Open();
                // 開啟交易
                using (SqlTransaction trans = conn.BeginTransaction())
                {
                    try
                    {
                        string sql = @"                                    
                                    DELETE  [TKBUSINESS].[dbo].[REPORT_PACKAGES]                                  
                                    WHERE [ID]=@ID                                       
                                        ";

                        using (SqlCommand command = new SqlCommand(sql, conn, trans))
                        {
                            command.CommandTimeout = 60;
                            // 使用參數化查詢，避免 SQL Injection
                            command.Parameters.AddWithValue("@ID", ID ?? (object)DBNull.Value);
                         
                            int rowsAffected = command.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                trans.Commit();
                            }
                            else
                            {
                                // 若更新筆數為 0，通常代表 ID 不存在
                                trans.Rollback();
                                MessageBox.Show("更新失敗。");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // 發生錯誤時回滾
                        if (trans.Connection != null) trans.Rollback();
                        MessageBox.Show("系統錯誤：" + ex.Message);
                    }
                }
            }

        }

        public void SETFASTREPORT(List<string> checkedRows)
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL1();

            Report report1 = new Report();
            report1.Load(@"REPORT\銷貨單憑証-直式.frx");

            //20210902密      
            Class1 TKID = new Class1();
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                           
                            ");

            return SB;

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
                DELETE_TOREPORT_PACKAGES(id);

                MessageBox.Show("刪除成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                SwitchState(EditState.Browse);
            }
        }

        // 【存檔】按鈕（依狀態執行 INSERT 或 UPDATE）
        private void btnTSSave_Click(object sender, EventArgs e)
        {
            string id = textBoxID.Text.Trim();
            string SHIPDATES = dateTimePicker2.Value.ToString("yyyyMMdd");
            string SHIPFROM = textBox1.Text.Trim();
            string SHIPTO = textBox2.Text.Trim();
            string ITEM = textBox3.Text.Trim();
            string PRODUCTDESCRIPTION = textBox4.Text.Trim();
            string CARRIER = textBox5.Text.Trim();
            string PO = textBox6.Text.Trim();
            string SUPPLIER = textBox7.Text.Trim();
            string EXP = textBox8.Text.Trim();
            string LOT = textBox9.Text.Trim();
            string COO = textBox10.Text.Trim();
            string WEIGHT = textBox11.Text.Trim();
            string CUBE = textBox12.Text.Trim();
            string SELLUNIT = textBox13.Text.Trim();
            string ORDERUNITS = textBox14.Text.Trim();
            string LABEL = textBox15.Text.Trim();
            string BATCH = textBox16.Text.Trim();

            // 基礎資料驗證
            if (string.IsNullOrEmpty(SHIPDATES) )
            {
                MessageBox.Show("請填寫所有必要欄位！", "驗證錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                if (currentState == EditState.Add)
                {
                    // TODO: 執行 SQL INSERT ...
                    ADD_TOREPORT_PACKAGES(
                        SHIPDATES,
                        SHIPFROM,
                        SHIPTO,
                        ITEM,
                        PRODUCTDESCRIPTION,
                        CARRIER,
                        PO,
                        SUPPLIER,
                        EXP,
                        LOT,
                        COO,
                        WEIGHT,
                        CUBE,
                        SELLUNIT,
                        ORDERUNITS,
                        LABEL,
                        BATCH
                        );
                    MessageBox.Show("新增成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else if (currentState == EditState.Edit)
                {
                    // TODO: 執行 SQL UPDATE ...
                    UPDATE_TOREPORT_PACKAGES(
                        id,
                        SHIPDATES,
                        SHIPFROM,
                        SHIPTO,
                        ITEM,
                        PRODUCTDESCRIPTION,
                        CARRIER,
                        PO,
                        SUPPLIER,
                        EXP,
                        LOT,
                        COO,
                        WEIGHT,
                        CUBE,
                        SELLUNIT,
                        ORDERUNITS,
                        LABEL,
                        BATCH
                        );

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
