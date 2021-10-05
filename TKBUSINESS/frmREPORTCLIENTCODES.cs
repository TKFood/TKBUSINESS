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
using System.Data.OleDb;

namespace TKBUSINESS
{
    public partial class frmREPORTCLIENTCODES : Form
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

        public frmREPORTCLIENTCODES()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void InsertExcelRecords(string FilePathName)
        {
            try
            {

                string conString = string.Empty;
                string extension = Path.GetExtension(FilePathName);
                switch (extension)
                {
                    case ".xls": //Excel 97-03
                        conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'";
                        break;
                    case ".xlsx": //Excel 07 or higher
                        conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'";
                        break;

                }

                conString = string.Format(conString, FilePathName);
                using (OleDbConnection excel_con = new OleDbConnection(conString))
                {
                    excel_con.Open();
                    string sheet1 = excel_con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["TABLE_NAME"].ToString();
                    DataTable dtExcelData = new DataTable();

                    //[OPTIONAL]: It is recommended as otherwise the data will be considered as String by default.
                    //dtExcelData.Columns.AddRange(new DataColumn[1] { new DataColumn("CODES", typeof(string)) });

                    using (OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [" + sheet1 + "]", excel_con))
                    {
                        oda.Fill(dtExcelData);
                    }
                    excel_con.Close();

                    //string consString = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
                    //using (SqlConnection con = new SqlConnection(consString))
                    //{
                    //    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                    //    {
                    //        //Set the database table name
                    //        sqlBulkCopy.DestinationTableName = "dbo.tblPersons";

                    //        //[OPTIONAL]: Map the Excel columns with that of the database table
                    //        sqlBulkCopy.ColumnMappings.Add("Id", "PersonId");
                    //        sqlBulkCopy.ColumnMappings.Add("Name", "Name");
                    //        sqlBulkCopy.ColumnMappings.Add("Salary", "Salary");
                    //        con.Open();
                    //        sqlBulkCopy.WriteToServer(dtExcelData);
                    //        con.Close();
                    //    }
                    //}
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Data has not been Imported due to :{0}", ex.Message), "Not Imported", MessageBoxButtons.OK, MessageBoxIcon.Warning);
           
            }


        
        }

        #endregion

        #region BUTTON
        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select file";
            dialog.InitialDirectory = ".\\";
            dialog.Filter = "xls files (*.*)|*.xls|xlsx files (*.*)|*.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                InsertExcelRecords(dialog.FileName);
                //MessageBox.Show(dialog.FileName);
            }
        }

        #endregion
    }
}
