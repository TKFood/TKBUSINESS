﻿using System;
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
    public partial class frmPRESALEV2018COPY : Form
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
        string SALSESID = null;
        int result;
        DataGridViewRow drPRESLAES = new DataGridViewRow();

        public frmPRESALEV2018COPY()
        {
            InitializeComponent();

            numericUpDown1.Value = (DateTime.Now.Year + 1);
            numericUpDown3.Value = (DateTime.Now.Year + 1);
        }

        #region FUNCTION
        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            SEARCHEMP();
        }

        public void SEARCHEMP()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
                    
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@" SELECT [Code],[CnName] FROM [HRMDB].[dbo].[Employee]   WHERE [Code]='{0}'", textBox7.Text.ToString());
            Sequel.AppendFormat(@"  ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);

            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("Code", typeof(string));
            dt.Columns.Add("CnName", typeof(string));
            da.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                textBox8.Text = dt.Rows[0]["CnName"].ToString();
            }
            else
            {
                textBox8.Text = null;
            }

            sqlConn.Close();
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            SEARCHCOPMA();
        }

        public void SEARCHCOPMA()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@" SELECT MA001,MA002  FROM [TK].dbo.COPMA   WHERE MA001='{0}'", textBox9.Text.ToString());
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
            else
            {
                textBox10.Text = null;
            }

            sqlConn.Close();
        }

        public void COPYPRESALE2018()
        {
            try
            {

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" INSERT INTO [TKBUSINESS].[dbo].[PRESALE2018]");
                sbSql.AppendFormat(" ([ID],[YEARS],[MONTHS],[SALESID],[SALESNAME],[CUSTOMERID],[CUSTOMERNAME],[MB001],[MB002],[PRICES],[NUM],[TMONEY],[MB003])");
                sbSql.AppendFormat(" SELECT NEWID(),'{0}','{1}',[SALESID],[SALESNAME],[CUSTOMERID],[CUSTOMERNAME],[MB001],[MB002],[PRICES],[NUM],[TMONEY],[MB003]", numericUpDown3.Value.ToString(), numericUpDown4.Value.ToString());
                sbSql.AppendFormat(" FROM [TKBUSINESS].[dbo].[PRESALE2018]");
                sbSql.AppendFormat(" WHERE [YEARS]='{0}' AND [MONTHS]='{1}' AND [SALESID]='{2}' AND [CUSTOMERID]='{3}'",numericUpDown1.Value.ToString(),numericUpDown2.Value.ToString(),textBox7.Text,textBox9.Text);
                sbSql.AppendFormat(" ");

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

        private void button1_Click(object sender, EventArgs e)
        {
            frmPRESALEV2018COPMA SUBfrmPRESALEV2018COPMA = new frmPRESALEV2018COPMA();
            SUBfrmPRESALEV2018COPMA.ShowDialog();
            textBox9.Text = SUBfrmPRESALEV2018COPMA.TextBoxMsg.Trim();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            COPYPRESALE2018();
            this.Close();
        }



        #endregion


    }
}
