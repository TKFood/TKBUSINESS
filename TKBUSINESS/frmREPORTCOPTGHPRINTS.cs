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

namespace TKBUSINESS
{
    public partial class frmREPORTCOPTGHPRINTS : Form
    {
        public frmREPORTCOPTGHPRINTS()
        {
            InitializeComponent();
        }

        private void frmREPORTCOPTGHPRINTS_Load(object sender, EventArgs e)
        {
            AddCheckBoxColumn(dataGridView1);
        }
        #region FUNCTION

        //在gridview新增checkbox欄位
        public void AddCheckBoxColumn(DataGridView dgv)
        {
            DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();
            chk.HeaderText = "選擇";
            chk.Name = "chk";
            chk.Width = 50;
            dgv.Columns.Insert(0, chk);
        }

        //記錄gridview被勾選的chekcbox
        public List<string> GetCheckedRows(DataGridView dgv)
        {
            List<string> checkedRows = new List<string>();
            foreach (DataGridViewRow row in dgv.Rows)
            {
                DataGridViewCheckBoxCell chkCell = row.Cells["chk"] as DataGridViewCheckBoxCell;
                if (chkCell != null && chkCell.Value != null && (bool)chkCell.Value)
                {
                    //假設你想要取得第一欄的值
                    string value = row.Cells[1].Value.ToString()+row.Cells[2].Value.ToString();
                    checkedRows.Add(value);
                }
            }
            return checkedRows;
        }

        public void SEARCH_DG1(string SDATE, string EDATE)
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
                                    TG001 AS '銷貨單別'
                                    ,TG002 AS '銷貨單號'
                                    ,MV002 AS '業務員'
                                    ,TG004 AS '客戶代號'
                                    ,TG007 AS '客戶名稱'

                                    FROM [TK].[dbo].[COPTG]
                                    INNER JOIN [TK].dbo.CMSMV ON MV001=TG006
                                    INNER JOIN [TK].dbo.COPMA ON MA001=TG004
                                    INNER JOIN [TK].dbo.CMSNA ON NA002=TG047
                                    WHERE  TG003>='{0}' AND TG003<='{1} '
                                    ORDER BY TG001,TG002

                                    ", SDATE, EDATE);
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
        public void SETFASTREPORT(List<string> checkedRows)
        {
            string SELECTED_ROWS = null;
            if (checkedRows!=null && checkedRows.Count>0)
            {
                SELECTED_ROWS = string.Join(",", checkedRows.Select(x => $"'{x}'"));
            }
            
           
            
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL1(SELECTED_ROWS);

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

        public StringBuilder SETSQL1(string SELECTED_ROWS)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                           --20260811 銷貨單憑証
                            SELECT 

                            TG010+MB002 AS '廠別'
                            ,TG001+MQ002 AS '銷貨單別'
                            ,TG002 AS '銷貨單號'
                            ,TG015 AS '統一編號'
                            --1.應稅內含、2.應稅外加、3.零稅率、4.免稅、9.不計稅
                            ,(CASE WHEN TG017='1' THEN '應稅內含' WHEN TG017='2' THEN '應稅外加'WHEN TG017='3' THEN '零稅率'WHEN TG017='4' THEN '免稅'WHEN TG017='9' THEN '不計稅' END  ) AS '課稅別'
                            ,TG044 AS '營業稅率'
                            ,TG005 AS '部門'
                            ,CONVERT(NVARCHAR,CONVERT(datetime,TG003),111) AS '單據日期'
                            --1.二聯式、2.三聯式、3.二聯式收銀機發票、4.三聯式收銀機發票、5.電子計算機發票、6.免用統一發票、A.增值稅專用發票、B.普通發票、C.免用發票   //890623 ADD 'A,B,C' BY 349 FOR 大陸用
                            ,(CASE WHEN TG016='1' THEN '二聯式'  WHEN TG016='2' THEN '三聯式' WHEN TG016='3' THEN '二聯式收銀機發票' WHEN TG016='4' THEN '三聯式收銀機發票' WHEN TG016='5' THEN '電子計算機發票' WHEN TG016='6' THEN '免用統一發票' WHEN TG016='7' THEN '電子發票' WHEN TG016='A' THEN '增值稅專用發票' WHEN TG016='B' THEN '普通發票' WHEN TG016='C' THEN '免用發票'  ELSE TG016 END ) AS '發票聯數'
                            ,MV002 AS '業務員'
                            ,TG032 AS '件數'
                            ,TG004 AS '客戶代號'
                            ,TG007 AS '客戶名稱'
                            ,CONVERT(NVARCHAR,CONVERT(datetime,TG021),111)  AS '發票日期'
                            ,TG014 AS '發票號碼'
                            ,TG012 AS '匯率'
                            ,TG011 AS '幣別'
                            ,MA008 AS '傳真'
                            ,TG106 AS '電話(一)'
                            ,TG107 AS '電話(二)'
                            ,TG008 AS '送貨地址'
                            ,TG047+NA003 AS '付款條件'
                            ,TG075 AS '收貨部門'
                            ,TG110 AS '指定日期'
                            --1.不分時段 2.早上(09~12點) 3.中午(12~17點) 4.下午(17~20點) 
                            --5.09點 6.10點 7.11點 8.12點 9.13點 A.14點 B.15點 C.16點 D.17點 E.18點 F.19點 G.20點  [DEF""1""]  //20101230 9.3 網購功能
                            ,(CASE WHEN TG111='1' THEN '不分時段' 
                            WHEN TG111='2' THEN '早上(09~12點)'
                            WHEN TG111='3' THEN '中午(12~17點)'
                            WHEN TG111='4' THEN '下午(17~20點)'
                            WHEN TG111='5' THEN '09點'
                            WHEN TG111='6' THEN '10點'
                            WHEN TG111='7' THEN '11點'
                            WHEN TG111='8' THEN '12點'
                            WHEN TG111='9' THEN '13點'
                            WHEN TG111='A' THEN '14點'
                            WHEN TG111='B' THEN '15點'
                            WHEN TG111='C' THEN '16點'
                            WHEN TG111='D' THEN '17點'
                            WHEN TG111='E' THEN '18點'
                            WHEN TG111='F' THEN '19點'
                            WHEN TG111='G' THEN '20點'
                            END) AS '配送時段'
                            ,TG113 AS '代收貨款'
                            ,TG027 AS '備註'
                            ,TH003 AS '序號'
                            ,TH004 AS '品號'
                            ,TH005 AS '品名'
                            ,TH006 AS '規格'
                            ,TH008 AS '銷貨數量'
                            --1.贈品量、2.備品量&&88-06-23
                            ,(CASE WHEN TH031='1' THEN '贈品量' WHEN TH031='2' THEN '備品量'  END ) AS '類型'
                            ,TH024 AS '贈/備品量'
                            ,TH009 AS '單位'
                            ,TH025 AS '折扣率%'
                            ,TH012 AS '單價'
                            ,TH013 AS '金額'
                            ,TH007 AS '銷貨庫別'
                            ,TH014+TH015 AS '訂單單號'
                            ,TH027+TH028+TH029 AS '結帳單號'
                            ,TH017 AS '批號'
                            ,TH026 AS '結帳碼'
                            ,TH018 AS '單身備註'
                            ,TG033 AS '數量合計'
                            ,TG013 AS '原幣銷貨金額'
                            ,TG025 AS '原幣銷貨稅額'
                            ,TG013+TG025 AS '原幣金額合計'
                            ,TG045 AS '本幣銷貨金額'
                            ,TG046 AS '本幣銷貨稅額'
                            ,TG045+TG046 AS '本幣金額合計'

                            FROM [TK].[dbo].[COPTH],[TK].[dbo].[COPTG]
                            INNER JOIN [TK].dbo.[CMSMB] ON MB001=TG010
                            INNER JOIN [TK].dbo.CMSMQ ON MQ001= TG001
                            INNER JOIN [TK].dbo.CMSMV ON MV001=TG006
                            INNER JOIN [TK].dbo.COPMA ON MA001=TG004
                            INNER JOIN [TK].dbo.CMSNA ON NA002=TG047
                            WHERE TG001=TH001 AND TG002=TH002
                            AND TG001+TG002 IN ({0})
                            ORDER BY TG001,TG002,TH003

                            ", SELECTED_ROWS);

            return SB;

        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            string SDATES=dateTimePicker1.Value.ToString("yyyyMMdd");
            string EDATES = dateTimePicker1.Value.ToString("yyyyMMdd");
            SEARCH_DG1(SDATES, EDATES);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var checkedRows = GetCheckedRows(dataGridView1);
            SETFASTREPORT(checkedRows);
        }
        #endregion

        
    }
}
