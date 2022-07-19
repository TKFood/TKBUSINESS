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
using System.Data.OleDb;

namespace TKBUSINESS
{
    public partial class frmADDCOPTCTD : Form
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

        string _path = null;

        public class DATACOPMD
        {
            public string COMPANY;
            public string CREATOR;
            public string USR_GROUP;
            public string CREATE_DATE;
            public string MODIFIER;
            public string MODI_DATE;
            public string FLAG;
            public string CREATE_TIME;
            public string MODI_TIME;
            public string TRANS_TYPE;
            public string TRANS_NAME;
            public string sync_date;
            public string sync_time;
            public string sync_mark;
            public string sync_count;
            public string DataUser;
            public string DataGroup;
            public string MD001;
            public string MD002;
            public string MD003;
            public string MD004;
            public string MD005;
            public string MD006;
            public string MD007;
            public string MD008;
            public string MD009;
            public string MD010;
            public string MD011;
            public string MD012;
            public string MD013;
            public string MD014;
            public string MD015;
            public string MD016;
            public string MD017;
            public string MD018;
            public string MD019;
            public string MD020;
            public string MD021;
            public string MD022;
            public string MD023;
            public string MD024;
            public string MD025;
            public string MD026;
            public string MD027;
            public string MD028;
            public string MD029;
            public string MD030;
            public string MD031;
            public string UDF01;
            public string UDF02;
            public string UDF03;
            public string UDF04;
            public string UDF05;
            public string UDF06;
            public string UDF07;
            public string UDF08;
            public string UDF09;
            public string UDF10;
        }

        public class DATACOPTC
        {

            public string COMPANY;
            public string CREATOR;
            public string USR_GROUP;
            public string CREATE_DATE;
            public string MODIFIER;
            public string MODI_DATE;
            public string FLAG;
            public string CREATE_TIME;
            public string MODI_TIME;
            public string TRANS_TYPE;
            public string TRANS_NAME;
            public string sync_date;
            public string sync_time;
            public string sync_mark;
            public string sync_count;
            public string DataUser;
            public string DataGroup;
            public string TC001;
            public string TC002;
            public string TC003;
            public string TC004;
            public string TC005;
            public string TC006;
            public string TC007;
            public string TC008;
            public string TC009;
            public string TC010;
            public string TC011;
            public string TC012;
            public string TC013;
            public string TC014;
            public string TC015;
            public string TC016;
            public string TC017;
            public string TC018;
            public string TC019;
            public string TC020;
            public string TC021;
            public string TC022;
            public string TC023;
            public string TC024;
            public string TC025;
            public string TC026;
            public string TC027;
            public string TC028;
            public string TC029;
            public string TC030;
            public string TC031;
            public string TC032;
            public string TC033;
            public string TC034;
            public string TC035;
            public string TC036;
            public string TC037;
            public string TC038;
            public string TC039;
            public string TC040;
            public string TC041;
            public string TC042;
            public string TC043;
            public string TC044;
            public string TC045;
            public string TC046;
            public string TC047;
            public string TC048;
            public string TC049;
            public string TC050;
            public string TC051;
            public string TC052;
            public string TC053;
            public string TC054;
            public string TC055;
            public string TC056;
            public string TC057;
            public string TC058;
            public string TC059;
            public string TC060;
            public string TC061;
            public string TC062;
            public string TC063;
            public string TC064;
            public string TC065;
            public string TC066;
            public string TC067;
            public string TC068;
            public string TC069;
            public string TC070;
            public string TC071;
            public string TC072;
            public string TC073;
            public string TC074;
            public string TC075;
            public string TC076;
            public string TC077;
            public string TC078;
            public string TC079;
            public string TC080;
            public string TC081;
            public string TC082;
            public string TC083;
            public string TC084;
            public string TC085;
            public string TC086;
            public string TC087;
            public string TC088;
            public string TC089;
            public string TC090;
            public string TC091;
            public string TC092;
            public string TC093;
            public string TC094;
            public string TC095;
            public string TC096;
            public string TC097;
            public string TC098;
            public string TC099;
            public string TC100;
            public string TC101;
            public string TC102;
            public string TC103;
            public string TC104;
            public string TC105;
            public string TC106;
            public string TC107;
            public string TC108;
            public string TC109;
            public string TC110;
            public string TC111;
            public string TC112;
            public string TC113;
            public string TC114;
            public string TC115;
            public string TC116;
            public string TC117;
            public string TC118;
            public string TC119;
            public string TC120;
            public string TC121;
            public string TC122;
            public string TC123;
            public string TC124;
            public string TC125;
            public string TC126;
            public string TC127;
            public string TC128;
            public string TC129;
            public string TC130;
            public string TC131;
            public string TC132;
            public string TC133;
            public string TC134;
            public string TC135;
            public string TC136;
            public string TC137;
            public string TC138;
            public string TC139;
            public string TC140;
            public string TC141;
            public string TC142;
            public string TC143;
            public string TC144;
            public string TC145;
            public string TC146;
            public string UDF01;
            public string UDF02;
            public string UDF03;
            public string UDF04;
            public string UDF05;
            public string UDF06;
            public string UDF07;
            public string UDF08;
            public string UDF09;
            public string UDF10;

        }
        public frmADDCOPTCTD()
        {
            InitializeComponent();

            SETTEXTBOX();
        }

        #region FUNCTION

        public void SETTEXTBOX()
        {
            textBox2.Text = "2221103200";
            textBox1.Text = SERACHCOPMA(textBox2.Text.ToString().Trim());
        }
        public void Search()
        {
            DataSet ds = new DataSet();

            try
            {
                sbSql.Clear();

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.AppendFormat(@"
                                       
                                        SELECT [SERNO]
                                    ,[預購單號]
                                    ,[收件者姓名]
                                    ,[電話(日)]
                                    ,[電話(夜)]
                                    ,[手機]
                                    ,[電子郵件]
                                    ,[預定到貨日期]
                                    ,[取貨時段]
                                    ,[郵遞區號]
                                    ,[縣市]
                                    ,[鄉鎮區]
                                    ,[收件者地址]
                                    ,[產品料號]
                                    ,[商品名稱]
                                    ,[預訂數量]
                                    ,[加油站代號]
                                    ,[加油站名]
                                    ,[TEMPCOPMAORDERRS].TC001
                                    ,[TEMPCOPMAORDERRS].TC002
                                    ,COPTC.TC001
                                    ,COPTC.TC002
                                    ,COPTD.TD001
                                    ,COPTD.TD002
                                    ,COPTD.TD003

                                    FROM [TKBUSINESS].[dbo].[TEMPCOPMAORDERRS]
                                    LEFT JOIN [TK].dbo.COPTC ON [預購單號]=TC012
                                    LEFT JOIN [TK].dbo.COPTD ON [預購單號]=TD086
                                    ORDER BY [SERNO]
                                         ");

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

                    //dataGridView1.CurrentCell = dataGridView1[0, 2];

                }



            }
            catch
            {

            }
            finally
            {

            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if(string.IsNullOrEmpty(textBox2.Text.ToString().Trim()))
            {
                textBox1.Text = SERACHCOPMA(textBox2.Text.ToString().Trim());
            }
            
        }

        public string SERACHCOPMA(string MA001)
        {
            DataSet ds = new DataSet();

            try
            {
                sbSql.Clear();

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.AppendFormat(@"
                                    SELECT MA001,MA002 
                                    FROM [TK].dbo.COPMA
                                    WHERE MA001='{0}'
                                         ",MA001);

                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 0)
                {
                    return ds.Tables["ds"].Rows[0]["MA002"].ToString().Trim();
                }
                else
                {
                    return null;
                }
              



            }
            catch
            {
                return null;
            }
            finally
            {

            }
        }


        public void ADDCOPMD(string MA001)
        {
            DATACOPMD COPMD = new DATACOPMD();
            COPMD.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
            COPMD.MD001 = MA001;

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

                sbSql.AppendFormat(@" 
                                    INSERT INTO [TK].[dbo].[COPMD]
                                    (
                                    [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]
                                    ,[MD001],[MD002],[MD003],[MD004],[MD005],[MD006],[MD007],[MD008],[MD009],[MD010]
                                    ,[MD011],[MD012],[MD013],[MD014],[MD015],[MD016],[MD017],[MD018],[MD019],[MD020]
                                    ,[MD021],[MD022],[MD023],[MD024],[MD025],[MD026],[MD027],[MD028],[MD029],[MD030]
                                    ,[MD031]
                                    ,[UDF01],[UDF02],[UDF03],[UDF04],[UDF05],[UDF06],[UDF07],[UDF08],[UDF09],[UDF10]
                                    )
                                    SELECT 
                                    'TK' AS [COMPANY],'160115' AS [CREATOR],'117000' AS [USR_GROUP],'{0}' AS [CREATE_DATE],'' AS [MODIFIER],'' AS [MODI_DATE],1 AS [FLAG],'' AS [CREATE_TIME],'' AS [MODI_TIME],'' AS [TRANS_TYPE],'' AS [TRANS_NAME],'' AS [sync_date],'' AS [sync_time],'' AS [sync_mark],0 AS [sync_count],'' AS [DataUser],'' AS [DataGroup]
                                    ,'{1}' AS [MD001],[加油站代號] AS [MD002],[縣市]+[鄉鎮區]+[收件者地址]+','+[收件者姓名]+','+[電話(日)]+'/ '+[手機] AS [MD003],'' AS [MD004],'' AS [MD005],[加油站名] AS [MD006],[收件者姓名] AS [MD007],'' AS [MD008],'' AS [MD009],'' AS [MD010]
                                    ,'' AS [MD011],[收件者姓名] AS [MD012],0 AS [MD013],0 AS [MD014],'' AS [MD015],'' AS [MD016],'' AS [MD017],[郵遞區號] AS [MD018],'' AS [MD019],'' AS [MD020]
                                    ,0 AS [MD021],'' AS [MD022],'' AS [MD023],'' AS [MD024],0 AS [MD025],'' AS [MD026],0 AS [MD027],'' AS [MD028],'' AS [MD029],'' AS [MD030]
                                    ,0 AS [MD031]
                                    ,'' AS [UDF01],'' AS [UDF02],'' AS [UDF03],'' AS [UDF04],'' AS [UDF05],0 AS [UDF06],0 AS [UDF07],0 AS [UDF08],0 AS [UDF09],0 AS [UDF10]
                                    FROM [TKBUSINESS].[dbo].[TEMPCOPMAORDERRS] 
                                    WHERE [加油站代號] NOT IN (SELECT [MD002] FROM [TK].[dbo].[COPMD] WHERE [MD001]='2221103200')
                                      
                                        ", COPMD.CREATE_DATE, COPMD.MD001);


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

                    MessageBox.Show("完成");
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

      

        public void ADDCOPTCCOPTD()
        {
            //DATACOPTC COPTC = new DATACOPTC();


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

                sbSql.AppendFormat(@" 
                                   --20220718 COPTC
                                    INSERT INTO [TK].[dbo].[COPTC]
                                    (
                                    [COMPANY]
                                    ,[CREATOR]
                                    ,[USR_GROUP]
                                    ,[CREATE_DATE]
                                    ,[MODIFIER]
                                    ,[MODI_DATE]
                                    ,[FLAG]
                                    ,[CREATE_TIME]
                                    ,[MODI_TIME]
                                    ,[TRANS_TYPE]
                                    ,[TRANS_NAME]
                                    ,[sync_date]
                                    ,[sync_time]
                                    ,[sync_mark]
                                    ,[sync_count]
                                    ,[DataUser]
                                    ,[DataGroup]
                                    ,[TC001]
                                    ,[TC002]
                                    ,[TC003]
                                    ,[TC004]
                                    ,[TC005]
                                    ,[TC006]
                                    ,[TC007]
                                    ,[TC008]
                                    ,[TC009]
                                    ,[TC010]
                                    ,[TC011]
                                    ,[TC012]
                                    ,[TC013]
                                    ,[TC014]
                                    ,[TC015]
                                    ,[TC016]
                                    ,[TC017]
                                    ,[TC018]
                                    ,[TC019]
                                    ,[TC020]
                                    ,[TC021]
                                    ,[TC022]
                                    ,[TC023]
                                    ,[TC024]
                                    ,[TC025]
                                    ,[TC026]
                                    ,[TC027]
                                    ,[TC028]
                                    ,[TC029]
                                    ,[TC030]
                                    ,[TC031]
                                    ,[TC032]
                                    ,[TC033]
                                    ,[TC034]
                                    ,[TC035]
                                    ,[TC036]
                                    ,[TC037]
                                    ,[TC038]
                                    ,[TC039]
                                    ,[TC040]
                                    ,[TC041]
                                    ,[TC042]
                                    ,[TC043]
                                    ,[TC044]
                                    ,[TC045]
                                    ,[TC046]
                                    ,[TC047]
                                    ,[TC048]
                                    ,[TC049]
                                    ,[TC050]
                                    ,[TC051]
                                    ,[TC052]
                                    ,[TC053]
                                    ,[TC054]
                                    ,[TC055]
                                    ,[TC056]
                                    ,[TC057]
                                    ,[TC058]
                                    ,[TC059]
                                    ,[TC060]
                                    ,[TC061]
                                    ,[TC062]
                                    ,[TC063]
                                    ,[TC064]
                                    ,[TC065]
                                    ,[TC066]
                                    ,[TC067]
                                    ,[TC068]
                                    ,[TC069]
                                    ,[TC070]
                                    ,[TC071]
                                    ,[TC072]
                                    ,[TC073]
                                    ,[TC074]
                                    ,[TC075]
                                    ,[TC076]
                                    ,[TC077]
                                    ,[TC078]
                                    ,[TC079]
                                    ,[TC080]
                                    ,[TC081]
                                    ,[TC082]
                                    ,[TC083]
                                    ,[TC084]
                                    ,[TC085]
                                    ,[TC086]
                                    ,[TC087]
                                    ,[TC088]
                                    ,[TC089]
                                    ,[TC090]
                                    ,[TC091]
                                    ,[TC092]
                                    ,[TC093]
                                    ,[TC094]
                                    ,[TC095]
                                    ,[TC096]
                                    ,[TC097]
                                    ,[TC098]
                                    ,[TC099]
                                    ,[TC100]
                                    ,[TC101]
                                    ,[TC102]
                                    ,[TC103]
                                    ,[TC104]
                                    ,[TC105]
                                    ,[TC106]
                                    ,[TC107]
                                    ,[TC108]
                                    ,[TC109]
                                    ,[TC110]
                                    ,[TC111]
                                    ,[TC112]
                                    ,[TC113]
                                    ,[TC114]
                                    ,[TC115]
                                    ,[TC116]
                                    ,[TC117]
                                    ,[TC118]
                                    ,[TC119]
                                    ,[TC120]
                                    ,[TC121]
                                    ,[TC122]
                                    ,[TC123]
                                    ,[TC124]
                                    ,[TC125]
                                    ,[TC126]
                                    ,[TC127]
                                    ,[TC128]
                                    ,[TC129]
                                    ,[TC130]
                                    ,[TC131]
                                    ,[TC132]
                                    ,[TC133]
                                    ,[TC134]
                                    ,[TC135]
                                    ,[TC136]
                                    ,[TC137]
                                    ,[TC138]
                                    ,[TC139]
                                    ,[TC140]
                                    ,[TC141]
                                    ,[TC142]
                                    ,[TC143]
                                    ,[TC144]
                                    ,[TC145]
                                    ,[TC146]
                                    ,[UDF01]
                                    ,[UDF02]
                                    ,[UDF03]
                                    ,[UDF04]
                                    ,[UDF05]
                                    ,[UDF06]
                                    ,[UDF07]
                                    ,[UDF08]
                                    ,[UDF09]
                                    ,[UDF10]
                                    )

                                    SELECT 
                                    'TK' AS [COMPANY],'160115' AS [CREATOR],'117000' AS [USR_GROUP],CONVERT(NVARCHAR,GETDATE(),112) AS [CREATE_DATE],'' AS [MODIFIER],'' AS [MODI_DATE],'1' AS [FLAG],CONVERT(NVARCHAR,GETDATE(),108) AS [CREATE_TIME],'' AS [MODI_TIME],'P001' AS [TRANS_TYPE],'COPMI06' AS [TRANS_NAME],'' AS [sync_date],'' AS [sync_time],'' AS [sync_mark],'0' AS [sync_count],'' AS [DataUser],'117000' AS [DataGroup]
                                    ,[TEMPCOPMAORDERRS].TC001 AS [TC001]
                                    ,[TEMPCOPMAORDERRS].TC002 AS [TC002]
                                    ,CONVERT(NVARCHAR,GETDATE(),112) AS [TC003]
                                    ,'2221103200' AS [TC004]
                                    ,'117100' AS [TC005]
                                    ,'160155' AS [TC006]
                                    ,'20' AS [TC007]
                                    ,'NTD' AS [TC008]
                                    ,1 AS [TC009]
                                    ,[縣市]+[鄉鎮區]+[收件者地址]+','+[收件者姓名]+','+[手機]+'/ '+[電話(日)] AS [TC010]
                                    ,'' AS [TC011]
                                    ,[預購單號] AS [TC012]
                                    ,'' AS [TC013]
                                    ,'出貨後30天匯款' AS [TC014]
                                    ,'' AS [TC015]
                                    ,'1' AS [TC016]
                                    ,'' AS [TC017]
                                    ,[收件者姓名] AS [TC018]
                                    ,'5' AS [TC019]
                                    ,'' AS [TC020]
                                    ,'' AS [TC021]
                                    ,'' AS [TC022]
                                    ,'' AS [TC023]
                                    ,'' AS [TC024]
                                    ,'' AS [TC025]
                                    ,0 AS [TC026]
                                    ,'N' AS [TC027]
                                    ,0 AS [TC028]
                                    ,(CASE WHEN COPMA.MA038='1' THEN ROUND((450*CONVERT(INT,[預訂數量])/1.05),0) WHEN COPMA.MA038='2' THEN ROUND(450*CONVERT(INT,[預訂數量]),0)  ELSE 450*CONVERT(INT,[預訂數量]) END) AS [TC029]
                                    , (CASE WHEN COPMA.MA038='1' THEN (450*CONVERT(INT,[預訂數量])-ROUND((450*CONVERT(INT,[預訂數量])/1.05),0)) WHEN COPMA.MA038='2' THEN ROUND((450*CONVERT(INT,[預訂數量])*0.05),0)  ELSE 0 END) AS [TC030]
                                    ,CONVERT(INT,[預訂數量]) AS [TC031]
                                    ,'2221103200' AS [TC032]
                                    ,'' AS [TC033]
                                    ,'' AS [TC034]
                                    ,'' AS [TC035]
                                    ,'' AS [TC036]
                                    ,'' AS [TC037]
                                    ,'' AS [TC038]
                                    ,CONVERT(NVARCHAR,GETDATE(),112) AS [TC039]
                                    ,'' AS [TC040]
                                    ,0.05 AS [TC041]
                                    ,'218' AS [TC042]
                                    ,0.980*CONVERT(INT,[預訂數量]) AS [TC043]
                                    ,0 AS [TC044]
                                    ,0 AS [TC045]
                                    ,0 AS [TC046]
                                    ,'' AS [TC047]
                                    ,'N' AS [TC048]
                                    ,'' AS [TC049]
                                    ,'N' AS [TC050]
                                    ,'' AS [TC051]
                                    ,0 AS [TC052]
                                    ,'台灣中油股份有限公司' AS [TC053]
                                    ,'' AS [TC054]
                                    ,[收件者姓名] AS [TC055]
                                    ,'' AS [TC056]
                                    ,'' AS [TC057]
                                    ,'' AS [TC058]
                                    ,'' AS [TC059]
                                    ,'' AS [TC060]
                                    ,'' AS [TC061]
                                    ,'' AS [TC062]
                                    ,'台北市信義區松仁路3號' AS [TC063]
                                    ,'' AS [TC064]
                                    ,'' AS [TC065]
                                    ,'' AS [TC066]
                                    ,'' AS [TC067]
                                    ,'' AS [TC068]
                                    ,'' AS [TC069]
                                    ,'N' AS [TC070]
                                    ,'' AS [TC071]
                                    ,0 AS [TC072]
                                    ,0 AS [TC073]
                                    ,0 AS [TC074]
                                    ,'' AS [TC075]
                                    ,'' AS [TC076]
                                    ,'' AS [TC077]
                                    ,'' AS [TC078]
                                    ,'' AS [TC079]
                                    ,'' AS [TC080]
                                    ,'' AS [TC081]
                                    ,'' AS [TC082]
                                    ,'' AS [TC083]
                                    ,'' AS [TC084]
                                    ,'' AS [TC085]
                                    ,'' AS [TC086]
                                    ,'' AS [TC087]
                                    ,'' AS [TC088]
                                    ,'' AS [TC089]
                                    ,'' AS [TC090]
                                    ,'' AS [TC091]
                                    ,'N' AS [TC092]
                                    ,'' AS [TC093]
                                    ,(CASE WHEN ISNULL([手機],'')<>'' THEN [手機] ELSE [電話(日)] END )AS [TC094]
                                    ,'' AS [TC095]
                                    ,'' AS [TC096]
                                    ,'' AS [TC097]
                                    ,'' AS [TC098]
                                    ,'1' AS [TC099]
                                    ,'N' AS [TC100]
                                    ,'' AS [TC101]
                                    ,'' AS [TC102]
                                    ,0 AS [TC103]
                                    ,'' AS [TC104]
                                    ,'' AS [TC105]
                                    ,'1' AS [TC106]
                                    ,0 AS [TC107]
                                    ,0 AS [TC108]
                                    ,0 AS [TC109]
                                    ,0 AS [TC110]
                                    ,0 AS [TC111]
                                    ,0 AS [TC112]
                                    ,'' AS [TC113]
                                    ,'' AS [TC114]
                                    ,'' AS [TC115]
                                    ,'1' AS [TC116]
                                    ,'' AS [TC117]
                                    ,0 AS [TC118]
                                    ,0 AS [TC119]
                                    ,0 AS [TC120]
                                    ,'7' AS [TC121]
                                    ,'' AS [TC122]
                                    ,'' AS [TC123]
                                    ,'03707901' AS [TC124]
                                    ,'' AS [TC125]
                                    ,'' AS [TC126]
                                    ,'' AS [TC127]
                                    ,'' AS [TC128]
                                    ,'' AS [TC129]
                                    ,'' AS [TC130]
                                    ,'' AS [TC131]
                                    ,'' AS [TC132]
                                    ,'' AS [TC133]
                                    ,'' AS [TC134]
                                    ,'' AS [TC135]
                                    ,'' AS [TC136]
                                    ,'' AS [TC137]
                                    ,'' AS [TC138]
                                    ,'' AS [TC139]
                                    ,'' AS [TC140]
                                    ,0 AS [TC141]
                                    ,'' AS [TC142]
                                    ,'' AS [TC143]
                                    ,'' AS [TC144]
                                    ,'' AS [TC145]
                                    ,'' AS [TC146]
                                    ,'' AS [UDF01]
                                    ,'' AS [UDF02]
                                    ,'' AS [UDF03]
                                    ,'' AS [UDF04]
                                    ,'' AS [UDF05]
                                    ,0 AS [UDF06]
                                    ,0 AS [UDF07]
                                    ,0 AS [UDF08]
                                    ,0 AS [UDF09]
                                    ,0 AS [UDF10]
                                    FROM [TKBUSINESS].[dbo].[TEMPCOPMAORDERRS]
                                    LEFT JOIN [TK].dbo.COPMA ON MA001='2221103200'
                                    WHERE [預購單號] NOT IN (SELECT TC012 FROM [TK].dbo.COPTC WHERE ISNULL(TC012,'')<>'')
                                      
                                    --20220719 COPTD
                                    INSERT INTO [TK].[dbo].[COPTD]
                                    (
                                    [COMPANY]
                                    ,[CREATOR]
                                    ,[USR_GROUP]
                                    ,[CREATE_DATE]
                                    ,[MODIFIER]
                                    ,[MODI_DATE]
                                    ,[FLAG]
                                    ,[CREATE_TIME]
                                    ,[MODI_TIME]
                                    ,[TRANS_TYPE]
                                    ,[TRANS_NAME]
                                    ,[sync_date]
                                    ,[sync_time]
                                    ,[sync_mark]
                                    ,[sync_count]
                                    ,[DataUser]
                                    ,[DataGroup]
                                    ,[TD001]
                                    ,[TD002]
                                    ,[TD003]
                                    ,[TD004]
                                    ,[TD005]
                                    ,[TD006]
                                    ,[TD007]
                                    ,[TD008]
                                    ,[TD009]
                                    ,[TD010]
                                    ,[TD011]
                                    ,[TD012]
                                    ,[TD013]
                                    ,[TD014]
                                    ,[TD015]
                                    ,[TD016]
                                    ,[TD017]
                                    ,[TD018]
                                    ,[TD019]
                                    ,[TD020]
                                    ,[TD021]
                                    ,[TD022]
                                    ,[TD023]
                                    ,[TD024]
                                    ,[TD025]
                                    ,[TD026]
                                    ,[TD027]
                                    ,[TD028]
                                    ,[TD029]
                                    ,[TD030]
                                    ,[TD031]
                                    ,[TD032]
                                    ,[TD033]
                                    ,[TD034]
                                    ,[TD035]
                                    ,[TD036]
                                    ,[TD037]
                                    ,[TD038]
                                    ,[TD039]
                                    ,[TD040]
                                    ,[TD041]
                                    ,[TD042]
                                    ,[TD043]
                                    ,[TD044]
                                    ,[TD045]
                                    ,[TD046]
                                    ,[TD047]
                                    ,[TD048]
                                    ,[TD049]
                                    ,[TD050]
                                    ,[TD051]
                                    ,[TD052]
                                    ,[TD053]
                                    ,[TD054]
                                    ,[TD055]
                                    ,[TD056]
                                    ,[TD057]
                                    ,[TD058]
                                    ,[TD059]
                                    ,[TD060]
                                    ,[TD061]
                                    ,[TD062]
                                    ,[TD063]
                                    ,[TD064]
                                    ,[TD065]
                                    ,[TD066]
                                    ,[TD067]
                                    ,[TD068]
                                    ,[TD069]
                                    ,[TD070]
                                    ,[TD071]
                                    ,[TD072]
                                    ,[TD073]
                                    ,[TD074]
                                    ,[TD075]
                                    ,[TD076]
                                    ,[TD077]
                                    ,[TD078]
                                    ,[TD079]
                                    ,[TD080]
                                    ,[TD081]
                                    ,[TD082]
                                    ,[TD083]
                                    ,[TD084]
                                    ,[TD085]
                                    ,[TD086]
                                    ,[TD087]
                                    ,[TD088]
                                    ,[TD089]
                                    ,[TD090]
                                    ,[TD091]
                                    ,[TD092]
                                    ,[TD093]
                                    ,[TD094]
                                    ,[TD095]
                                    ,[TD096]
                                    ,[TD097]
                                    ,[TD098]
                                    ,[TD099]
                                    ,[TD100]
                                    ,[TD101]
                                    ,[TD102]
                                    ,[TD103]
                                    ,[TD104]
                                    ,[TD105]
                                    ,[TD106]
                                    ,[TD107]
                                    ,[TD108]
                                    ,[TD109]
                                    ,[TD110]
                                    ,[TD111]
                                    ,[TD112]
                                    ,[TD113]
                                    ,[UDF01]
                                    ,[UDF02]
                                    ,[UDF03]
                                    ,[UDF04]
                                    ,[UDF05]
                                    ,[UDF06]
                                    ,[UDF07]
                                    ,[UDF08]
                                    ,[UDF09]
                                    ,[UDF10]
                                    ,[TD200]
                                    )

                                    SELECT 
                                    'TK' AS [COMPANY],'160115' AS [CREATOR],'117000' AS [USR_GROUP],CONVERT(NVARCHAR,GETDATE(),112) AS [CREATE_DATE],'' AS [MODIFIER],'' AS [MODI_DATE],'1' AS [FLAG],CONVERT(NVARCHAR,GETDATE(),108) AS [CREATE_TIME],'' AS [MODI_TIME],'P001' AS [TRANS_TYPE],'COPMI06' AS [TRANS_NAME],'' AS [sync_date],'' AS [sync_time],'' AS [sync_mark],0 AS [sync_count],'' AS [DataUser],'117000' AS [DataGroup]
                                    ,[TEMPCOPMAORDERRS].TC001 AS [TD001]
                                    ,[TEMPCOPMAORDERRS].TC002 AS [TD002]
                                    ,'0001' AS [TD003]
                                    ,MB001 AS [TD004]
                                    ,MB002 AS [TD005]
                                    ,MB003 AS [TD006]
                                    ,'20001' AS [TD007]
                                    ,CONVERT(INT,[預訂數量]) AS [TD008]
                                    ,0 AS [TD009]
                                    ,MB004 AS [TD010]
                                    ,450 AS [TD011]
                                    ,450*CONVERT(INT,[預訂數量]) AS [TD012]
                                    ,CONVERT(NVARCHAR,CONVERT(DATETIME,[預定到貨日期]),112) AS [TD013]
                                    ,'' AS [TD014]
                                    ,'' AS [TD015]
                                    ,'N' AS [TD016]
                                    ,'' AS [TD017]
                                    ,'' AS [TD018]
                                    ,'' AS [TD019]
                                    ,'' AS [TD020]
                                    ,'N' AS [TD021]
                                    ,0 AS [TD022]
                                    ,'' AS [TD023]
                                    ,0 AS [TD024]
                                    ,0 AS [TD025]
                                    ,1 AS [TD026]
                                    ,'' AS [TD027]
                                    ,'' AS [TD028]
                                    ,'' AS [TD029]
                                    ,5.200 AS [TD030]
                                    ,0 AS [TD031]
                                    ,0 AS [TD032]
                                    ,0 AS [TD033]
                                    ,0 AS [TD034]
                                    ,0 AS [TD035]
                                    ,'' AS [TD036]
                                    ,'' AS [TD037]
                                    ,'' AS [TD038]
                                    ,'' AS [TD039]
                                    ,'' AS [TD040]
                                    ,'' AS [TD041]
                                    ,0 AS [TD042]
                                    ,'' AS [TD043]
                                    ,'' AS [TD044]
                                    ,'9' AS [TD045]
                                    ,'' AS [TD046]
                                    ,'' AS [TD047]
                                    ,'' AS [TD048]
                                    ,'1' AS [TD049]
                                    ,0 AS [TD050]
                                    ,0 AS [TD051]
                                    ,0 AS [TD052]
                                    ,0 AS [TD053]
                                    ,0 AS [TD054]
                                    ,0 AS [TD055]
                                    ,'' AS [TD056]
                                    ,'' AS [TD057]
                                    ,'' AS [TD058]
                                    ,0 AS [TD059]
                                    ,'1' AS [TD060]
                                    ,0 AS [TD061]
                                    ,'N' AS [TD062]
                                    ,'' AS [TD063]
                                    ,'' AS [TD064]
                                    ,'' AS [TD065]
                                    ,'' AS [TD066]
                                    ,'' AS [TD067]
                                    ,'' AS [TD068]
                                    ,'N' AS [TD069]
                                    ,0 AS [TD070]
                                    ,'' AS [TD071]
                                    ,'' AS [TD072]
                                    ,'' AS [TD073]
                                    ,'' AS [TD074]
                                    ,'' AS [TD075]
                                    ,0 AS [TD076]
                                    ,'' AS [TD077]
                                    ,'' AS [TD078]
                                    ,'N' AS [TD079]
                                    ,'1' AS [TD080]
                                    ,0 AS [TD081]
                                    ,'' AS [TD082]
                                    ,'' AS [TD083]
                                    ,'' AS [TD084]
                                    ,0 AS [TD085]
                                    ,[預購單號] AS [TD086]
                                    ,0 AS [TD087]
                                    ,'' AS [TD088]
                                    ,'3' AS [TD089]
                                    ,0 AS [TD090]
                                    ,0 AS [TD091]
                                    ,'' AS [TD092]
                                    ,0 AS [TD093]
                                    ,'' AS [TD094]
                                    ,0 AS [TD095]
                                    ,'1' AS [TD096]
                                    ,'' AS [TD097]
                                    ,0 AS [TD098]
                                    ,0 AS [TD099]
                                    ,0 AS [TD100]
                                    ,'' AS [TD101]
                                    ,'' AS [TD102]
                                    ,'' AS [TD103]
                                    ,'' AS [TD104]
                                    ,'' AS [TD105]
                                    ,'' AS [TD106]
                                    ,0 AS [TD107]
                                    ,0 AS [TD108]
                                    ,0 AS [TD109]
                                    ,0 AS [TD110]
                                    ,0 AS [TD111]
                                    ,0 AS [TD112]
                                    ,0 AS [TD113]
                                    ,'' AS [UDF01]
                                    ,'' AS [UDF02]
                                    ,'' AS [UDF03]
                                    ,'' AS [UDF04]
                                    ,'' AS [UDF05]
                                    ,0 AS [UDF06]
                                    ,0 AS [UDF07]
                                    ,0 AS [UDF08]
                                    ,0 AS [UDF09]
                                    ,0 AS [UDF10]
                                    ,'' AS [TD200]
                                    FROM [TKBUSINESS].[dbo].[TEMPCOPMAORDERRS]
                                    LEFT JOIN [TK].dbo.COPMA ON MA001='2221103200'
                                    LEFT JOIN [TK].dbo.INVMB ON MB001='40117031018000'
                                    WHERE [預購單號] NOT IN (SELECT [TD086] FROM [TK].dbo.COPTD WHERE ISNULL([TD086],'')<>'')    
                                        ");


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

                    MessageBox.Show("完成");
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

        public string GETMAXTC002(string TC001, string TC003)
        {
            SqlDataAdapter adapter4 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
            DataSet ds4 = new DataSet();
            string TC002 = null;

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


                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TC002),'00000000000') AS TC002
                                       FROM [TK].[dbo].[COPTC] 
                                       WHERE  TC001='{0}' AND TC003='{1}'
                                    ",TC001,TC003);

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        TC002 = SETTC002(ds4.Tables["TEMPds4"].Rows[0]["TC002"].ToString());
                        return TC002;

                    }
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public string SETTC002(string TC002)
        {
            if (TC002.Equals("00000000000"))
            {
                return DateTime.Now.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(TC002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return DateTime.Now.ToString("yyyyMMdd") + temp.ToString();
            }
        }


        public void UPDATETEMPCOPMAORDERRS()
        {
            DataSet ds = new DataSet();

            try
            {
                sbSql.Clear();

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.AppendFormat(@"                                       
                                    SELECT [預購單號]
                                    FROM [TKBUSINESS].[dbo].[TEMPCOPMAORDERRS]
                                    WHERE [預購單號] NOT IN (SELECT TC012 FROM [TK].dbo.COPTC WHERE ISNULL(TC012,'')<>'')

                                         ");

                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 0)
                {
                    //多筆預購訂單，每陶1次新增TC001、TC002，避免重覆
                    string TC001 = "A221";
                    string TC002 = GETMAXTC002(TC001, DateTime.Now.ToString("yyyyMMdd"));

                    int serno = Convert.ToInt16(TC002.Substring(8, 3));
                    serno = serno - 1;

                    foreach (DataRow DR in ds.Tables["ds"].Rows)
                    {
                        string 預購單號 = DR["預購單號"].ToString();

                        //流水號+1
                        serno = serno + 1;
                        string temp = serno.ToString();
                        temp = temp.PadLeft(3, '0');
                        TC002=DateTime.Now.ToString("yyyyMMdd") + temp.ToString();

                        UPDATETEMPCOPMAORDERRSTC001TC002(預購單號, TC001, TC002);

                       
                    }

                }
               



            }
            catch
            {

            }
            finally
            {

            }
        }

        public void UPDATETEMPCOPMAORDERRSTC001TC002(string 預購單號,string TC001,string TC002)
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

                sbSql.AppendFormat(@" 
                                    UPDATE [TKBUSINESS].[dbo].[TEMPCOPMAORDERRS]
                                    SET TC001='{1}',TC002='{2}'
                                    WHERE  [預購單號]='{0}'
                                        ", 預購單號, TC001, TC002);


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

        public void OPENFILE(string MAXSERNO)
        {
            //記錄選到的檔案路徑
            _path = null;

            OpenFileDialog od = new OpenFileDialog();
            od.Filter = "Excell|*.xls;*.xlsx;";
           
            DialogResult dr = od.ShowDialog();
            if (dr == DialogResult.Abort)
            {
                return;
            }               
            if (dr == DialogResult.Cancel)
            {
                return;
            }
                
            textBox3.Text = od.FileName.ToString();
            _path = od.FileName.ToString();

            //匯入excel到db
            InsertExcelRecords(MAXSERNO);
        }


        private void InsertExcelRecords(string MAXSERNO)
        {

            try
            {
                //  ExcelConn(_path);
                //找出不同excel的格式，設定連接字串
                //xls跟非xls
                string constr = null;
                if (_path.CompareTo(".xls") == 0)
                {
                    constr = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _path + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
                }                    
                else
                {
                    constr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _path + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
                }
                    
                //找出excel的第1張分頁名稱，用query中                
                OleDbConnection Econ = new OleDbConnection(constr);
                Econ.Open();

              

                DataTable excelShema = Econ.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string firstSheetName = excelShema.Rows[0]["TABLE_NAME"].ToString();

                string Query = string.Format("Select * FROM [{0}]", firstSheetName);
                OleDbCommand Ecom = new OleDbCommand(Query, Econ);


                DataTable dtExcelData = new DataTable();
                
                OleDbDataAdapter oda = new OleDbDataAdapter(Query, Econ);
                Econ.Close();
                oda.Fill(dtExcelData);
                DataTable Exceldt = dtExcelData;

                //把第一列的欄位名移除
                Exceldt.Rows[0].Delete();
                //再去除特定條件
                for (int i = Exceldt.Rows.Count - 1; i >=1; i--)
                {
                    if (Convert.ToInt32(Exceldt.Rows[i][0].ToString())<=Convert.ToInt32(MAXSERNO))
                    {
                        Exceldt.Rows[i].Delete();
                    }
                }

                Exceldt.AcceptChanges();

                if(Exceldt.Rows.Count>0)
                {
                    //把Exceldt匯到db中
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    sbSql.Clear();

                    foreach (DataRow DR in Exceldt.Rows)
                    {
                        sbSql.AppendFormat(@" 
                                  
                                    INSERT INTO [TKBUSINESS].[dbo].[TEMPCOPMAORDERRS]
                                    (
                                    [SERNO]
                                    ,[預購單號]
                                    ,[收件者姓名]
                                    ,[電話(日)]
                                    ,[電話(夜)]
                                    ,[手機]
                                    ,[電子郵件]
                                    ,[預定到貨日期]
                                    ,[取貨時段]
                                    ,[郵遞區號]
                                    ,[縣市]
                                    ,[鄉鎮區]
                                    ,[收件者地址]
                                    ,[產品料號]
                                    ,[商品名稱]
                                    ,[預訂數量]
                                    ,[加油站代號]
                                    ,[加油站名]

                                    )
                                    VALUES
                                    (
                                    '{0}'
                                    ,'{1}'
                                    ,'{2}'
                                    ,'{3}'
                                    ,'{4}'
                                    ,'{5}'
                                    ,'{6}'
                                    ,'{7}'
                                    ,'{8}'
                                    ,'{9}'
                                    ,'{10}'
                                    ,'{11}'
                                    ,'{12}'
                                    ,'{13}'
                                    ,'{14}'
                                    ,'{15}'
                                    ,'{16}'
                                    ,'{17}'
                                    )
                                    ", DR[0].ToString()
                                        , DR[1].ToString()
                                        , DR[2].ToString()
                                        , DR[3].ToString()
                                        , DR[4].ToString()
                                        , DR[5].ToString()
                                        , DR[6].ToString()
                                        , DR[7].ToString()
                                        , DR[8].ToString()
                                        , DR[9].ToString()
                                        , DR[10].ToString()
                                        , DR[11].ToString()
                                        , DR[12].ToString()
                                        , DR[13].ToString()
                                        , DR[14].ToString()
                                        , DR[15].ToString()
                                        , DR[16].ToString()
                                        , DR[17].ToString()


                                        );

                    }


                    sqlConn.Open();
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

                        sqlConn.Close();
                    }

                    //SqlBulkCopy objbulk = new SqlBulkCopy(sqlConn);
                    ////assigning Destination table name
                    //objbulk.DestinationTableName = "dbo.TEMPCOPMAORDERRS";
                    ////Mapping Table column
                    //objbulk.ColumnMappings.Add("F1"," SERNO");
                    //objbulk.ColumnMappings.Add("F2", "預購單號");

                    //sqlConn.Open();
                    //objbulk.WriteToServer(Exceldt);
                    //sqlConn.Close();
                    //MessageBox.Show("Data has been Imported successfully.", "Imported", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("錯誤:{0}", ex.Message), "Not Imported", MessageBoxButtons.OK, MessageBoxIcon.Warning);
             
            }
        }


        public void TEST()
        {
            //記錄選到的檔案路徑
            _path = null;

            OpenFileDialog od = new OpenFileDialog();
            od.Filter = "Excell|*.xls;*.xlsx;";

            DialogResult dr = od.ShowDialog();
            if (dr == DialogResult.Abort)
            {
                return;
            }
            if (dr == DialogResult.Cancel)
            {
                return;
            }

            textBox3.Text = od.FileName.ToString();
            _path = od.FileName.ToString();

            string conString = string.Empty;
         

            if (_path.CompareTo(".xls") == 0)
            {
                conString = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _path + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            }
            else
            {
                conString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _path + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            }
            
            conString = string.Format(conString, _path);
            using (OleDbConnection excel_con = new OleDbConnection(conString))
            {
                excel_con.Open();
                string sheet1 = excel_con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["TABLE_NAME"].ToString();
                DataTable dtExcelData = new DataTable();

                //[OPTIONAL]: It is recommended as otherwise the data will be considered as String by default.
                dtExcelData.Columns.AddRange(
                    new DataColumn[3] { new DataColumn("Id", typeof(int)),
                    new DataColumn("Name", typeof(string)),
                    new DataColumn("Salary", typeof(decimal))

                    });

                using (OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [" + sheet1 + "]", excel_con))
                {
                    oda.Fill(dtExcelData);
                }
                excel_con.Close();

                string consString = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
                using (SqlConnection con = new SqlConnection(consString))
                {
                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                    {
                        //Set the database table name
                        sqlBulkCopy.DestinationTableName = "dbo.tblPersons";

                        //[OPTIONAL]: Map the Excel columns with that of the database table
                        sqlBulkCopy.ColumnMappings.Add("Id", "PersonId");
                        sqlBulkCopy.ColumnMappings.Add("Name", "Name");
                        sqlBulkCopy.ColumnMappings.Add("Salary", "Salary");
                        con.Open();
                        sqlBulkCopy.WriteToServer(dtExcelData);
                        con.Close();
                    }
                }
            }
        }

        public string FINDMAXSERNO()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();
           

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


                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT MAX([SERNO])  [SERNO] FROM [TKBUSINESS].[dbo].[TEMPCOPMAORDERRS]
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 1)
                {                    
                    return ds.Tables["ds"].Rows[0]["SERNO"].ToString();
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                return null;
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
            Search();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            ADDCOPMD(textBox2.Text.ToString().Trim());
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //先在[TKBUSINESS].[dbo].[TEMPCOPMAORDERRS] 指定好TC001、TC002，方便做ERP訂單的新增
            //用TC012的客戶單號=預購單號做比較，找出那些還沒有轉入ERP的訂單中
            UPDATETEMPCOPMAORDERRS();
            ADDCOPTCCOPTD();

            Search();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            textBox5.Text = FINDMAXSERNO();

            OPENFILE(textBox5.Text);
            //TEST();

            textBox5.Text = FINDMAXSERNO();
            Search();
        }

        #endregion


    }
}
