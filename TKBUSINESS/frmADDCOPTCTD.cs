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
                                        ,TC001
                                        ,TC002

                                        FROM [TKBUSINESS].[dbo].[TEMPCOPMAORDERRS]
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
            DATACOPTC COPTC = new DATACOPTC();

            COPTC.COMPANY = "TK";
            COPTC.CREATOR = "160115";
            COPTC.USR_GROUP = "117000";
            COPTC.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
            COPTC.MODIFIER = "160115";
            COPTC.MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            COPTC.FLAG = "1";
            COPTC.CREATE_TIME = DateTime.Now.ToString("HH:mm:ss");
            COPTC.MODI_TIME = DateTime.Now.ToString("HH:mm:ss");
            COPTC.TRANS_TYPE = "P001";
            COPTC.TRANS_NAME = "COPMI06";
            COPTC.sync_date = "";
            COPTC.sync_time = "";
            COPTC.sync_mark = "";
            COPTC.sync_count = "0";
            COPTC.DataUser = "";
            COPTC.DataGroup = "117000";
            COPTC.TC001 = "A221";
            COPTC.TC002 = GETMAXTC002(COPTC.TC001, DateTime.Now.ToString("yyyyMMdd"));
            COPTC.TC003 = DateTime.Now.ToString("yyyyMMdd");
            COPTC.TC004 = "2221103200";
            COPTC.TC005 = "117100";
            COPTC.TC006 = "160155";
            COPTC.TC007 = "20";
            COPTC.TC008 = "NTD";
            COPTC.TC009 = "1";
            COPTC.TC010 = "";
            COPTC.TC011 = "";
            COPTC.TC012 = "";
            COPTC.TC013 = "";
            COPTC.TC014 = "出貨後30天匯款";
            COPTC.TC015 = "1";
            COPTC.TC016 = "";
            COPTC.TC017 = "";
            COPTC.TC018 = "";
            COPTC.TC019 = "5";
            COPTC.TC020 = "";
            COPTC.TC021 = "";
            COPTC.TC022 = "";
            COPTC.TC023 = "";
            COPTC.TC024 = "";
            COPTC.TC025 = "";
            COPTC.TC026 = "0";
            COPTC.TC027 = "N";
            COPTC.TC028 = "0";
            COPTC.TC029 = "0";
            COPTC.TC030 = "0";
            COPTC.TC031 = "0";
            COPTC.TC032 = "2221103200";
            COPTC.TC033 = "";
            COPTC.TC034 = "";
            COPTC.TC035 = "";
            COPTC.TC036 = "";
            COPTC.TC037 = "";
            COPTC.TC038 = "";
            COPTC.TC039 = DateTime.Now.ToString("yyyyMMdd");
            COPTC.TC040 = "";
            COPTC.TC041 = "0.0500";
            COPTC.TC042 = "218";
            COPTC.TC043 = "0.980";
            COPTC.TC044 = "0";
            COPTC.TC045 = "0";
            COPTC.TC046 = "0";
            COPTC.TC047 = "";
            COPTC.TC048 = "N";
            COPTC.TC049 = "";
            COPTC.TC050 = "N";
            COPTC.TC051 = "";
            COPTC.TC052 = "0";
            COPTC.TC053 = "台灣中油股份有限公司";
            COPTC.TC054 = "";
            COPTC.TC055 = "";
            COPTC.TC056 = "";
            COPTC.TC057 = "";
            COPTC.TC058 = "";
            COPTC.TC059 = "";
            COPTC.TC060 = "";
            COPTC.TC061 = "";
            COPTC.TC062 = "";
            COPTC.TC063 = "台北市信義區松仁路3號";
            COPTC.TC064 = "";
            COPTC.TC065 = "";
            COPTC.TC066 = "";
            COPTC.TC067 = "";
            COPTC.TC068 = "";
            COPTC.TC069 = "";
            COPTC.TC070 = "N";
            COPTC.TC071 = "";
            COPTC.TC072 = "0";
            COPTC.TC073 = "0";
            COPTC.TC074 = "0";
            COPTC.TC075 = "";
            COPTC.TC076 = "";
            COPTC.TC077 = "";
            COPTC.TC078 = "";
            COPTC.TC079 = "";
            COPTC.TC080 = "";
            COPTC.TC081 = "";
            COPTC.TC082 = "";
            COPTC.TC083 = "";
            COPTC.TC084 = "";
            COPTC.TC085 = "";
            COPTC.TC086 = "";
            COPTC.TC087 = "";
            COPTC.TC088 = "";
            COPTC.TC089 = "";
            COPTC.TC090 = "";
            COPTC.TC091 = "";
            COPTC.TC092 = "N";
            COPTC.TC093 = "";
            COPTC.TC094 = "";
            COPTC.TC095 = "";
            COPTC.TC096 = "";
            COPTC.TC097 = "";
            COPTC.TC098 = "";
            COPTC.TC099 = "1";
            COPTC.TC100 = "N";
            COPTC.TC101 = "";
            COPTC.TC102 = "";
            COPTC.TC103 = "0";
            COPTC.TC104 = "";
            COPTC.TC105 = "";
            COPTC.TC106 = "1";
            COPTC.TC107 = "0";
            COPTC.TC108 = "0";
            COPTC.TC109 = "0";
            COPTC.TC110 = "0";
            COPTC.TC111 = "0";
            COPTC.TC112 = "0";
            COPTC.TC113 = "";
            COPTC.TC114 = "";
            COPTC.TC115 = "";
            COPTC.TC116 = "1";
            COPTC.TC117 = "";
            COPTC.TC118 = "0";
            COPTC.TC119 = "0";
            COPTC.TC120 = "0";
            COPTC.TC121 = "7";
            COPTC.TC122 = "";
            COPTC.TC123 = "";
            COPTC.TC124 = "03707901";
            COPTC.TC125 = "";
            COPTC.TC126 = "";
            COPTC.TC127 = "";
            COPTC.TC128 = "";
            COPTC.TC129 = "";
            COPTC.TC130 = "";
            COPTC.TC131 = "";
            COPTC.TC132 = "";
            COPTC.TC133 = "";
            COPTC.TC134 = "";
            COPTC.TC135 = "";
            COPTC.TC136 = "";
            COPTC.TC137 = "";
            COPTC.TC138 = "";
            COPTC.TC139 = "";
            COPTC.TC140 = "";
            COPTC.TC141 = "0";
            COPTC.TC142 = "";
            COPTC.TC143 = "";
            COPTC.TC144 = "";
            COPTC.TC145 = "";
            COPTC.TC146 = "";
            COPTC.UDF01 = "";
            COPTC.UDF02 = "";
            COPTC.UDF03 = "";
            COPTC.UDF04 = "";
            COPTC.UDF05 = "";
            COPTC.UDF06 = "0";
            COPTC.UDF07 = "0";
            COPTC.UDF08 = "0";
            COPTC.UDF09 = "0";
            COPTC.UDF10 = "0";




            

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
            UPDATETEMPCOPMAORDERRS();
        }

        #endregion


    }
}
