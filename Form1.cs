using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using HtmlAgilityPack;
using System.IO;
using System.Net;
using System.Threading;
using System.Data.SqlClient;
using System.Collections.Specialized;
using System.Web;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace NaNa_exchangerate
{
    public partial class Form1 : Form
    {
        public string err = "";
        private string DBHostStr = DBLink1.DBHost.Trim();
        private string DBStr = "NaNa";
        private string DBUser = "sa";
        private string DBPasd = "portweLL$";
        private bool debug = true;
        public Form1()
        {
            InitializeComponent();
            Load += new System.EventHandler(this.Form1_Load);
            Shown += new System.EventHandler(this.Form1_Shown);
            Closing += new CancelEventHandler(Form1_Closing);
            Closed += new EventHandler(Form1_Closed);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
                System.Net.ServicePointManager.SecurityProtocol = (SecurityProtocolType)192 | (SecurityProtocolType)768 | (SecurityProtocolType)3072;

                WebClient url = new WebClient();
                time.Text = "牌告時間：" + DateTime.Now;
                MemoryStream ms = new MemoryStream(url.DownloadData("https://www.ctbcbank.com/twrbo/zh_tw/dep_index/dep_ratequery/dep_foreign_rates.html"));    //中國信託
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.Load(ms, Encoding.Default);

                //string usdrate = doc.DocumentNode.SelectSingleNode("//*[@id='USD_saleRate2']").InnerText;
                //string eurrate = doc.DocumentNode.SelectSingleNode("//*[@id='EUR_saleRate2']").InnerText;
                //string gbprate = doc.DocumentNode.SelectSingleNode("//*[@id='GBP_saleRate2']").InnerText;
                //string inrrate = doc.DocumentNode.SelectSingleNode(@"/html/body/div/div[4]/div/div[1]/table/tr[20]/td[3]").InnerText.Trim();
                //string jpyrate = doc.DocumentNode.SelectSingleNode("//*[@id='JPY_saleRate2']").InnerText;
                //string krwrate = doc.DocumentNode.SelectSingleNode(@"/html/body/div/div[4]/div/div[1]/table/tr[17]/td[3]").InnerText.Trim();
                //string myrrate = doc.DocumentNode.SelectSingleNode(@"/html/body/div/div[4]/div/div[1]/table/tr[15]/td[3]").InnerText.Trim();
                //string rmbrate = doc.DocumentNode.SelectSingleNode("//*[@id='CNY_saleRate2']").InnerText;
                //string sgdrate = doc.DocumentNode.SelectSingleNode("//*[@id='SGD_saleRate2']").InnerText;




                System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
                System.Net.ServicePointManager.SecurityProtocol = (SecurityProtocolType)192 | (SecurityProtocolType)768 | (SecurityProtocolType)3072;
                WebClient bosan = new WebClient();
                MemoryStream ma = new MemoryStream(bosan.DownloadData("https://rate.bot.com.tw/xrt?Lang=en-US"));    //台灣銀行
                HtmlAgilityPack.HtmlDocument test = new HtmlAgilityPack.HtmlDocument();
                test.Load(ma, Encoding.Default);

                string usdrate = test.DocumentNode.SelectSingleNode(@"/html/body/div[1]/main/div[4]/table/tbody/tr[1]/td[5]").InnerText.Trim();
                string hkdrate = test.DocumentNode.SelectSingleNode(@"/html/body/div[1]/main/div[4]/table/tbody/tr[2]/td[5]").InnerText.Trim();
                string eurrate = test.DocumentNode.SelectSingleNode(@"/html/body/div[1]/main/div[4]/table/tbody/tr[15]/td[5]").InnerText.Trim();
                string gbprate = test.DocumentNode.SelectSingleNode(@"/html/body/div[1]/main/div[4]/table/tbody/tr[3]/td[5]").InnerText.Trim();
                string jpyrate = test.DocumentNode.SelectSingleNode(@"/html/body/div[1]/main/div[4]/table/tbody/tr[8]/td[5]").InnerText.Trim();
                string cnyrate = test.DocumentNode.SelectSingleNode(@"/html/body/div[1]/main/div[4]/table/tbody/tr[19]/td[5]").InnerText.Trim();
                string sgdrate = test.DocumentNode.SelectSingleNode(@"/html/body/div[1]/main/div[4]/table/tbody/tr[6]/td[5]").InnerText.Trim();


                USDrate.Text = usdrate;
                HKDrate.Text = hkdrate;
                EURrate.Text = eurrate;
                GBPrate.Text = gbprate;
                //INRrate.Text = inrrate;
                JPYrate.Text = jpyrate;
                //KRWrate.Text = krwrate;
                //MYRrate.Text = myrrate;
                CNYrate.Text = cnyrate;
                SGDrate.Text = sgdrate;

                doc = null;
                url = null;
                ms.Close();
                test = null;
                bosan = null;
                ma.Close();

                /*
                HtmlMeta meta = new HtmlMeta();
                meta.Attributes.Add("http-equiv", "refresh");
                //設定秒數，一天後後執行頁面更新
                meta.Content = "86400";
                this.Header.Controls.Add(meta);
                */

                //匯入資料庫

                insertUSDRATE(usdrate);
                insertEURRATE(eurrate);
                insertGBPRATE(gbprate);
                //insertINRRATE(inrrate);
                insertJPYRATE(jpyrate);
                //insertKRWRATE(krwrate);
                //insertMYRRATE(myrrate);
                insertCNYRATE(cnyrate);
                insertSGDRATE(sgdrate);
                insertHKDRATE(hkdrate);

                this.Close();
            }
            catch (Exception ex)
            {
                Label1.Text = ex.Message;
            }
        }
        private void Form1_Shown(object sender, EventArgs e)
        {
        }
        private void Form1_Closing(object sender, EventArgs e)
        {
        }
        private void Form1_Closed(object sender, EventArgs e)
        {
        }


        public void insertUSDRATE(string usdrate)
        {
            SqlConnectionStringBuilder builder2 = new SqlConnectionStringBuilder();
            builder2.Add("User id", DBUser.Trim());
            builder2.Add("Initial Catalog", DBStr.Trim());
            builder2.Add("Data Source", DBHostStr.Trim());
            builder2.Add("Password", DBPasd.Trim());
            SqlConnection cn2 = new SqlConnection(builder2.ConnectionString);
            SqlCommand cmd2 = cn2.CreateCommand();
            cmd2.Parameters.Add(new SqlParameter("@From_Currency_code", "USD"));
            cmd2.Parameters.Add(new SqlParameter("@To_Currency_code", "TWD"));
            cmd2.Parameters.Add(new SqlParameter("@Rate_YYYYMMDD", DateTime.Now.ToString("yyyyMMdd")));
            cmd2.Parameters.Add(new SqlParameter("@Exchange_rate", Convert.ToSingle(usdrate)));
            cmd2.Parameters.Add(new SqlParameter("@Creation_date", DateTime.Now));
            cmd2.Parameters.Add(new SqlParameter("@Created_by", 195));
            cmd2.CommandText = "insert into Exchange_date_rate(From_Currency_code,To_Currency_code,Rate_YYYYMMDD,Exchange_rate,Creation_date,Created_by) values (@From_Currency_code,@To_Currency_code,@Rate_YYYYMMDD,@Exchange_rate,@Creation_date,@Created_by) ";
            try
            {
                cn2.Open();
                cmd2.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                err = ex.Message;
            }
            finally
            {
                cn2.Close();
                cn2.Dispose();
            }
            return;
        }
        public void insertEURRATE(string eurrate)
        {
            SqlConnectionStringBuilder builder2 = new SqlConnectionStringBuilder();
            builder2.Add("User id", DBUser.Trim());
            builder2.Add("Initial Catalog", DBStr.Trim());
            builder2.Add("Data Source", DBHostStr.Trim());
            builder2.Add("Password", DBPasd.Trim());
            SqlConnection cn2 = new SqlConnection(builder2.ConnectionString);
            SqlCommand cmd2 = cn2.CreateCommand();
            cmd2.Parameters.Add(new SqlParameter("@From_Currency_code", "EUR"));
            cmd2.Parameters.Add(new SqlParameter("@To_Currency_code", "TWD"));
            cmd2.Parameters.Add(new SqlParameter("@Rate_YYYYMMDD", DateTime.Now.ToString("yyyyMMdd")));
            cmd2.Parameters.Add(new SqlParameter("@Exchange_rate", Convert.ToSingle(eurrate)));
            cmd2.Parameters.Add(new SqlParameter("@Creation_date", DateTime.Now));
            cmd2.Parameters.Add(new SqlParameter("@Created_by", 195));
            cmd2.CommandText = "insert into Exchange_date_rate(From_Currency_code,To_Currency_code,Rate_YYYYMMDD,Exchange_rate,Creation_date,Created_by) values (@From_Currency_code,@To_Currency_code,@Rate_YYYYMMDD,@Exchange_rate,@Creation_date,@Created_by) ";
            try
            {
                cn2.Open();
                cmd2.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                err = ex.Message;
            }
            finally
            {
                cn2.Close();
                cn2.Dispose();
            }
            return;
        }
        public void insertGBPRATE(string gbprate)
        {
            SqlConnectionStringBuilder builder2 = new SqlConnectionStringBuilder();
            builder2.Add("User id", DBUser.Trim());
            builder2.Add("Initial Catalog", DBStr.Trim());
            builder2.Add("Data Source", DBHostStr.Trim());
            builder2.Add("Password", DBPasd.Trim());
            SqlConnection cn2 = new SqlConnection(builder2.ConnectionString);
            SqlCommand cmd2 = cn2.CreateCommand();
            cmd2.Parameters.Add(new SqlParameter("@From_Currency_code", "GBP"));
            cmd2.Parameters.Add(new SqlParameter("@To_Currency_code", "TWD"));
            cmd2.Parameters.Add(new SqlParameter("@Rate_YYYYMMDD", DateTime.Now.ToString("yyyyMMdd")));
            cmd2.Parameters.Add(new SqlParameter("@Exchange_rate", Convert.ToSingle(gbprate)));
            cmd2.Parameters.Add(new SqlParameter("@Creation_date", DateTime.Now));
            cmd2.Parameters.Add(new SqlParameter("@Created_by", 195));
            cmd2.CommandText = "insert into Exchange_date_rate(From_Currency_code,To_Currency_code,Rate_YYYYMMDD,Exchange_rate,Creation_date,Created_by) values (@From_Currency_code,@To_Currency_code,@Rate_YYYYMMDD,@Exchange_rate,@Creation_date,@Created_by) ";
            try
            {
                cn2.Open();
                cmd2.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                err = ex.Message;
            }
            finally
            {
                cn2.Close();
                cn2.Dispose();
            }
            return;
        }
        public void insertINRRATE(string inrrate)
        {
            SqlConnectionStringBuilder builder2 = new SqlConnectionStringBuilder();
            builder2.Add("User id", DBUser.Trim());
            builder2.Add("Initial Catalog", DBStr.Trim());
            builder2.Add("Data Source", DBHostStr.Trim());
            builder2.Add("Password", DBPasd.Trim());
            SqlConnection cn2 = new SqlConnection(builder2.ConnectionString);
            SqlCommand cmd2 = cn2.CreateCommand();
            cmd2.Parameters.Add(new SqlParameter("@From_Currency_code", "INR"));
            cmd2.Parameters.Add(new SqlParameter("@To_Currency_code", "TWD"));
            cmd2.Parameters.Add(new SqlParameter("@Rate_YYYYMMDD", DateTime.Now.ToString("yyyyMMdd")));
            cmd2.Parameters.Add(new SqlParameter("@Exchange_rate", Convert.ToSingle(inrrate)));
            cmd2.Parameters.Add(new SqlParameter("@Creation_date", DateTime.Now));
            cmd2.Parameters.Add(new SqlParameter("@Created_by", 195));
            cmd2.CommandText = "insert into Exchange_date_rate(From_Currency_code,To_Currency_code,Rate_YYYYMMDD,Exchange_rate,Creation_date,Created_by) values (@From_Currency_code,@To_Currency_code,@Rate_YYYYMMDD,@Exchange_rate,@Creation_date,@Created_by) ";
            try
            {
                cn2.Open();
                cmd2.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                err = ex.Message;
            }
            finally
            {
                cn2.Close();
                cn2.Dispose();
            }
            return;
        }
        public void insertJPYRATE(string jpyrate)
        {
            SqlConnectionStringBuilder builder2 = new SqlConnectionStringBuilder();
            builder2.Add("User id", DBUser.Trim());
            builder2.Add("Initial Catalog", DBStr.Trim());
            builder2.Add("Data Source", DBHostStr.Trim());
            builder2.Add("Password", DBPasd.Trim());
            SqlConnection cn2 = new SqlConnection(builder2.ConnectionString);
            SqlCommand cmd2 = cn2.CreateCommand();
            cmd2.Parameters.Add(new SqlParameter("@From_Currency_code", "JPY"));
            cmd2.Parameters.Add(new SqlParameter("@To_Currency_code", "TWD"));
            cmd2.Parameters.Add(new SqlParameter("@Rate_YYYYMMDD", DateTime.Now.ToString("yyyyMMdd")));
            cmd2.Parameters.Add(new SqlParameter("@Exchange_rate", Convert.ToSingle(jpyrate)));
            cmd2.Parameters.Add(new SqlParameter("@Creation_date", DateTime.Now));
            cmd2.Parameters.Add(new SqlParameter("@Created_by", 195));
            cmd2.CommandText = "insert into Exchange_date_rate(From_Currency_code,To_Currency_code,Rate_YYYYMMDD,Exchange_rate,Creation_date,Created_by) values (@From_Currency_code,@To_Currency_code,@Rate_YYYYMMDD,@Exchange_rate,@Creation_date,@Created_by) ";
            try
            {
                cn2.Open();
                cmd2.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                err = ex.Message;
            }
            finally
            {
                cn2.Close();
                cn2.Dispose();
            }
            return;
        }
        public void insertKRWRATE(string krwrate)
        {
            SqlConnectionStringBuilder builder2 = new SqlConnectionStringBuilder();
            builder2.Add("User id", DBUser.Trim());
            builder2.Add("Initial Catalog", DBStr.Trim());
            builder2.Add("Data Source", DBHostStr.Trim());
            builder2.Add("Password", DBPasd.Trim());
            SqlConnection cn2 = new SqlConnection(builder2.ConnectionString);
            SqlCommand cmd2 = cn2.CreateCommand();
            cmd2.Parameters.Add(new SqlParameter("@From_Currency_code", "KRW"));
            cmd2.Parameters.Add(new SqlParameter("@To_Currency_code", "TWD"));
            cmd2.Parameters.Add(new SqlParameter("@Rate_YYYYMMDD", DateTime.Now.ToString("yyyyMMdd")));
            cmd2.Parameters.Add(new SqlParameter("@Exchange_rate", Convert.ToSingle(krwrate)));
            cmd2.Parameters.Add(new SqlParameter("@Creation_date", DateTime.Now));
            cmd2.Parameters.Add(new SqlParameter("@Created_by", 195));
            cmd2.CommandText = "insert into Exchange_date_rate(From_Currency_code,To_Currency_code,Rate_YYYYMMDD,Exchange_rate,Creation_date,Created_by) values (@From_Currency_code,@To_Currency_code,@Rate_YYYYMMDD,@Exchange_rate,@Creation_date,@Created_by) ";
            try
            {
                cn2.Open();
                cmd2.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                err = ex.Message;
            }
            finally
            {
                cn2.Close();
                cn2.Dispose();
            }
            return;
        }
        public void insertMYRRATE(string myrrate)
        {
            SqlConnectionStringBuilder builder2 = new SqlConnectionStringBuilder();
            builder2.Add("User id", DBUser.Trim());
            builder2.Add("Initial Catalog", DBStr.Trim());
            builder2.Add("Data Source", DBHostStr.Trim());
            builder2.Add("Password", DBPasd.Trim());
            SqlConnection cn2 = new SqlConnection(builder2.ConnectionString);
            SqlCommand cmd2 = cn2.CreateCommand();
            cmd2.Parameters.Add(new SqlParameter("@From_Currency_code", "MYR"));
            cmd2.Parameters.Add(new SqlParameter("@To_Currency_code", "TWD"));
            cmd2.Parameters.Add(new SqlParameter("@Rate_YYYYMMDD", DateTime.Now.ToString("yyyyMMdd")));
            cmd2.Parameters.Add(new SqlParameter("@Exchange_rate", Convert.ToSingle(myrrate)));
            cmd2.Parameters.Add(new SqlParameter("@Creation_date", DateTime.Now));
            cmd2.Parameters.Add(new SqlParameter("@Created_by", 195));
            cmd2.CommandText = "insert into Exchange_date_rate(From_Currency_code,To_Currency_code,Rate_YYYYMMDD,Exchange_rate,Creation_date,Created_by) values (@From_Currency_code,@To_Currency_code,@Rate_YYYYMMDD,@Exchange_rate,@Creation_date,@Created_by) ";
            try
            {
                cn2.Open();
                cmd2.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                err = ex.Message;
            }
            finally
            {
                cn2.Close();
                cn2.Dispose();
            }
            return;
        }
        public void insertCNYRATE(string cnyrate)
        {
            SqlConnectionStringBuilder builder2 = new SqlConnectionStringBuilder();
            builder2.Add("User id", DBUser.Trim());
            builder2.Add("Initial Catalog", DBStr.Trim());
            builder2.Add("Data Source", DBHostStr.Trim());
            builder2.Add("Password", DBPasd.Trim());
            SqlConnection cn2 = new SqlConnection(builder2.ConnectionString);
            SqlCommand cmd2 = cn2.CreateCommand();
            cmd2.Parameters.Add(new SqlParameter("@From_Currency_code", "CNY"));
            cmd2.Parameters.Add(new SqlParameter("@To_Currency_code", "TWD"));
            cmd2.Parameters.Add(new SqlParameter("@Rate_YYYYMMDD", DateTime.Now.ToString("yyyyMMdd")));
            cmd2.Parameters.Add(new SqlParameter("@Exchange_rate", Convert.ToSingle(cnyrate)));
            cmd2.Parameters.Add(new SqlParameter("@Creation_date", DateTime.Now));
            cmd2.Parameters.Add(new SqlParameter("@Created_by", 195));
            cmd2.CommandText = "insert into Exchange_date_rate(From_Currency_code,To_Currency_code,Rate_YYYYMMDD,Exchange_rate,Creation_date,Created_by) values (@From_Currency_code,@To_Currency_code,@Rate_YYYYMMDD,@Exchange_rate,@Creation_date,@Created_by) ";
            try
            {
                cn2.Open();
                cmd2.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                err = ex.Message;
            }
            finally
            {
                cn2.Close();
                cn2.Dispose();
            }
            return;
        }
        public void insertSGDRATE(string sgdrate)
        {
            SqlConnectionStringBuilder builder2 = new SqlConnectionStringBuilder();
            builder2.Add("User id", DBUser.Trim());
            builder2.Add("Initial Catalog", DBStr.Trim());
            builder2.Add("Data Source", DBHostStr.Trim());
            builder2.Add("Password", DBPasd.Trim());
            SqlConnection cn2 = new SqlConnection(builder2.ConnectionString);
            SqlCommand cmd2 = cn2.CreateCommand();
            cmd2.Parameters.Add(new SqlParameter("@From_Currency_code", "SGD"));
            cmd2.Parameters.Add(new SqlParameter("@To_Currency_code", "TWD"));
            cmd2.Parameters.Add(new SqlParameter("@Rate_YYYYMMDD", DateTime.Now.ToString("yyyyMMdd")));
            cmd2.Parameters.Add(new SqlParameter("@Exchange_rate", Convert.ToSingle(sgdrate)));
            cmd2.Parameters.Add(new SqlParameter("@Creation_date", DateTime.Now));
            cmd2.Parameters.Add(new SqlParameter("@Created_by", 195));
            cmd2.CommandText = "insert into Exchange_date_rate(From_Currency_code,To_Currency_code,Rate_YYYYMMDD,Exchange_rate,Creation_date,Created_by) values (@From_Currency_code,@To_Currency_code,@Rate_YYYYMMDD,@Exchange_rate,@Creation_date,@Created_by) ";
            try
            {
                cn2.Open();
                cmd2.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                err = ex.Message;
            }
            finally
            {
                cn2.Close();
                cn2.Dispose();
            }
            return;
        }
        public void insertHKDRATE(string hkdrate)
        {
            SqlConnectionStringBuilder builder2 = new SqlConnectionStringBuilder();
            builder2.Add("User id", DBUser.Trim());
            builder2.Add("Initial Catalog", DBStr.Trim());
            builder2.Add("Data Source", DBHostStr.Trim());
            builder2.Add("Password", DBPasd.Trim());
            SqlConnection cn2 = new SqlConnection(builder2.ConnectionString);
            SqlCommand cmd2 = cn2.CreateCommand();
            cmd2.Parameters.Add(new SqlParameter("@From_Currency_code", "HKD"));
            cmd2.Parameters.Add(new SqlParameter("@To_Currency_code", "TWD"));
            cmd2.Parameters.Add(new SqlParameter("@Rate_YYYYMMDD", DateTime.Now.ToString("yyyyMMdd")));
            cmd2.Parameters.Add(new SqlParameter("@Exchange_rate", Convert.ToSingle(hkdrate)));
            cmd2.Parameters.Add(new SqlParameter("@Creation_date", DateTime.Now));
            cmd2.Parameters.Add(new SqlParameter("@Created_by", 195));
            cmd2.CommandText = "insert into Exchange_date_rate(From_Currency_code,To_Currency_code,Rate_YYYYMMDD,Exchange_rate,Creation_date,Created_by) values (@From_Currency_code,@To_Currency_code,@Rate_YYYYMMDD,@Exchange_rate,@Creation_date,@Created_by) ";
            try
            {
                cn2.Open();
                cmd2.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                err = ex.Message;
            }
            finally
            {
                cn2.Close();
                cn2.Dispose();
            }
            return;
        }
    }
}

