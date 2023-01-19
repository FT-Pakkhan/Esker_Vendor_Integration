using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace Esker_Vendor_Integration
{
    public partial class FTS00OVI : Form
    {
        public FTS00OVI()
        {
            InitializeComponent();
            ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }

        private void FTS00OVI_Load(object sender, EventArgs e)
        {
            btnExecute.PerformClick();
            this.Close();
        }

        private void btnExecute_Click(object sender, EventArgs e)
        {
            try
            {
                string connString = null;

                connString = @"Persist Security Info=True" +
                    ";Data Source=" + ConfigurationManager.AppSettings.Get("Server").ToString() +
                    ";Initial Catalog=" + ConfigurationManager.AppSettings.Get("CommonDb").ToString() +
                    ";User ID=" + ConfigurationManager.AppSettings.Get("DbUserName").ToString() +
                    ";Password=" + ConfigurationManager.AppSettings.Get("DbPassword").ToString() + ";";

                using (SqlConnection conn = new SqlConnection(connString))
                {
                    if (conn.State == ConnectionState.Open)
                        conn.Close();

                    conn.Open();
                    Log.AppendText("[Message] " + DateTime.Now.ToString() + " : " + ConfigurationManager.AppSettings.Get("CommonDb").ToString() + " successfully Connected!" + Environment.NewLine);


                    using (SqlDataAdapter da = new SqlDataAdapter("", conn))
                    {
                        da.SelectCommand.CommandText = "SELECT * FROM ODBC ";
                        da.SelectCommand.CommandTimeout = 0;
                        DataTable dt = new DataTable();
                        da.Fill(dt);

                        da.Dispose();

                        if (dt.Rows.Count > 0)
                        {
                            foreach (DataRow row in dt.Rows)
                            {
                                string SvrType = row["SvrType"].ToString();
                                string Server = row["Server"].ToString();
                                string LicSvr = row["LicSvr"].ToString();
                                string SQLLogin = row["SQLLogin"].ToString();
                                string SQLPass = row["SQLPass"].ToString();
                                string SAPDBName = row["SAPDBName"].ToString();
                                string SAPLogin = row["SAPLogin"].ToString();
                                string SAPPass = row["SAPPass"].ToString();
                                string InputFilePath = row["V_InputFilePath"].ToString();
                                string InputProcessedFilePath = row["V_InputProcessedFilePath"].ToString();
                                string OutputFilePath = row["V_OutputFilePath"].ToString();

                                VendorIntegration.Execute(this, SvrType, Server, LicSvr, SQLLogin, SQLPass, SAPDBName, SAPLogin, SAPPass,
                                    InputFilePath, InputProcessedFilePath, OutputFilePath);
                            }
                        }
                        dt.Dispose();
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                Log.AppendText("[Error] " + DateTime.Now.ToString() + " : " + ex.Message + Environment.NewLine);
            }
        }

        private void FTS00OVI_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Log.Text.Length >= 0)
            {
                Log.AppendText("[Message] " + DateTime.Now.ToString() + " : " + "System Ended." + Environment.NewLine);

                string path = @"" + ConfigurationManager.AppSettings.Get("LogFilePath").ToString() + DateTime.Now.ToString("yyyy-MM-dd");
                if (!System.IO.Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path);
                }

                string insideFileLog = "\\" + DateTime.Now.ToString("yyyy-MM-dd_hhmm") + ".txt";
                System.IO.File.WriteAllLines(path + insideFileLog, Log.Lines);
            }
        }
    }
}
