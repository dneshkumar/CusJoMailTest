using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Timers;

namespace EmailSample
{
    public partial class Service1 : ServiceBase
    {
        Timer timer = new Timer();
        Timer timer2 = new Timer();
        public static StringBuilder sbStatus = new StringBuilder();
        SqlConnection SqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnectionString"].ConnectionString);
        public int TimeInterval = Convert.ToInt16(ConfigurationManager.AppSettings["ScheduledTime"]);
        public int DailyRemainderInterval = Convert.ToInt16(ConfigurationManager.AppSettings["DailyRemainderInterval"]);
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("Service is started at " + DateTime.Now);
            timer.Elapsed += new ElapsedEventHandler(MailService);
            int timing = TimeInterval * 86400000; // for one day interval
            timer.Interval = timing; //number in milisecinds  
            timer.Enabled = true;

            timer2.Elapsed += new ElapsedEventHandler(DailyRemainder);
            int timing2 = DailyRemainderInterval * 86400000; // one day after mail send
            timer2.Interval = timing2; //number in milisecinds  
            timer2.Enabled = true;
        }
        private void MailService(object source, ElapsedEventArgs e)
        {
            MailCreation();
        }

        private void DailyRemainder(object source, ElapsedEventArgs e)
        {
            RemainderMail();
        }

        #region Insert Vendor in SAP
        private void MailCreation()
        {
            //bool result = false;
            try
            {
                DataTable dtMailIds = GetMailIds();

                if (dtMailIds.Rows.Count > 0)
                {
                    SendEmail(dtMailIds);
                    WriteToFile("Email sent to users");
                }
                else
                {
                    WriteToFile("No Record Found !");
                }
            }
            catch (Exception ex)
            {
                SendErrorToText(ex, "Error_");

            }
            finally
            {

            }
        }
        #endregion

        #region SAP to SQL Update Vendor Code
        private void RemainderMail()
        {
            try
            {
                DataTable dtMailIds = GetRemainderMailIds();

                if (dtMailIds.Rows.Count > 0)
                {
                    SendEmail(dtMailIds);
                    WriteToFile("Email sent to users");
                }
                else
                {
                    WriteToFile("No Record Found !");
                }
                int timing2 = 1 * 86400000; // updating the interval to one day
                timer2.Interval = timing2; //number in milisecinds 

            }
            catch (Exception ex)
            {
                SendErrorToText(ex, "Error_");
            }
        }
        #endregion

        #region Email
        void SendEmail(DataTable dtListEmail)
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");

                mail.From = new MailAddress(ConfigurationManager.AppSettings["FromMailId"].ToString());
                foreach (DataRow drrecipient in dtListEmail.Rows)
                {
                    string url = "http://localhost:50905/Index.aspx" + "?emailid=" + drrecipient["email"].ToString(); // to uniquly identify the user
                    mail.To.Add(drrecipient["email"].ToString());
                    mail.Body = @"<a href= '" + url + "'>Click Here</a>";
                }
                mail.Subject = "Test Mail- Click the below link";

                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["MailId"].ToString(), ConfigurationManager.AppSettings["Password"].ToString());
                SmtpServer.EnableSsl = true;

                SmtpServer.Send(mail);
            }
            catch (Exception ex)
            {
                SendErrorToText(ex, "Error_");
            }
        }
        #endregion

        #region Execution
        public DataTable GetMailIds()
        {
            DataTable dt = ExecuteDataTableProcedure();
            return dt;
        }

        public DataTable GetRemainderMailIds()
        {
            DataTable dt = GetRemainderMailId();
            return dt;
        }
        public static DataTable ExecuteDataTableProcedure()
        {
            SqlConnection SqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnectionString"].ConnectionString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = SqlConn;
            cmd.CommandText = "uspGetEmailIds";
            cmd.CommandType = CommandType.StoredProcedure;
            SqlConn.Open();
            SqlDataAdapter objda = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            objda.Fill(dt);
            SqlConn.Close();
            return dt;
        }

        public static DataTable GetRemainderMailId()
        {
            SqlConnection SqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnectionString"].ConnectionString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = SqlConn;
            cmd.CommandText = "uspGetRemainderEmailIds";
            cmd.CommandType = CommandType.StoredProcedure;
            SqlConn.Open();
            SqlDataAdapter objda = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            objda.Fill(dt);
            SqlConn.Close();
            return dt;
        }

        #endregion

        #region Log 
        private void SendErrorToText(Exception ex, string ErrorFileName)
        {
            
            var line = Environment.NewLine;
            string path = System.Configuration.ConfigurationManager.AppSettings["LogPath"] + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = System.Configuration.ConfigurationManager.AppSettings["LogPath"] + "\\Logs\\" + ErrorFileName + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    string error = "Error Message:" + " " + ex.Message.ToString() + " " + ErrorFileName + line + "Exception Type:" + " " + ex.GetType().Name.ToString() + line + "StackTrace: " + ex.StackTrace;
                    sw.WriteLine("-----------Exception Details on " + " " + DateTime.Now.ToString() + "-----------------");
                    sw.WriteLine("-------------------------------------------------------------------------------------");
                    sw.WriteLine(error);
                    sw.WriteLine("--------------------------------*End*------------------------------------------");
                    sw.WriteLine(line);
                    sw.Flush();
                    sw.Close();
                }
            }
        }

        public void WriteToFile(string Message)
        {
            string path = System.Configuration.ConfigurationManager.AppSettings["LogPath"] + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = System.Configuration.ConfigurationManager.AppSettings["LogPath"] + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(DateTime.Now.ToString() + " : " + Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(DateTime.Now.ToString() + " : " + Message);
                }
            }
        }
        #endregion
        protected override void OnStop()
        {
            WriteToFile("Service is stopped at " + DateTime.Now);
        }
    }
}
