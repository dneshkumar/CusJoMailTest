using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net.Mail;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace EmailWebApp
{
    public partial class Index : System.Web.UI.Page
    {

        public SqlConnection _sqlConnection = new SqlConnection(@"Data Source=(local)\SQLEXPRESS;Database=Test; Integrated Security=True");
        protected void Page_Load(object sender, EventArgs e)
        {
            string result = string.Empty;
            try
            {
                if (Request.QueryString["emailid"] != null && Request.QueryString["emailid"] != string.Empty)
                {
                    string emailid = Request.QueryString["emailid"];
                    SendEmail(emailid);
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = _sqlConnection;
                    cmd.CommandText = "uspUpdateEmailStatus";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@email", emailid);
                    _sqlConnection.Open();
                    result = cmd.ExecuteScalar().ToString();
                }

            }
            catch (Exception ex)
            {
            }

        }

        void SendEmail(string emailid)
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
                mail.From = new MailAddress(ConfigurationManager.AppSettings["FromMailId"].ToString());
                mail.To.Add(emailid);
                mail.Body = "Thanks you";
                mail.Subject = "Thank you";
                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["MailId"].ToString(), ConfigurationManager.AppSettings["Password"].ToString());
                SmtpServer.EnableSsl = true;
                SmtpServer.Send(mail);
            }
            catch (Exception ex)
            {
                {
                }
            }
        }
    }
}
    