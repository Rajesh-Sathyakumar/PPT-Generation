using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using System.Configuration;
using System.Net;
using System.Net.Mail;
using System.Web;

namespace AOAService
{
    class Utilities
    {
        public string SVNUserName = ConfigurationManager.AppSettings["SVNUserID"];
        public string SVNPassword = ConfigurationManager.AppSettings["SVNPassword"];

        public string SqlLogReadQueryExecutor(SqlConnection connection, string SQLQuery)
        {
            connection.Open();
            var command = new SqlCommand(SQLQuery, connection);
            SqlDataReader result = command.ExecuteReader();

            string results = string.Empty;

            if (result.Read())
            {
                results = result["LEVEL"].ToString() + "|" + result["MESSAGE"].ToString();
                connection.Close();
                return results;
            }
            connection.Close();
            return null;
        }

        public void SQLQueryExecutor(string connectionString, string SQLQuery)
        {
            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                var command = new SqlCommand(SQLQuery, connection);
                command.CommandTimeout = 0;
                command.ExecuteNonQuery();
                connection.Close();
            }
        }

        public string SQLQueryResults(string connectionString, string SQLQuery)
        {
            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                var command = new SqlCommand(SQLQuery, connection);
                SqlDataReader result = command.ExecuteReader();

                string results = string.Empty;

                if (result.Read())
                {
                    results = result[0].ToString();
                    connection.Close();
                    return results;
                }
                connection.Close();
                return null;
            }
        }

        public void SVNQueryExecutor(string SVNPath, string SVNQueryName, string connectionString)
        {
            using (var SVNWebCall = new WebClient())
            {

                System.Net.ServicePointManager.ServerCertificateValidationCallback +=
                delegate(object sender, System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                System.Security.Cryptography.X509Certificates.X509Chain chain,
                System.Net.Security.SslPolicyErrors sslPolicyErrors)
                {
                    return true; // **** Always accept
                };

                SVNWebCall.Credentials = new System.Net.NetworkCredential(SVNUserName, SVNPassword);
                var Query = SVNWebCall.DownloadString(SVNPath + SVNQueryName);
                SQLQueryExecutor(connectionString, Query);
            }
        }

        public string ReplaceDatabaseName(string Statement, string DatabaseName)
        {
            string output;

            output = Regex.Replace(Statement, "<databasename>", DatabaseName);

            return output;
        }

        public string LoaderURLBuilder(string URLFromDatabase, string release, string LoaderName, string Server, string Uid, string Password, string DatabaseName)
        {
            string URL;

            URL = Replace(URLFromDatabase, "<ReleaseName>", release);
            URL = Replace(URL, "<LoaderName>", LoaderName);
            URL = Replace(URL, "<server>", Server);
            URL = Replace(URL, "<databasename>", DatabaseName);
            URL = Replace(URL, "<uid>", Uid);
            URL = Replace(URL, "<password>", Password);

            return URL;
        }

        public string Replace(string Statement, string characterToFind, string replaceCharacter)
        {
            return Regex.Replace(Statement, characterToFind, replaceCharacter);
        }

        public string ConnectionStringBuilder(string server, string databaseName)
        {
            var ConnectionString = "Server=" + server + ";Database=" + databaseName + ";Integrated Security=True"; ;
            return ConnectionString;
        }

        public DataSet LoadDataSet(string connectionString, string query)
        {
            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                var command = new SqlCommand(query, connection);

                var ds = new DataSet();
                var adapter = new SqlDataAdapter(command);
                adapter.Fill(ds);

                return ds;
            }
        }

        public DataTable LoadDataTable(string connectionString, string query)
        {
            return LoadDataSet(connectionString, query).Tables[0];
        }

        public void SendEmail(string AddTo, string AddBCC, string Subject, string Body, string File)
        {
            MailMessage msg = new MailMessage();
            msg.IsBodyHtml = true;
            string[] multito = AddTo.Split(',');
            foreach (string MultiEmailTo in multito)
            {
                msg.To.Add(new MailAddress(MultiEmailTo));
            }
            string[] multiCC = AddBCC.Split(',');
            foreach (string MultiEmailCC in multiCC)
            {
                msg.Bcc.Add(new MailAddress(MultiEmailCC));
            }
            msg.From = new MailAddress("DARA@Advisory.com");
            msg.Subject = Subject;
            msg.Body = Body;
            msg.Attachments.Add(new Attachment(File));
            SmtpClient smclient = new SmtpClient("192.168.17.46");
            smclient.Send(msg);
        }

        public void SendExceptionEmail(string AddTo, string Subject, string Body, string File)
        {
            MailMessage msg = new MailMessage();
            msg.IsBodyHtml = true;
            string[] multito = AddTo.Split(',');
            foreach (string MultiEmailTo in multito)
            {
                msg.To.Add(new MailAddress(MultiEmailTo));
            }
            msg.From = new MailAddress("DARA@Advisory.com");
            msg.Subject = Subject;
            msg.Body = Body;
            SmtpClient smclient = new SmtpClient("192.168.17.46");
            smclient.Send(msg);
        }
    }
}
