using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Data;
//using Microsoft.Office.Interop;
//using Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Core;
using System.Reflection;
using System.Drawing;
using System.Globalization;

namespace ReportEdit
{
    class sendEmail
    {
        public static void SendMyMail()
        {


            /* !!NEW!! READING IN THE .CSV TO GET LIST OF ADDRESSES. Below comment is old 

             List<string> addyList = new List<string>();
             foreach (string line in File.ReadLines(Properties.Settings.Default.AddyList))
             {
                 addyList.Add(line);
             }
             */

            SmtpClient companySmtpClient = new SmtpClient("smtprelay.company.com");

            companySmtpClient.UseDefaultCredentials = true;

            MailAddress from = new MailAddress("ActiveBatchRunReport@company.com", "ActiveBatchRunReport");
            MailAddress to = new MailAddress("GrpDstISOne@company.com", "DstOne");
            MailAddress Recipient1 = new MailAddress("Bill@company.com", "Bill");
            MailAddress Recipient2 = new MailAddress("Tom@company.com", "Tom");
            MailAddress Recipient3 = new MailAddress("Gena@company.com", "Gena");
            MailAddress Recipient4 = new MailAddress("Clifford@company.com", "Clifford");
            MailAddress ccKate = new MailAddress("Kate@company.com", "Kate");


            MailMessage myMail = new System.Net.Mail.MailMessage(from, to);

            myMail.To.Add(Recipient1);
            myMail.To.Add(Recipient2);
            myMail.To.Add(Recipient3);
            myMail.To.Add(Recipient4);

            myMail.CC.Add(ccKate);

            myMail.Subject = "Daily Job Runs";
            myMail.SubjectEncoding = System.Text.Encoding.UTF8;

            myMail.Body = "Attached you will find an Excel spreadsheet" +
            "Total Job counts are listed at the bottom.";

            myMail.BodyEncoding = System.Text.Encoding.UTF8;
            myMail.IsBodyHtml = true;

            myMail.Attachments.Add(new Attachment(@"PathToAttachment"));

            companySmtpClient.Send(myMail);
        }

    }
}