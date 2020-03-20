using System;
using System.Net.Mail;
using System.Text;
using System.Net.Mime;
using System.Net;
using System.IO;

public partial class Outlook_mail : System.Web.UI.Page
{
    protected void Send_email_with_outlookCalender(object sender, EventArgs e)
    {

        // FYLD UD SELV TIL TEST!
        string emailFrom = Request.Form["Hvem sender Email?"];
        string emailFromPassword  Request.Form["password fra sender"];
        string emailTO = Request.Form["Email som skal modtage"];
        string calenderSubject  Request.Form["Emne"];
        string calenderBody  Request.Form["Text i emailen"];
        string location  Request.Form["Adresse"];
        string startDate  Request.Form["Start dato"];
        string endDate  Request.Form["Slut dato"];

        // Credentials
        var credentials = new NetworkCredential(emailFrom, emailFromPassword);

        // Mail message
        var mail = new MailMessage()
        {
            From = new MailAddress(emailFrom),
            Subject = calenderSubject,
            Body = calenderBody
        };

        mail.To.Add(new MailAddress(emailTO));

        // Smtp client 
        var client = new SmtpClient()
        {
            // Verdierne skal lige ændres i forhold til Serveren/MailServeren
            Port = 587,
            DeliveryMethod = SmtpDeliveryMethod.Network,
            UseDefaultCredentials = false,
            Host = "smtp.gmail.com",
            EnableSsl = true,
            Credentials = credentials
        };

        // Dette er filens indhold, syntaksten her laver filen af sig selv.
        StringBuilder str = new StringBuilder();
        str.AppendLine("BEGIN:VCALENDAR");
        str.AppendLine("PRODID:-//Outzource ApS//");
        str.AppendLine("VERSION:2.0");
        str.AppendLine("METHOD:REQUEST");
        str.AppendLine("BEGIN:VEVENT");
        str.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHHmmssZ}", startDate));
        str.AppendLine(string.Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", DateTime.UtcNow));
        str.AppendLine(string.Format("DTEND:{0:yyyyMMddTHHmmssZ}", endDate));
        str.AppendLine("LOCATION:" + location);
        str.AppendLine(string.Format("UID:{0}", Guid.NewGuid()));
        str.AppendLine(string.Format("DESCRIPTION:{0}", mail.Body));
        str.AppendLine(string.Format("X-ALT-DESC;FMTTYPE=text/html:{0}", mail.Body));
        str.AppendLine(string.Format("SUMMARY:{0}", mail.Subject));
        str.AppendLine(string.Format("ORGANIZER:MAILTO:{0}", mail.From.Address));


        ContentType contype = new ContentType("text/calendar");
        contype.Parameters.Add("method", "REQUEST");
        contype.Parameters.Add(calenderSubject, "AddToCalender.ics");
        AlternateView avCal = AlternateView.CreateAlternateViewFromString(str.ToString(), contype);
        mail.AlternateViews.Add(avCal);


        // Send it...         
        client.Send(mail);

    }
}
