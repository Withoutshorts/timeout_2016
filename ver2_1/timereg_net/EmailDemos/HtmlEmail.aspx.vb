'Add the namespace for the email-related classes and for working w/streams
Imports System.Net.Mail
Imports System.IO


Partial Class HtmlEmail
    Inherits System.Web.UI.Page

    'Sends an HTML-formatted email using the IsBodyHtml property...
    Protected Sub SendHTMLEmail_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SendHTMLEmail.Click
        '!!! UPDATE THIS VALUE TO YOUR EMAIL ADDRESS
        Const ToAddress As String = "YOUR EMAIL ADDRESS HERE!!"

        '(1) Create the MailMessage instance
        Dim mm As New MailMessage(UsersEmail.Text, ToAddress)

        '(2) Assign the MailMessage's properties
        mm.Subject = "HTML-Formatted Email Demo Using the IsBodyHtml Property"
        mm.Body = "<h2>This is an HTML-Formatted Email Send Using the <code>IsBodyHtml</code> Property</h2><p>Isn't HTML <em>neat</em>?</p><p>You can make all sorts of <span style=""color:red;font-weight:bold;"">pretty colors!!</span>.</p>"
        mm.IsBodyHtml = True

        '(3) Create the SmtpClient object
        Dim smtp As New SmtpClient

        '(4) Send the MailMessage (will use the Web.config settings)
        smtp.Send(mm)



        'Display a client-side popup, explaining that the email has been sent
        ClientScript.RegisterStartupScript(Me.GetType(), "HiMom!", String.Format("alert('An HTML-formatted email using the IsBodyHtml property has been sent to {0}');", ToAddress.Replace("'", "\'")), True)
    End Sub

End Class
