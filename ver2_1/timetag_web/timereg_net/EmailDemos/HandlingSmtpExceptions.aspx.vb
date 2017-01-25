'Add the namespace for the email-related classes AND System.Net for the NetworkCredential class
Imports System.Net.Mail
Imports System.Net


Partial Class HandlingSmtpExceptions
    Inherits System.Web.UI.Page

    Protected Sub SendEmail_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SendEmail.Click
        '!!! UPDATE THIS VALUE TO YOUR EMAIL ADDRESS
        Const ToAddress As String = "YOUR EMAIL ADDRESS HERE!!"

        Try
            '(1) Create the MailMessage instance
            Dim mm As New MailMessage(UsersEmail.Text, ToAddress)

            '(2) Assign the MailMessage's properties
            mm.Subject = "Test Email... DO NOT PANIC!!!1!!!111!!"
            mm.Body = "This is a test message..." & vbCrLf & vbCrLf & "** Do not panic. This is only a test **"
            mm.IsBodyHtml = False

            '(3) Create the SmtpClient object
            Dim smtp As New SmtpClient

            'Set the SMTP settings...
            smtp.Host = Hostname.Text
            If Not String.IsNullOrEmpty(Port.Text) Then
                smtp.Port = Convert.ToInt32(Port.Text)
            End If

            If Not String.IsNullOrEmpty(Username.Text) Then
                smtp.Credentials = New NetworkCredential(Username.Text, Password.Text)
            End If

            '(4) Send the MailMessage (will use the Web.config settings)
            smtp.Send(mm)

            'Display a client-side popup, explaining that the email has been sent
            ClientScript.RegisterStartupScript(Me.GetType(), "HiMom!", String.Format("alert('An test email has successfully been sent to {0}');", ToAddress.Replace("'", "\'")), True)

        Catch smtpEx As SmtpException
            'A problem occurred when sending the email message
            ClientScript.RegisterStartupScript(Me.GetType(), "OhCrap", String.Format("alert('There was a problem in sending the email: {0}');", smtpEx.Message.Replace("'", "\'")), True)

        Catch generalEx As Exception
            'Some other problem occurred
            ClientScript.RegisterStartupScript(Me.GetType(), "OhCrap", String.Format("alert('There was a general problem: {0}');", generalEx.Message.Replace("'", "\'")), True)

        End Try
    End Sub

    Protected Sub LoadBadValues_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LoadBadValues.Click
        UsersEmail.Text = "scott@example.com"
        Hostname.Text = "smtp.example.com"
        Port.Text = "25"
        Username.Text = "Scott"
    End Sub
End Class
