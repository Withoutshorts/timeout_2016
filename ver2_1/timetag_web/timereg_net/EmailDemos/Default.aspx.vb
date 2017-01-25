'Add the namespace for the email-related classes
Imports System.Net.Mail

Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub SendEmail_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SendEmail.Click
        '!!! UPDATE THIS VALUE TO YOUR EMAIL ADDRESS
        Const ToAddress As String = "sk@outzoure.dk"

        '(1) Create the MailMessage instance
        Dim mm As New MailMessage(UsersEmail.Text, ToAddress)

        '(2) Assign the MailMessage's properties
        mm.Subject = Subject.Text
        mm.Body = Body.Text
        mm.IsBodyHtml = False

        '(3) Create the SmtpClient object
        Dim smtp As New SmtpClient

        '(4) Send the MailMessage (will use the Web.config settings)
        smtp.Send(mm)


        'Show the EmailSentForm Panel and hide the EmailForm Panel
        EmailSentForm.Visible = True
        EmailForm.Visible = False
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            'On the first page load, hide the EmailSentForm Panel
            EmailSentForm.Visible = False
        End If
    End Sub
End Class
