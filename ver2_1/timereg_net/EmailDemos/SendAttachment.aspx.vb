'Add the namespace for the email-related classes
Imports System.Net.Mail

Partial Class SendAttachment
    Inherits System.Web.UI.Page

    Protected Sub SendEmail_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SendEmail.Click
        'Make sure a file has been uploaded
        If String.IsNullOrEmpty(AttachmentFile.FileName) OrElse AttachmentFile.PostedFile Is Nothing Then
            Throw New ApplicationException("Egad, a file wasn't uploaded... you should probably use more graceful error handling than this, though...")
        End If

        '!!! UPDATE THIS VALUE TO YOUR EMAIL ADDRESS
        Const ToAddress As String = "YOUR EMAIL ADDRESS HERE!!"

        '(1) Create the MailMessage instance
        Dim mm As New MailMessage(UsersEmail.Text, ToAddress)

        '(2) Assign the MailMessage's properties
        mm.Subject = "Emailing an Uploaded File as an Attachment Demo"
        mm.Body = Body.Text
        mm.IsBodyHtml = False

        'Attach the file
        mm.Attachments.Add(New Attachment(AttachmentFile.PostedFile.InputStream, AttachmentFile.FileName))

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
