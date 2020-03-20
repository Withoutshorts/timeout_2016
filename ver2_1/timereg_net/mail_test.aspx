<%@ Page Language="vb" Debug="true" %>

<%@ Import Namespace="System.Web.Services.Description" %>

<%@  Import Namespace="Microsoft.VisualBasic" %>
<%@  Import Namespace="System.Data"%>
<%@  Import Namespace="System.Data.Odbc" %>
<%@  Import Namespace="System.Web.Services" %> 

<%@  Import Namespace="System.Diagnostics" %>
<%@  Import Namespace="System.Web.UI" %> 
<%@  Import Namespace="System.Web.UI.Webcontrols" %>
<%@  Import Namespace="System.Collections.Generic" %>



<%@  Import Namespace="System.Net.Mail" %>







<!--@ Page Language="C#" AutoEventWireup="true" CodeFile="importer_job_wilke.cs" Inherits="importer_job_wilke" 
<!--@  Import Namespace="System.Globalization" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<script runat="server" language="vbscript">

    Public lto As String = "xxx"
    Public strOptions As String = ""


    Public emailto As String = "all"

    Sub Page_load()




    End Sub


    Private Sub btnOverRide_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bt.Click


        Dim Smtp_Server As New SmtpClient
        Dim e_mail As New MailMessage()
        'Smtp_Server.UseDefaultCredentials = True
        'Smtp_Server.Credentials = New Net.NetworkCredential("administrator", "Sok!2637")
        Smtp_Server.Port = 25 '587
        'Smtp_Server.EnableSsl = True
        Smtp_Server.Host = "formrelay.rackhosting.com"

        'TEST
        Dim email As String = "sk@outzource.dk"

        e_mail = New MailMessage()
        e_mail.From = New MailAddress("timeout_no_reply@outzource.dk")
        e_mail.To.Add(email)

        e_mail.Subject = "TimeOut - LM costcenter 600 - project joblog report"

        e_mail.IsBodyHtml = True
        e_mail.Body = "Hi TEST<br><br>"
        e_mail.Body += "The following projects has project hours last month:<br><br>"


        e_mail.Body += "<br><br>You can check hours by logging into TimeOut: <a href=""https://timeout.cloud/tia"">timeout.cloud/tia</a><br><br><br><br>"
        e_mail.Body += "This mail is sent from the TimeOut email service."

        'Dim FileName As String = "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_14\inc\upload\" & ltofolder & "\" & file

        'If System.IO.File.Exists(FileName) Then e_mail.Attachments.Add(New System.Net.Mail.Attachment(FileName))


        Smtp_Server.Send(e_mail)



        'Return "Mails Sent"

        '' etc...
    End Sub




</script>
   


<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>TimeOut Send TEST mail</title>
</head>


     
    <body style="font-family:Arial" leftmargin="20" topmargin="20">

    


  <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	
	</td></tr></table>
  
         <h4>TimeOut Send TEST mail</h4>

    <form id="form1" runat="server" method="post">
    <div style="position:relative; left:10px; top:40px; width:600px; padding:20px; border:1px #999999 solid;">
    <asp:TextBox runat="server" ID="fm_lto" Style="visibility:hidden;"></asp:TextBox>
    <asp:TextBox runat="server" ID="fm_thisuser" Style="visibility:hidden;"></asp:TextBox>
    <asp:TextBox runat="server" ID="test"></asp:TextBox>
   
        Send to specifik user: 


       
       

     <asp:Button runat="server" Text="Send test email now >>" ID="bt"  Style="float:left;"  /> <!-- OnClick="hentData(0,0,0,0)" -->

    <br /><br />

    <asp:Label runat="server" ID="meMTxt" Style="position:relative; width:300px; height:400px; vertical-align:top;">..</asp:Label>


<br /><br /><a href="#" onclick="Javascript:window.close();" style="color:red; font-size:12px; float:right;">[Close window X]</a>
    </div>
   
    
    
    </form>
</body>
</html>






        




