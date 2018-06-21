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



    Sub Page_load()



        Dim thisuser As String = Request("user")
        lto = Request("lto")
        fm_lto.Text = lto
        fm_thisuser.Text = thisuser

    End Sub






    Sub hentData()




        'Dim kategori As Integer
        Dim lto As String
        Dim emailto As String
        Dim limit As String
        Dim thisuser As String
        Dim flname As String = "xxx.csv"

        lto = fm_lto.Text
        limit = "" 'fm_limit.Value
        thisuser = fm_thisuser.Text
        'emailto = sendtospecificemail.Value
        emailto = Request("sendtospecificemail")

        Dim Table1 As DataTable
        Table1 = New DataTable("tb_to_var")


        Dim column1 As DataColumn = New DataColumn("lto")
        column1.DataType = System.Type.GetType("System.String")

        Table1.Columns.Add(column1)

        Dim column2 As DataColumn = New DataColumn("limit")
        column2.DataType = System.Type.GetType("System.String")

        Table1.Columns.Add(column2)

        Dim column3 As DataColumn = New DataColumn("emailto")
        column3.DataType = System.Type.GetType("System.String")

        Table1.Columns.Add(column3)

        Dim row As DataRow

        row = Table1.NewRow()
        row("lto") = lto
        row("limit") = limit
        row("emailto") = emailto

        'Add row
        Table1.Rows.Add(row)

        Dim ds As DataSet = New DataSet("ds")
        ds.Tables.Add(Table1)



        Dim CallWebService As New dk_rack.weekrapport_lto.ozreportws_lto()

        'Dim fname As Array = CallWebService.oz_report()
        CallWebService.Timeout = -1 '// Or -1 For infinite (ellers timer servicen ud ved ca. 70 mails)
        Dim a As String = CallWebService.oz_report_lto(ds)

        'meMTxt.Text = "Number of mails sent: " & a & "<br>"


        Dim ltofolder As String = lto

        If lto = "outz" Then
            lto = "intranet"
        End If


        Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" & lto & ";user=to_outzource2;Password=SKba200473;"

        '** Åbner Connection ***'
        Dim objConn As OdbcConnection
        objConn = New OdbcConnection(strConn)
        objConn.Open()

        Dim objCmd As OdbcCommand
        Dim objCmd2 As OdbcCommand
        'Dim objDataSet As New DataSet
        Dim objDR As OdbcDataReader
        Dim objDR2 As OdbcDataReader

        Dim t As Integer
        t = 0
        If t = 0 Then

            Dim strSQLfilenames As String = "SELECT afe_id, afe_email, afe_file FROM abonner_file_email WHERE afe_sent = 0 ORDER BY afe_id"

            Dim emailSendTo As String = ""
            objCmd = New OdbcCommand(strSQLfilenames, objConn)
            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

            While objDR.Read() = True


                Call sendmail(objDR("afe_email"), objDR("afe_file"), lto, ltofolder, thisuser)

                emailSendTo += "; " + objDR("afe_email")

                Dim strSQLtimerOverfort As String = "UPDATE abonner_file_email SET afe_sent = 1 WHERE afe_id = " & objDR("afe_id")

                objCmd2 = New OdbcCommand(strSQLtimerOverfort, objConn)
                objDR2 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)

            End While
            objDR.Close()





            meMTxt.Text = "TimeOut Week-report sent to: " + emailSendTo + "<br>Number of mails:" + a

        Else
            meMTxt.Text = "Rapport klar - no mail afsendt"
        End If














    End Sub


    Function sendmail(email, file, lto, ltofolder, thisuser) As String

        Dim Smtp_Server As New SmtpClient
        Dim e_mail As New MailMessage()
        'Smtp_Server.UseDefaultCredentials = True
        'Smtp_Server.Credentials = New Net.NetworkCredential("administrator", "Sok!2637")
        Smtp_Server.Port = 25 '587
        'Smtp_Server.EnableSsl = True
        Smtp_Server.Host = "formrelay.rackhosting.com"

        e_mail = New MailMessage()
        e_mail.From = New MailAddress("timeout_no_reply@outzource.dk")
        e_mail.To.Add(email)
        e_mail.Subject = "TimeOut - Weekreport"
        e_mail.IsBodyHtml = True
        e_mail.Body = "Hi TimeOut User.<br><br>"
        e_mail.Body += "This mail is sent from the TimeOut email service. You can change the settings for your subscriptions in TimeOut under TSA - administration - Subscribe.<br><br> This report is sent by the Timeout Email agent." '+ thisuser

        Dim FileName As String = "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_14\inc\upload\" & ltofolder & "\" & file

        If System.IO.File.Exists(FileName) Then e_mail.Attachments.Add(New System.Net.Mail.Attachment(FileName))


        'e_mail.Attachments.Add()

        Smtp_Server.Send(e_mail)
        'MsgBox("Mail Sent")

        Return "Mails Sent"

    End Function


</script>
   


<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>TimeOut Send week-report manual</title>
</head>


     
    <body style="font-family:Arial" leftmargin="20" topmargin="20">

    


  <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	
	</td></tr></table>
  
         <h4>TimeOut Send week-report manual</h4>

    <form id="form1" runat="server" method="post">
    <div style="position:relative; left:10px; top:40px; width:600px; padding:20px; border:1px #999999 solid;">
    <asp:TextBox runat="server" ID="fm_lto" Style="visibility:hidden;"></asp:TextBox>
    <asp:TextBox runat="server" ID="fm_thisuser" Style="visibility:hidden;"></asp:TextBox>
    <!--<asp:TextBox runat="server" ID="meid"></asp:TextBox>-->
    <!--
        <br /><br />
    
    If you got many subscribers, please select limits here: 
    <select runat="server" id="fm_limit">
        <option value=" LIMIT 0,50">First 0-49</option>
        <option value=" LIMIT 50,50">Next 50-99</option>
        <option value=" LIMIT 100,50">Next 100-149</option>
        <option value=" LIMIT 150,50">Next 150-199</option>
        <option value=" LIMIT 200,50">Next 200-249</option>
        <option value=" LIMIT 300,50">Next 250-299</option>
    </select>-->
    <br /><br />

        Send to specifik user: 


       
       

        <% 

            'Dim strSQLfilename1s As String = "SELECT email FROM rapport_abo WHERE lto = '" & lto & "' ORDER BY email"
            'Response.Write(strSQLfilename1s)



            Dim lto_db As String = ""
            If lto = "outz" Then
                lto_db = "intranet"
            Else
                lto_db = lto
            End If






            Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_admin;user=to_outzource2;Password=SKba200473;"

            '** Åbner Connection ***'
            Dim objConn As OdbcConnection
            objConn = New OdbcConnection(strConn)
            objConn.Open()

            Dim objCmd As OdbcCommand
            Dim objDR As OdbcDataReader

            Dim strSQLfilenames As String = "SELECT email FROM rapport_abo WHERE lto = '" & lto & "' ORDER BY email"
            objCmd = New OdbcCommand(strSQLfilenames, objConn)
            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)




            While objDR.Read() = True



                strOptions += "<option value=" & objDR("email") & ">" & objDR("email") & "</option>"



            End While
            objDR.Close()
            %>

          
           <select name="sendtospecificemail"  id="sendtospecificemail">
            <option value="all">All</option>
               <%=strOptions %>
            </select>

        <br /><br />

     <asp:Button runat="server" Text="Send week-report now >>" ID="bt" OnClick="hentData" Style="float:left;"  />

    <br /><br />

    <asp:Label runat="server" ID="meMTxt" Style="position:relative; width:300px; height:400px; vertical-align:top;">..</asp:Label>


<br /><br /><a href="#" onclick="Javascript:window.close();" style="color:red; font-size:12px; float:right;">[Close window X]</a>
    </div>
   
    
    
    </form>
</body>
</html>






        




