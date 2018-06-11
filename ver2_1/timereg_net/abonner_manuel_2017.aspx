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


        Dim dDato As Date = Date.Now 'tirsdag morgen 07:00
        Dim weekDay As Integer = DatePart("w", dDato, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        Dim weekDayTime As Integer = DatePart("h", dDato, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        test.Text = dDato & " # " & weekDay & " ## " & weekDayTime


        Dim func As String = Request("func") '= fromtimerservice 
        '*** Func = fromtimerservice Service bliver kaldt fra timer på server

        If (func = "fromtimerservice") Then

            '** 
            emailto = "all"

            Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_admin;user=to_outzource2;Password=SKba200473;"

            '** Åbner Connection ***'
            Dim objConn As OdbcConnection
            objConn = New OdbcConnection(strConn)
            objConn.Open()

            Dim objCmd As OdbcCommand
            'Dim objDataSet As New DataSet
            Dim objDR As OdbcDataReader








            'type_send = 0 = UGE
            'Henter abonneter der passer til Uge, Måned, Sidste dag i uge osv. 
            Dim rapportAboSQL As String = "SELECT lto, rapporttype, klokkeslet, ugedag, navn, type_send FROM abonner_raptype_tidspunkt WHERE type_send = 0 ORDER BY id"
            objCmd = New OdbcCommand(rapportAboSQL, objConn)
            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

            While objDR.Read() = True

                'TEST
                'PKT 2
                'If CInt(weekDay) = 5 And CInt(weekDayTime) = 11 And objDR("rapporttype") = 5 And objDR("lto") = "outz" Then
                'Call hentData("tia_20180509", emailto, objDR("rapporttype"), 1, 1)
                'End If

                'PKT. 3 liste til LM Monday 07.00 med dem der ikke har afsluttet - Special costcenters - KLAR TIL TEST
                'If CInt(weekDay) = 5 And CInt(weekDayTime) = 11 And objDR("rapporttype") = 9 And objDR("lto") = "outz" Then
                'Call hentData("tia_20180509", emailto, objDR("rapporttype"), 1, 9)
                'End If

                'PKT 4
                'If CInt(weekDay) = 2 And CInt(weekDayTime) = 14 And objDR("rapporttype") = 6 And objDR("lto") = "outz" Then
                'Call hentData("tia_20180509", emailto, objDR("rapporttype"), 1, 2)
                'End If

                'PKT. 5 ALL Timesheet default completed - KLAR TIL TEST
                'If CInt(weekDay) = 2 And CInt(weekDayTime) = 14 And objDR("rapporttype") = 8 And objDR("lto") = "outz" Then
                'Call defaultComplete("tia_20180509", 8)
                'End If

                'PKT. 7 LM REMINDER ALL LM who have not Approved WEEKLY. Monday 13.00 - Special costcenters
                'If CInt(weekDay) = 2 And CInt(weekDayTime) = 14 And objDR("rapporttype") = 10 And objDR("lto") = "outz" Then
                'Call hentData("tia_20180509", emailto, objDR("rapporttype"), 1, 10)
                'End If

                'PKT. 7 PM REMINDER ALL PM who have not Approved Project hours. Monday 13.00 - Special costcenters
                'If CInt(weekDay) = 7 And CInt(weekDayTime) = 10 And objDR("rapporttype") = 13 And objDR("lto") = "outz" Then
                'Call hentData("tia_20180509", emailto, objDR("rapporttype"), 1, 13)
                'End If


                'PKT 8 Second reminder is sent to SnS mangers Deadline: Monday 3pm
                'If CInt(weekDay) = 2 And CInt(weekDayTime) = 14 And objDR("rapporttype") = 12 And objDR("lto") = "outz" Then
                'Call hentData("tia_20180509", emailto, objDR("rapporttype"), 1, 12)
                'End If

                'PKT 9 Defalut godkend alle timer - KLAR TIL TEST
                'If CInt(weekDay) = 2 And CInt(weekDayTime) = 14 And objDR("rapporttype") = 11 And objDR("lto") = "outz" Then
                'Call defaultApprove("tia_20180509", 11)
                'End If



                'PRODUKTION

                'Der skal være EN rapport type 5 - KLAR TIL TEST


                'PKT. 3 liste til LM Sunday 23.59 med dem der ikke har afsluttet - Special costcenters - KLAR TIL TEST
                If CInt(weekDay) = 1 And CInt(weekDayTime) = 0 And objDR("rapporttype") = 9 And objDR("lto") = "tia" Then
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 9)
                End If


                'PKT. 2 FIRST reminder ALL monday 07.00 (ændret fra søndag 23:59)
                If CInt(weekDay) = 1 And CInt(weekDayTime) = 7 And objDR("rapporttype") = 5 And objDR("lto") = "tia" Then
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 1)
                End If



                'PKT. 4 SECOND reminder ALL in SnS (cost center 300, 320 & 330)  Monday 09.00 - KLAR TIL TEST
                If CInt(weekDay) = 1 And CInt(weekDayTime) = 9 And objDR("rapporttype") = 6 And objDR("lto") = "tia" Then
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 2)
                End If

                'PKT. 5 ALL Timesheet default completed - KLAR TIL TEST
                If CInt(weekDay) = 1 And CInt(weekDayTime) = 10 And objDR("rapporttype") = 8 And objDR("lto") = "tia" Then
                    Call defaultComplete(objDR("lto"), 8)
                End If


                'PKT. 7 REMINDER ALL PM + LM who have not Approved WEEKLY & Project hours. Monday 13.00 - Special costcenters
                If CInt(weekDay) = 1 And CInt(weekDayTime) = 13 And objDR("rapporttype") = 10 And objDR("lto") = "tia" Then
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 10)
                End If

                'PKT. 7 PM REMINDER ALL PM who have Not Approved Project hours. Monday 13.00 - Special costcenters
                If CInt(weekDay) = 1 And CInt(weekDayTime) = 13 And objDR("rapporttype") = 13 And objDR("lto") = "tia" Then
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 13)
                End If

                'PKT 8 Second reminder is sent to SnS mangers Deadline: Monday 3pm
                If CInt(weekDay) = 1 And CInt(weekDayTime) = 15 And objDR("rapporttype") = 12 And objDR("lto") = "tia" Then
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 12)
                End If

                'PKT 8 Second reminder is sent to SnS mangers Deadline: Monday 3pm
                If CInt(weekDay) = 1 And CInt(weekDayTime) = 15 And objDR("rapporttype") = 14 And objDR("lto") = "tia" Then
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 14)
                End If


                'PKT 9 Defalut godkend alle timer - KLAR TIL TEST
                If CInt(weekDay) = 1 And CInt(weekDayTime) = 16 And objDR("rapporttype") = 11 And objDR("lto") = "tia" Then
                    Call defaultApprove(objDR("lto"), 11)
                End If

                'Pkt 10 Sender godkendte til NAV

            End While
            objDR.Close()


        End If


        Dim thisuser As String = Request("user")
        lto = Request("lto")
        fm_lto.Text = lto
        fm_thisuser.Text = thisuser

    End Sub


    Private Sub btnOverRide_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bt.Click

        'test.Text = "HEj"
        'Dim retval = SomeFunction(value)
        lto = "outz"
        emailto = "sk@outzource.dk"
        Call hentData(lto, emailto, 5, 1, 0)


        '' etc...
    End Sub



    Sub hentData(lto, emailto, rapporttype, manual_auto, reminder_rule_no)


        'Response.End()



        'Dim kategori As Integer

        Dim limit As String
        Dim thisuser As String
        Dim flname As String = "xxx.csv"

        Dim this_rapporttype As String = rapporttype
        Dim this_manauto As Integer = manual_auto '1 Auto, 0 Manual

        If CInt(manual_auto) = 1 Then 'auto
            lto = lto 'fm_lto.Text
            limit = "" 'fm_limit.Value
            thisuser = "TimeOut user"
        Else
            lto = fm_lto.Text
            limit = fm_limit.Value
            thisuser = fm_thisuser.Text
        End If



        'emailto = sendtospecificemail.Value
        emailto = "all" 'Request("sendtospecificemail")

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

        Dim column4 As DataColumn = New DataColumn("rapporttype")
        column4.DataType = System.Type.GetType("System.String")

        Table1.Columns.Add(column4)

        Dim column5 As DataColumn = New DataColumn("manuel_auto")
        column5.DataType = System.Type.GetType("System.String")

        Table1.Columns.Add(column5)

        Dim row As DataRow

        row = Table1.NewRow()
        row("lto") = lto
        row("limit") = limit
        row("emailto") = emailto
        row("rapporttype") = this_rapporttype
        row("manuel_auto") = this_manauto

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


        Dim strSQLtimerOverfort As String = ""
        Dim t As Integer
        t = 0
        If t = 0 Then



            Dim strSQLfilenames As String = "SELECT afe_id, afe_email, afe_file, mnavn, init, afe_progrp FROM abonner_file_email"
            Select Case rapporttype
                Case 0, 1, 2, 3, 4
                    strSQLfilenames += " LEFT JOIN medarbejdere ON (email = afe_email AND mansat = 1)"
                Case 5, 6, 9, 10, 12, 13, 14
                    strSQLfilenames += " LEFT JOIN medarbejdere ON (init = afe_init AND mansat = 1)"
                Case Else
                    strSQLfilenames += " LEFT JOIN medarbejdere ON (email = afe_email AND mansat = 1)"

            End Select

            Select Case rapporttype
                Case 0, 1, 2, 3, 4, 5, 6
                    strSQLfilenames += " WHERE afe_sent = 0 AND mansat = 1 GROUP BY init ORDER BY afe_id"
                Case 9, 10, 12, 13, 14
                    strSQLfilenames += " WHERE afe_sent = 0 AND mansat = 1 ORDER BY afe_id"
                Case Else
                    strSQLfilenames += " WHERE afe_sent = 0 AND mansat = 1 GROUP BY init ORDER BY afe_id"

            End Select



            Dim emailSendTo As String = ""
            objCmd = New OdbcCommand(strSQLfilenames, objConn)
            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)



            While objDR.Read() = True

                Select Case this_rapporttype
                    Case 5, 6 'TIA WEEKLY

                        thisuser = objDR("mnavn") & " [" & objDR("init") & "]"

                        Call sendmailWeeklyNotify(objDR("afe_email"), objDR("afe_file"), lto, ltofolder, thisuser, reminder_rule_no)

                    Case 9 'TIA WEEKLY LM - NOT Closed Weeks

                        thisuser = objDR("mnavn") & " [" & objDR("init") & "]"

                        Call sendmailWeeklyNotifyLM(objDR("afe_email"), objDR("afe_file"), lto, ltofolder, thisuser, reminder_rule_no, objDR("afe_progrp"))

                    Case 10, 12, 13, 14 'TIA WEEKLY LM (10,12) +PM (13,14) NOT Approved hours 13:00 + 15:00 Sns 300

                        thisuser = objDR("mnavn") & " [" & objDR("init") & "]"

                        Call sendmailWeeklyNotifyLM(objDR("afe_email"), objDR("afe_file"), lto, ltofolder, thisuser, reminder_rule_no, objDR("afe_progrp"))
                    Case Else
                        Call sendmail(objDR("afe_email"), objDR("afe_file"), lto, ltofolder, thisuser)
                End Select



                emailSendTo += "; " & objDR("afe_id") & ". " & objDR("afe_email")

                strSQLtimerOverfort = "UPDATE abonner_file_email SET afe_sent = 1 WHERE afe_id = " & objDR("afe_id")

                objCmd2 = New OdbcCommand(strSQLtimerOverfort, objConn)
                objDR2 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)

            End While
            objDR.Close()



            Dim strSQLtimerOverfortm As String = "UPDATE abonner_file_email SET afe_sent = -1 WHERE afe_sent = 0"

            objCmd2 = New OdbcCommand(strSQLtimerOverfortm, objConn)
            objDR2 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)


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
        e_mail.Body += "This mail is sent from the TimeOut email service. You can change the settings for your subscriptions in TimeOut under TSA - administration - Subscribe." '+ thisuser

        Dim FileName As String = "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_14\inc\upload\" & ltofolder & "\" & file

        If System.IO.File.Exists(FileName) Then e_mail.Attachments.Add(New System.Net.Mail.Attachment(FileName))


        'e_mail.Attachments.Add()

        Smtp_Server.Send(e_mail)
        'MsgBox("Mail Sent")

        Return "Mails Sent"

    End Function



    Function sendmailWeeklyNotify(email, file, lto, ltofolder, thisuser, reminder_rule_no) As String

        Dim thisDateDatformat As String = ""
        Dim ddnow As Date = Date.Now
        Dim dd As Date = Date.Now
        Dim ddw As Integer = DatePart("w", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)

        If ddw = 1 Then
            dd = DateAdd("d", -1, dd)
        Else
            dd = DateAdd("d", -(ddw), dd)
        End If


        Dim thisDate As String = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)
        Dim thisWeek As String = DatePart("ww", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)

        Dim deadline As String = "Sunday 23:59"



        If CInt(reminder_rule_no) = 1 Or CInt(reminder_rule_no) = 2 Then


            dd = dd 'DateAdd("d", -1, dd)
            thisDateDatformat = DatePart("d", dd) & "." & DatePart("m", dd) & "." & DatePart("yyyy", dd)
            thisDate = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)
            thisWeek = DatePart("ww", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)

            deadline = "Sunday " & thisDateDatformat & " at 23:59 (11:59 PM)"

        End If


        Dim Smtp_Server As New SmtpClient
        Dim e_mail As New MailMessage()
        'Smtp_Server.UseDefaultCredentials = True
        'Smtp_Server.Credentials = New Net.NetworkCredential("administrator", "Sok!2637")
        Smtp_Server.Port = 25 '587
        'Smtp_Server.EnableSsl = True
        Smtp_Server.Host = "formrelay.rackhosting.com"

        'TEST
        If email = "" Then 'Or email = "raminta.ivoskute@nordgain.com"
            email = "sk@outzource.dk"
        End If


        e_mail = New MailMessage()
        e_mail.From = New MailAddress("timeout_no_reply@outzource.dk")
        e_mail.To.Add(email)





        If CInt(reminder_rule_no) = 1 Then
            e_mail.Subject = "TimeOut - Close Week First Reminder"
            e_mail.IsBodyHtml = True
            e_mail.Body = "Hi " & thisuser & "<br>"
        End If

        If CInt(reminder_rule_no) = 2 Then
            e_mail.Subject = "TimeOut - Close Week Second Reminder"
            e_mail.IsBodyHtml = True
            e_mail.Body = "Hi " & thisuser & "<br>"
            'e_mail.Body += "This is your second reminder. All your work will be autocompleted 10:00.<br>"
        End If


        e_mail.Body += "You have not completed your week no. " & thisWeek & " before deadline: " & deadline & ", please do it immediately.<br><br>"

        If CInt(reminder_rule_no) = 2 Then
            e_mail.Body += "Notice: This Is the second reminder; all time registration will be autocompleted at 10.00 AM DK Time!<br><br>"
        End If

        e_mail.Body += "You can complete your week by logging into TimeOut: <a href=""https://timeout.cloud/tia"">timeout.cloud/tia</a><br><br><br><br>"
        e_mail.Body += "This mail is sent from the TimeOut email service. " & ddnow & "" '+ thisuser

        'Dim FileName As String = "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_14\inc\upload\" & ltofolder & "\" & file

        'If System.IO.File.Exists(FileName) Then e_mail.Attachments.Add(New System.Net.Mail.Attachment(FileName))

        'If email = "sk@outzource.dk" Then
        Smtp_Server.Send(e_mail)
        'End If
        'MsgBox("Mail Sent")

        Return "Mails Sent"

    End Function



    Function sendmailWeeklyNotifyLM(email, file, lto, ltofolder, thisuser, reminder_rule_no, progrpId) As String


        ' Using writer As StreamWriter = New StreamWriter("D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_14\inc\upload\" & lto & "\" & flname9, False, Encoding.GetEncoding("iso-8859-1"))


        'ExpTxtheader = "Project;Project No.;Activity;Employee;Init;Forcast Hours Total;Real. Hours Month;Saldo;Forcast Hours " & MonthName(Month(dDato)) & " - " & Year(dDato) & ";Real. Hours " & MonthName(Month(dDato)) & " - " & Year(dDato) & ";"
        'writer.WriteLine(ExpTxtheader)
        'writer.WriteLine(objDR.Item("jobnavn") & ";" & objDR.Item("jobnr") & ";" & objDR2.Item("aktnavn") & ";" & expTxt)

        'End Using 'writer


        'reminder_rule_no = 9


        Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" & lto & ";user=to_outzource2;Password=SKba200473;"

        '** Åbner Connection ***'
        Dim objConn As OdbcConnection
        objConn = New OdbcConnection(strConn)
        objConn.Open()

        Dim objCmd As OdbcCommand
        'Dim objDataSet As New DataSet

        Dim objDR2 As OdbcDataReader
        Dim objDR3 As OdbcDataReader
        Dim objDR4 As OdbcDataReader

        Dim ddnow As Date = Date.Now
        Dim dd As Date = Date.Now
        Dim ddw As Integer = DatePart("w", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)


        'If CInt(reminder_rule_no) = 9 Then 'udsendes søndag

        'dd = dd

        'Else


        If ddw = 1 Then
            dd = DateAdd("d", -1, dd)
        Else
            dd = DateAdd("d", -(ddw), dd)
        End If


        'End If


        Dim thisDate As String = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)
        Dim thisWeek As String = DatePart("ww", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)

        Dim deadline As String = "Sunday 23:59"



        'Dim thisDate As String = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)
        'Dim thisWeek As String = DatePart("ww", dd)


        'Dim dgsondag As Date = Date.Now
        'dgsondag = dd 'DateAdd("d", -1, dd)
        'Dim slutDatoWeekSQL As String = DatePart("yyyy", dgsondag) & "-" & DatePart("m", dgsondag) & "-" & DatePart("d", dgsondag)

        Dim empNotCompleted As String = ""
        If CInt(reminder_rule_no) = 9 Then
            'empNotCompleted = "Employees not completed week no. " & thisWeek & ", it should be closed before: " & deadline & ".<br>"
            empNotCompleted = "The following employee/s from your team has Not completed their week no. " & thisWeek & " before deadline: Sunday 11.59 PM DK Time.<br>"
        End If

        If CInt(reminder_rule_no) = 10 Then
            empNotCompleted = "You have not approved all weeks for the employees in your team before deadline: today 1.00 PM DK Time.<br><br>" 'Employees hours/weeks not Approved:<br>"
            'empNotCompleted += "The following employee/s from your team still has non-approved weeks:<br><br>"
        End If

        If CInt(reminder_rule_no) = 12 Then
            empNotCompleted = "You have Not approved all weeks for the employees in your team before deadline: Today 3.0 PM DK Time.<br><br>"
            'empNotCompleted += "The following employee/s from your team still has non-approved weeks:<br><br>"
        End If


        Dim antalEmails As Integer
        Dim weekclosedbyuser As Integer = 0

        Dim strSQLprogrpids As String = ""
        Dim strSQLjob As String = ""
        Dim strSQL100400700 As String = ""
        Dim medidMember100400700 As Integer = 0
        Dim medidHours100400700 As Integer = 0

        Dim ddSondayLastWeek As Date = Date.Now
        ddSondayLastWeek = dd 'DateAdd("d", -(ddw), ddSondayLastWeek)
        Dim ddSondayLastWeekSQL As String = DatePart("yyyy", ddSondayLastWeek) & "-" & DatePart("m", ddSondayLastWeek) & "-" & DatePart("d", ddSondayLastWeek)

        Dim ddMondayLastWeek As Date = Date.Now
        ddMondayLastWeek = DateAdd("d", -(7), dd) 'DateAdd("d", -(ddw + 6), ddMondayLastWeek)
        Dim ddMondayLastWeekSQL As String = DatePart("yyyy", ddMondayLastWeek) & "-" & DatePart("m", ddMondayLastWeek) & "-" & DatePart("d", ddMondayLastWeek)

        empNotCompleted += "Period: " & DatePart("yyyy", ddMondayLastWeek) & "-" & DatePart("m", ddMondayLastWeek) & "-" & DatePart("d", ddMondayLastWeek) & " And " & DatePart("yyyy", ddSondayLastWeek) & "-" & DatePart("m", ddSondayLastWeek) & "-" & DatePart("d", ddSondayLastWeek) & "<br><br>"


        Dim hoursAppStatusSQL As String = ""
        Dim weekAppStatusSQL As String = ""
        Dim strSQLugeafsluttet As String = ""

        If CInt(reminder_rule_no) = 10 Or CInt(reminder_rule_no) = 12 Or CInt(reminder_rule_no) = 13 Or CInt(reminder_rule_no) = 14 Then
            hoursAppStatusSQL = " And (godkendtstatus = 0 Or godkendtstatus = 3)" 'Ikke taget stilling / Tentative
            weekAppStatusSQL = " And ugegodkendt = 0 AND splithr = 0"
        End If


        Select Case CInt(reminder_rule_no)
            Case 9, 10, 12

                strSQLprogrpids = "SELECT pr.id AS pid, m.mnavn, init, m.email, mid As medid, teamleder, med_cal FROM progrupperelationer pr "
                strSQLprogrpids += " Left Join medarbejdere m ON (m.mid = pr.medarbejderid And mansat = 1) "
                strSQLprogrpids += " WHERE pr.projektgruppeid = " & progrpId & " And teamleder = 0 And mansat = 1 "
                strSQLprogrpids += " ORDER BY pr.id LIMIT 2000"

            Case 13, 14

                strSQLprogrpids = "SELECT pr.id AS pid, m.mnavn, init, m.email, mid As medid, teamleder, med_cal FROM progrupperelationer pr "
                strSQLprogrpids += " Left Join medarbejdere m ON (m.mid = pr.medarbejderid And mansat = 1) "
                strSQLprogrpids += " WHERE pr.projektgruppeid = " & progrpId & " And teamleder = 1 And mansat = 1 "
                strSQLprogrpids += " ORDER BY pr.id LIMIT 20"


        End Select

        Dim lastPid As Integer = 0




        objCmd = New OdbcCommand(strSQLprogrpids, objConn)
        objDR2 = objCmd.ExecuteReader
        While objDR2.Read()



            weekclosedbyuser = 0
            medidMember100400700 = 0
            medidHours100400700 = 0
            'PM 13, 14
            If ((CInt(reminder_rule_no) = 13 Or CInt(reminder_rule_no) = 14) And progrpId <> 16) Then
                '13+14: NOT PM in DEV progrpId = 16


                medidMember100400700 = 1

                strSQLjob = "SELECT j.id, j.jobnavn, j.jobnr, k.kkundenavn, k.kkundenr FROM job j "
                strSQLjob += "LEFT Join kunder k ON (kid = jobknr) WHERE (j.projektgruppe1 = " & progrpId & " Or j.projektgruppe2 = " & progrpId & " OR j.projektgruppe3 = " & progrpId & " OR j.projektgruppe4 = " & progrpId & " OR j.projektgruppe5 = " & progrpId
                strSQLjob += " Or j.projektgruppe6 = " & progrpId & " Or j.projektgruppe7 = " & progrpId & " OR j.projektgruppe8 = " & progrpId & " OR j.projektgruppe9 = " & progrpId & " OR j.projektgruppe10 = " & progrpId & ") AND (jobans1 = " & objDR2("medid") & " OR jobans2 = " & objDR2("medid") & ")"

                objCmd = New OdbcCommand(strSQLjob, objConn)
                objDR3 = objCmd.ExecuteReader
                While objDR3.Read()

                    strSQL100400700 = "Select SUM(timer) As timer FROM timer WHERE tjobnr = '" & objDR3("jobnr") & "' And tdato BETWEEN '" & ddMondayLastWeekSQL & "' AND '" & ddSondayLastWeekSQL & "' " & hoursAppStatusSQL & " AND (tfaktim = 1 OR tfaktim = 2) GROUP BY tjobnr"
                    objCmd = New OdbcCommand(strSQL100400700, objConn)
                    objDR4 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    If objDR4.Read() Then

                        medidHours100400700 = objDR4("timer")

                        If medidHours100400700 <> 0 Then
                            empNotCompleted += objDR3("kkundenavn") & " (" & objDR3("kkundenr") & ") Job: " & objDR3("jobnavn") & " (" & objDR3("jobnr") & ") Hours: " & FormatNumber(objDR4("timer"), 2) & "<br>" ' Employee: " & objDR2.Item("mnavn") & " [" & objDR2.Item("init") & "] hours: " & objDR4("timer") & "<br>"
                            antalEmails = antalEmails + 1
                        End If

                    End If
                    objDR4.Close()


                End While
                objDR3.Close()

            End If



            'LM 9,10,12
            If (CInt(reminder_rule_no) = 9 And objDR2("med_cal") <> "LT" And (progrpId = 11 Or progrpId = 15 Or progrpId = 17)) _
            Or (CInt(reminder_rule_no) = 10 And progrpId <> 16) _
            Or (CInt(reminder_rule_no) = 12 And (progrpId = 12 Or progrpId = 13 Or progrpId = 14)) Then
                'Employees in Commercial, HR and Finance (cost center 100, 400 & 700)  ONLY IF THEY GOT HOURS
                '10: NOT LM in DEV progrpId = 16


                medidMember100400700 = 1

                strSQL100400700 = "Select SUM(timer) As timer FROM timer WHERE tmnr = " & objDR2("medid") & " And tdato BETWEEN '" & ddMondayLastWeekSQL & "' AND '" & ddSondayLastWeekSQL & "' " & hoursAppStatusSQL & " GROUP BY tmnr"
                objCmd = New OdbcCommand(strSQL100400700, objConn)
                objDR4 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                If objDR4.Read() Then

                    medidHours100400700 = objDR4("timer")

                End If

                objDR4.Close()


            End If



            '*** HEnter medarbejdere der ikke har afsluttet WEEK i grupppen
            If CInt(reminder_rule_no) = 9 Or CInt(reminder_rule_no) = 10 Or CInt(reminder_rule_no) = 12 Then

                strSQLugeafsluttet = "Select mid FROM ugestatus WHERE mid = " & objDR2("medid") & " And (uge = '" & ddSondayLastWeekSQL & "') " & weekAppStatusSQL & ""
                objCmd = New OdbcCommand(strSQLugeafsluttet, objConn)
                objDR4 = objCmd.ExecuteReader
                weekclosedbyuser = 0
                If objDR4.Read() = True Then
                    weekclosedbyuser = 1

                End If
                objDR4.Close()

            End If


            If CInt(reminder_rule_no) = 9 Then

                If (CInt(weekclosedbyuser) = 0 And CInt(medidMember100400700) = 0) Or (CInt(weekclosedbyuser) = 0 And CInt(medidMember100400700) = 1 And medidHours100400700 > 0) Then

                    antalEmails = antalEmails + 1
                    empNotCompleted += objDR2.Item("mnavn") & " [" & objDR2.Item("init") & "] " & objDR2.Item("email") & "<br>"

                End If


            End If


            If CInt(reminder_rule_no) = 10 Then 'LM not approved

                If (CInt(weekclosedbyuser) = 1 And medidMember100400700 = 1) Then

                    antalEmails = antalEmails + 1
                    empNotCompleted += objDR2.Item("mnavn") & " [" & objDR2.Item("init") & "] " & objDR2.Item("email") & " <br>" '" & strSQLugeafsluttet & "

                End If


            End If


            If CInt(reminder_rule_no) = 12 Then 'LM not approved Second SnS

                If (CInt(weekclosedbyuser) = 1 And medidMember100400700 = 1) Then

                    antalEmails = antalEmails + 1
                    empNotCompleted += objDR2.Item("mnavn") & " [" & objDR2.Item("init") & "] " & objDR2.Item("email") & "<br>"

                End If


            End If



        End While
        objDR2.Close()


        If antalEmails > 0 Then

            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            'Smtp_Server.UseDefaultCredentials = True
            'Smtp_Server.Credentials = New Net.NetworkCredential("administrator", "Sok!2637")
            Smtp_Server.Port = 25 '587
            'Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "formrelay.rackhosting.com"

            'TEST
            'email = "sk@outzource.dk"

            e_mail = New MailMessage()
            e_mail.From = New MailAddress("timeout_no_reply@outzource.dk")
            e_mail.To.Add(email)

            If CInt(reminder_rule_no) = 9 Then
                e_mail.Subject = "TimeOut - LM - costcenter members complete Week reminder"
            End If

            If CInt(reminder_rule_no) = 10 Then
                e_mail.Subject = "TimeOut - LM - weeks not Approved"
            End If

            If CInt(reminder_rule_no) = 12 Then
                e_mail.Subject = "TimeOut - LM - weeks not Approved second reminder"
            End If


            If CInt(reminder_rule_no) = 13 Then
                e_mail.Subject = "TimeOut - PM - projecthours not Approved"
            End If

            If CInt(reminder_rule_no) = 14 Then
                e_mail.Subject = "TimeOut - PM - projecthours not Approved - second reminder"
            End If

            e_mail.IsBodyHtml = True
            e_mail.Body = "Hi " & thisuser & "<br><br>"


            Select Case CInt(reminder_rule_no)
                Case 10, 12
                    e_mail.Body += "The following employee/s from your team still has non-approved weeks:<br><br>"
                Case 13, 14
                    e_mail.Body += "The following projects still has non-approved project hours:<br><br>"

            End Select



            e_mail.Body += empNotCompleted


            Select Case CInt(reminder_rule_no)
                Case 999
                    e_mail.Body += "<br><br>Please make sure to complete all weeks immediately.<br><br>"
                Case 10, 12
                    e_mail.Body += "<br><br>Please make sure to approve all weeks immediately.<br><br>"
                Case 13, 14
                    e_mail.Body += "<br><br>Please make sure to approve all project hours immediately.<br><br>"
            End Select

            If CInt(reminder_rule_no) = 10 Or CInt(reminder_rule_no) = 13 Then
                e_mail.Body += "<b>Notice : All time registration will be automatically approved by 4.00 PM DK Time!</b>"
            End If


            If CInt(reminder_rule_no) = 12 Or CInt(reminder_rule_no) = 14 Then
                e_mail.Body += "<b>Notice: This is your second reminder and all-time registration will be automatically approved by 4.00 PM DK Time!</b><br><br>"
            End If

            Select Case CInt(reminder_rule_no)
                Case 9
                    e_mail.Body += "<br><br>You can complete weeks by logging into TimeOut: <a href=""https://timeout.cloud/tia"">timeout.cloud/tia</a><br><br><br><br>"
                Case 10, 12
                    e_mail.Body += "<br><br>You can approve weeks by logging into TimeOut: <a href=""https://timeout.cloud/tia"">timeout.cloud/tia</a><br><br><br><br>"
                Case 13, 14
                    e_mail.Body += "<br><br>You can approve hours by logging into TimeOut: <a href=""https://timeout.cloud/tia"">timeout.cloud/tia</a><br><br><br><br>"
            End Select


            e_mail.Body += "This mail is sent from the TimeOut email service. " & ddnow & "" '+ thisuser

            'Dim FileName As String = "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_14\inc\upload\" & ltofolder & "\" & file

            'If System.IO.File.Exists(FileName) Then e_mail.Attachments.Add(New System.Net.Mail.Attachment(FileName))

            Smtp_Server.Send(e_mail)
            'MsgBox("Mail Sent")



        End If

        Return "Mails Sent"

    End Function




    Function defaultComplete(lto, reminder_rule_no) As String


        Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" & lto & ";user=to_outzource2;Password=SKba200473;"

        '** Åbner Connection ***'
        Dim objConn As OdbcConnection
        objConn = New OdbcConnection(strConn)
        objConn.Open()

        Dim objCmd As OdbcCommand
        'Dim objDataSet As New DataSet

        Dim objDR2 As OdbcDataReader
        Dim objDR4 As OdbcDataReader

        Dim dd As Date = Date.Now
        Dim thisDateSQL As String = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd) & " " & DatePart("h", dd) & ":" & DatePart("n", dd) & ":00"

        dd = DateAdd("d", -1, dd)
        Dim thisWeek As String = DatePart("ww", dd)

        Dim lastWeekSundaySQL As String = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)

        dd = DateAdd("d", -6, dd)
        Dim lastWeekMondaySQL As String = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)


        Dim deadline As String = "Monday 09:00"



        If CInt(reminder_rule_no) = 8 Then 'Sætter specille for rapport 8. Autocomplte alle medarbejderes uge + timer mandag 10.00. pkt 5.



        End If


        Dim strSQLmedids As String = "SELECT mid AS medid, email, mnavn AS navn FROM medarbejdere WHERE mansat = 1 ORDER BY mid LIMIT 2000"
        objCmd = New OdbcCommand(strSQLmedids, objConn)
        objDR2 = objCmd.ExecuteReader
        While objDR2.Read()


            Dim strSQLugeafsluttet As String = "Select mid FROM ugestatus WHERE mid = " & objDR2.Item("medid") & " AND uge = '" & lastWeekSundaySQL & "' GROUP BY mid"
            objCmd = New OdbcCommand(strSQLugeafsluttet, objConn)
            objDR4 = objCmd.ExecuteReader
            Dim weekclosedbyuser As Integer = 0
            If objDR4.Read() = True Then
                weekclosedbyuser = 1

            End If
            objDR4.Close()


            If CInt(weekclosedbyuser) = 0 Then

                Dim intStatusAfs As Integer = 2 '** Afsl. forsent


                '**** SLETTER (Opdaterer for revisionsspor) EVT HRsplit midlertidig godkendt WEEK ***
                '**** SPLIT HR bruger altid datoen fra den sidste uge i måneden. 
                '**** FULD uge godkendes igen fra WEEK approval
                Dim strSQLafslutDelHRsplit As String = "UPDATE ugestatus SET uge = '2002-01-01' WHERE WEEK(uge, 3) = WEEK('" & lastWeekSundaySQL & "', 3) AND year(uge) = '" & Year(lastWeekSundaySQL) & "' AND mid = " & objDR2.Item("medid") & " AND splithr = 1"
                objCmd = New OdbcCommand(strSQLafslutDelHRsplit, objConn)
                objDR4 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                objDR4.Close()


                Dim dgs As Date = Date.Now
                Dim strSQLupdateugestatus As String = "INSERT INTO ugestatus (uge, afsluttet, mid, status) VALUES ('" & lastWeekSundaySQL & "', '" & thisDateSQL & "', " & objDR2.Item("medid") & ", " & intStatusAfs & ")"
                objCmd = New OdbcCommand(strSQLupdateugestatus, objConn)
                objDR4 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                objDR4.Close()

                Dim godkendtstatus As Integer = 3



                Dim strSQLupdatetimer As String = "Update timer Set godkendtstatus = " & godkendtstatus & ", godkendtstatusaf = 'TimeOut close Week service' WHERE"
                strSQLupdatetimer += " tmnr = " & objDR2.Item("medid") & " AND tdato BETWEEN '" & lastWeekMondaySQL & "' AND '" & lastWeekSundaySQL & "' AND godkendtstatus <> 1 AND overfort = 0"
                objCmd = New OdbcCommand(strSQLupdatetimer, objConn)
                objDR4 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                objDR4.Close()



            End If



        End While
        objDR2.Close()

        Return "All weeks autocompleted"

    End Function


    Function defaultApprove(lto, reminder_rule_no) As String


        Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" & lto & ";user=to_outzource2;Password=SKba200473;"

        '** Åbner Connection ***'
        Dim objConn As OdbcConnection
        objConn = New OdbcConnection(strConn)
        objConn.Open()

        Dim objCmd As OdbcCommand
        'Dim objDataSet As New DataSet

        Dim objDR2 As OdbcDataReader
        Dim objDR3 As OdbcDataReader
        Dim objDR4 As OdbcDataReader

        Dim dd As Date = Date.Now
        Dim thisDateSQL As String = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)

        dd = DateAdd("d", -1, dd)
        Dim thisWeek As String = DatePart("ww", dd)

        Dim lastWeekSundaySQL As String = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)

        dd = DateAdd("d", -7, dd)
        Dim lastWeekMondaySQL As String = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)


        Dim deadline As String = "Monday 16:00"



        If CInt(reminder_rule_no) = 11 Then 'Sætter specille for rapport 11. 



        End If





        Dim strSQLmedids As String = "SELECT mid AS medid, email, mnavn AS navn FROM medarbejdere WHERE mansat = 1 ORDER BY mid LIMIT 2000"
        objCmd = New OdbcCommand(strSQLmedids, objConn)
        objDR2 = objCmd.ExecuteReader
        While objDR2.Read()


            Dim strSQLugeafsluttet As String = "Select id, mid FROM ugestatus WHERE mid = " & objDR2.Item("medid") & " AND uge = '" & lastWeekSundaySQL & "' AND ugegodkendt = 0 GROUP BY mid"
            objCmd = New OdbcCommand(strSQLugeafsluttet, objConn)
            objDR4 = objCmd.ExecuteReader

            If objDR4.Read() = True Then


                '**** SLETTER (Opdaterer for revisionsspor) EVT HRsplit midlertidig godkendt WEEK ***
                '**** SPLIT HR bruger altid datoen fra den sidste uge i måneden. 
                '**** FULD uge godkendes igen fra WEEK approval
                Dim strSQLafslutDelHRsplit As String = "UPDATE ugestatus SET uge = '2002-01-01' WHERE WEEK(uge, 3) = WEEK('" & lastWeekSundaySQL & "', 3) AND year(uge) = '" & Year(lastWeekSundaySQL) & "' AND mid = " & objDR4("mid") & " AND splithr = 1"
                objCmd = New OdbcCommand(strSQLafslutDelHRsplit, objConn)
                objDR4 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                objDR4.Close()


                Dim dgs As Date = Date.Now
                Dim strSQLupdateugestatus As String = "UPDATE ugestatus SET ugegodkendt = 1, ugegodkendtaf = 1, ugegodkendtdt = '" & thisDateSQL & "' WHERE id = " & objDR4("id")
                objCmd = New OdbcCommand(strSQLupdateugestatus, objConn)
                objDR3 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                objDR3.Close()






            End If



        End While
        objDR2.Close()


        Dim godkendtstatus As Integer = 1

        Dim strSQLupdatetimer As String = "Update timer Set godkendtstatus = " & godkendtstatus & ", godkendtstatusaf = 'TimeOut Approve Week service', godkendtdato = '" & thisDateSQL & "' WHERE"
        strSQLupdatetimer += " tdato BETWEEN '" & lastWeekMondaySQL & "' AND '" & lastWeekSundaySQL & "' AND godkendtstatus <> 1 AND overfort = 0" 'tmnr = " & objDR2.Item("medid") & " AND
        objCmd = New OdbcCommand(strSQLupdatetimer, objConn)
        objDR3 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
        objDR3.Close()

        Return "All weeks and hours auto approved"

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
    <asp:TextBox runat="server" ID="test"></asp:TextBox>
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

        <%
            ' Dim lto As String = "outz"
            Dim emailtoAll As String = "all"

            Dim rapporttype As Integer = 0%>

         <%=lto %>,<%=emailtoAll%>,<%=rapporttype %>,0

     <asp:Button runat="server" Text="Send week-report now >>" ID="bt"  Style="float:left;"  /> <!-- OnClick="hentData(0,0,0,0)" -->

    <br /><br />

    <asp:Label runat="server" ID="meMTxt" Style="position:relative; width:300px; height:400px; vertical-align:top;">..</asp:Label>


<br /><br /><a href="#" onclick="Javascript:window.close();" style="color:red; font-size:12px; float:right;">[Close window X]</a>
    </div>
   
    
    
    </form>
</body>
</html>






        




