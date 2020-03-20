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
        Dim firstDayInmonth = Day(dDato)
        Dim weekDay As Integer = DatePart("w", dDato, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        Dim thisWeekDay As Integer = 1
        Dim thisWeekDayPlan_1 As Integer = 1
        Dim thisWeekDayPlan_2 As Integer = 2
        Dim thisWeekDayTest As Integer = 4
        Dim weekDayTime As Integer = DatePart("h", dDato, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        test.Text = dDato & " # " & weekDay & " ## " & weekDayTime


        Dim thisuser As String = Request("user")
        lto = Request("lto")
        fm_lto.Text = lto
        fm_thisuser.Text = thisuser





        Dim func As String = Request("func") '= fromtimerservice 

        '*** Func = fromtimerservice Service bliver kaldt fra timer p� server
        If (func = "fromtimerservice") Then 'Reminder Service auto (bl.a TIA)

            '** 
            emailto = "all"

            Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_admin;user=to_outzource2;Password=SKba200473;"

            '** �bner Connection ***'
            Dim objConn As OdbcConnection
            objConn = New OdbcConnection(strConn)
            objConn.Open()

            Dim objCmd As OdbcCommand
            'Dim objDataSet As New DataSet
            Dim objDR As OdbcDataReader
            Dim objDR2 As OdbcDataReader






            'type_send = 0 = UGE
            'Henter abonneter der passer til Uge, M�ned, Sidste dag i uge osv. 
            Dim rapportAboSQL As String = "SELECT lto, rapporttype, klokkeslet, ugedag, navn, type_send FROM abonner_raptype_tidspunkt WHERE type_send = 0 ORDER BY id"
            objCmd = New OdbcCommand(rapportAboSQL, objConn)
            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            Dim antalloops As Integer = 0
            While objDR.Read() = True



                '*** tjekker om det er en helligdag
                Dim tjekdennedagSQL As Date = Year(dDato) & "-" & Month(dDato) & "-" & Day(dDato)
                Dim erHellig As Integer = 0
                Dim helligdagnavn As String = ""

                If (lto = "tia") And CInt(antalloops) = 0 Then


                    Dim strConn_tia As String = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_tia;user=to_outzource2;Password=SKba200473;"

                    '** �bner Connection ***'
                    Dim objConn_tia As OdbcConnection
                    objConn_tia = New OdbcConnection(strConn_tia)
                    objConn_tia.Open()


                    Dim strSQLhellig As String = "SELECT nh_id, nh_country, nh_name, nh_duration, nh_date, nh_editor_date, nh_open, nh_sortorder, nh_projgrp FROM national_holidays "
                    strSQLhellig += " WHERE nh_id <> 0 AND nh_date = '" & tjekdennedagSQL & "' AND (nh_country = 'LT' OR nh_country = 'DK') ORDER BY nh_country, nh_sortorder, nh_name, nh_date"

                    'if session("mid") = 1 then
                    'rESPONSE.WRITE strSQL8
                    'Response.flush
                    'end if

                    objCmd = New OdbcCommand(strSQLhellig, objConn_tia)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                    While objDR2.Read() = True


                        erHellig = objDR2("nh_id")


                    End While
                    objDR2.Close()

                    objConn_tia.Close()

                    '** ER HELLIG ELLER en Lordag eller s�ndag



                End If


                '**** SLUT hellighdag



                'PRODUKTION

                'Der skal v�re EN rapport type 5 - KLAR TIL TEST

                firstDayInmonth = 0 '0
                erHellig = 0
                'weekDay = 4 pga Jul og nyt�r

                thisWeekDay = 1 '1 = mandag
                thisWeekDayTest = 2
                thisWeekDayPlan_1 = 1
                thisWeekDayPlan_2 = 2

                '** LM

                'Daily
                'PKT. 220 Opfyldning
                If CInt(weekDay) = thisWeekDayTest And CInt(firstDayInmonth) <> 1 And CInt(erHellig) = 0 And CInt(weekDayTime) = 10 And objDR("rapporttype") = 200 And (objDR("lto") = "outz" Or objDR("lto") = "wwf") Then '0
                    'Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 200)
                End If


                '** WWF

                'WEEKLY
                'PKT. 200 LM dem her der mere end 15 timers flex 
                If CInt(weekDay) = thisWeekDay And CInt(firstDayInmonth) <> 1 And CInt(erHellig) = 0 And CInt(weekDayTime) = 13 And objDR("rapporttype") = 200 And (objDR("lto") = "outz" Or objDR("lto") = "wwf") Then '0
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 200)
                End If

                'Daily
                'PKT. 210 Alle der ikke har registreret timer 
                If (CInt(weekDay) = 1 Or CInt(weekDay) = 2 Or CInt(weekDay) = 3 Or CInt(weekDay) = 4 Or CInt(weekDay) = 5) And CInt(firstDayInmonth) <> 1 And CInt(erHellig) = 0 And CInt(weekDayTime) = 15 And objDR("rapporttype") = 210 And (objDR("lto") = "outz" Or objDR("lto") = "wwf") Then '0
                    Call sendmailDailyNotifyNoHours(weekDay, objDR("lto"), 210)
                End If



                '** PLAN / Outz

                'WEEKLY



                'PKT. 2 FIRST reminder ALL monday 09.00 (�ndret fra s�ndag 23:59) 
                If CInt(weekDay) = thisWeekDayPlan_1 And CInt(firstDayInmonth) <> 1 And CInt(erHellig) = 0 And CInt(weekDayTime) = 9 And objDR("rapporttype") = 5 And (objDR("lto") = "xplan" Or objDR("lto") = "outz") Then '9
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 1)
                End If


                'PKT. 3 liste til LM Monday 13.00 med dem der ikke har afsluttet 
                If CInt(weekDay) = thisWeekDayPlan_1 And CInt(firstDayInmonth) <> 1 And CInt(erHellig) = 0 And CInt(weekDayTime) = 13 And objDR("rapporttype") = 9 And (objDR("lto") = "xplan" Or objDR("lto") = "outz") Then '13
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 9)
                End If

                'PKT. 7 REMINDER ALL LM who have not Approved WEEKLY & Project hours. Tuesday 10.00 
                If CInt(weekDay) = thisWeekDayPlan_2 And CInt(firstDayInmonth) <> 1 And CInt(erHellig) = 0 And CInt(weekDayTime) = 10 And objDR("rapporttype") = 10 And (objDR("lto") = "xplan" Or objDR("lto") = "outz") Then '10
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 10)
                End If




                '** TIA

                'WEEKLY
                'PKT. 3 liste til LM Sunday 23.59 med dem der ikke har afsluttet - Special costcenters - KLAR TIL TEST
                If CInt(weekDay) = thisWeekDay And CInt(firstDayInmonth) <> 1 And CInt(erHellig) = 0 And CInt(weekDayTime) = 0 And objDR("rapporttype") = 9 And objDR("lto") = "tia" Then '0
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 9)
                End If


                'PKT. 2 FIRST reminder ALL monday 07.00 (�ndret fra s�ndag 23:59) 
                If CInt(weekDay) = thisWeekDay And CInt(firstDayInmonth) <> 1 And CInt(erHellig) = 0 And CInt(weekDayTime) = 7 And objDR("rapporttype") = 5 And objDR("lto") = "tia" Then '7
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 1)
                End If



                'PKT. 4 SECOND reminder ALL in SnS (cost center 300, 320 & 330) Monday 09.00 - KLAR TIL TEST
                If CInt(weekDay) = thisWeekDay And CInt(firstDayInmonth) <> 1 And CInt(erHellig) = 0 And CInt(weekDayTime) = 9 And objDR("rapporttype") = 6 And objDR("lto") = "tia" Then ' 9
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 2)
                End If

                'PKT. 5 ALL Timesheet default completed 10:00 - KLAR TIL TEST -- FEJL ved helligdag. N�r den .feks k�rer tirsdag ska lden tr�kke en dage mere fra, da den ellers lukker mandag, dvs. indev�rende uge. == RETET 20190423
                If CInt(weekDay) = thisWeekDay And CInt(firstDayInmonth) <> 1 And CInt(erHellig) = 0 And CInt(weekDayTime) = 10 And objDR("rapporttype") = 8 And objDR("lto") = "tia" Then '1-10
                    Call defaultComplete(objDR("lto"), 8)
                End If



                'PKT. 7 REMINDER ALL PM + LM who have not Approved WEEKLY & Project hours. Monday 13.00 - Special costcenters
                If CInt(weekDay) = thisWeekDay And CInt(firstDayInmonth) <> 1 And CInt(erHellig) = 0 And CInt(weekDayTime) = 13 And objDR("rapporttype") = 10 And objDR("lto") = "tia" Then '13
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 10)
                End If

                'PKT. 7 PM REMINDER ALL PM who have Not Approved Project hours. Monday 13.00 - Special costcenters
                If CInt(weekDay) = thisWeekDay And CInt(firstDayInmonth) <> 1 And CInt(erHellig) = 0 And CInt(weekDayTime) = 13 And objDR("rapporttype") = 13 And objDR("lto") = "tia" Then '13
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 13)
                End If

                'PKT 8 Second reminder is sent to SnS mangers Deadline: Monday 3pm
                If CInt(weekDay) = thisWeekDay And CInt(firstDayInmonth) <> 1 And CInt(erHellig) = 0 And CInt(weekDayTime) = 15 And objDR("rapporttype") = 12 And objDR("lto") = "tia" Then '15
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 12)
                End If

                'PKT 8 Second reminder is sent to SnS mangers Deadline: Monday 3pm
                If CInt(weekDay) = thisWeekDay And CInt(firstDayInmonth) <> 1 And CInt(erHellig) = 0 And CInt(weekDayTime) = 15 And objDR("rapporttype") = 14 And objDR("lto") = "tia" Then '15
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 14)
                End If


                'PKT 9 Defalut godkend alle timer & uger - KLAR TIL TEST
                If CInt(weekDay) = thisWeekDay And CInt(firstDayInmonth) <> 1 And CInt(erHellig) = 0 And CInt(weekDayTime) = 16 And objDR("rapporttype") = 11 And objDR("lto") = "tia" Then '16
                    Call defaultApprove(objDR("lto"), 11)
                End If











                'MONTHLY
                'PKT. 3 liste til LM Sunday 23.59 med dem der ikke har afsluttet - Special costcenters - KLAR TIL TEST

                'weekDayTime = 0

                If CInt(erHellig) = 0 And CInt(firstDayInmonth) = 1 And CInt(weekDayTime) = 0 And objDR("rapporttype") = 9 And objDR("lto") = "tia" Then '0
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 109)
                End If

                'PKT. 2 FIRST reminder ALL monday 07.00 (�ndret fra s�ndag 23:59)
                '** Rapport skal v�re = 105 n�r Webservice kaldes. S� den tjekker splithr
                If CInt(erHellig) = 0 And CInt(firstDayInmonth) = 1 And CInt(weekDayTime) = 7 And objDR("rapporttype") = 105 And objDR("lto") = "tia" Then '7
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 105)
                End If

                'PKT. 4 SECOND reminder ALL in SnS (cost center 300, 320 & 330) Monday 09.00 - KLAR TIL TEST
                If CInt(erHellig) = 0 And CInt(firstDayInmonth) = 1 And CInt(weekDayTime) = 9 And objDR("rapporttype") = 106 And objDR("lto") = "tia" Then
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 106)
                End If


                'PKT. 5 ALL Timesheet default completed 10:00 - KLAR TIL TEST
                If CInt(erHellig) = 0 And CInt(firstDayInmonth) = 1 And CInt(weekDayTime) = 10 And objDR("rapporttype") = 8 And objDR("lto") = "tia" Then '10
                    Call defaultComplete(objDR("lto"), 108)
                End If



                'PKT. 6 LM costceter 600 joblog report
                If CInt(erHellig) = 0 And CInt(firstDayInmonth) = 1 And CInt(weekDayTime) = 10 And objDR("rapporttype") = 116 And objDR("lto") = "tia" Then
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 116)
                End If



                'PKT. 7 REMINDER ALL PM + LM who have not Approved WEEKLY & Project hours. Monday 13.00 - Special costcenters
                If CInt(erHellig) = 0 And CInt(firstDayInmonth) = 1 And CInt(weekDayTime) = 13 And objDR("rapporttype") = 10 And objDR("lto") = "tia" Then '13
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 110)
                End If

                'PKT. 7 PM REMINDER ALL PM who have Not Approved Project hours. Monday 13.00 - Special costcenters
                If CInt(erHellig) = 0 And CInt(firstDayInmonth) = 1 And CInt(weekDayTime) = 13 And objDR("rapporttype") = 13 And objDR("lto") = "tia" Then
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 113)
                End If

                'PKT 8 Second reminder is sent to SnS mangers Approved WEEKLY HR Deadline: Monday 3pm
                If CInt(erHellig) = 0 And CInt(firstDayInmonth) = 1 And CInt(weekDayTime) = 15 And objDR("rapporttype") = 12 And objDR("lto") = "tia" Then
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 112)
                End If

                'PKT 8 Second reminder is sent to SnS mangers Deadline: Monday 3pm
                If CInt(erHellig) = 0 And CInt(firstDayInmonth) = 1 And CInt(weekDayTime) = 15 And objDR("rapporttype") = 14 And objDR("lto") = "tia" Then
                    Call hentData(objDR("lto"), emailto, objDR("rapporttype"), 1, 114)
                End If


                'PKT 9 Defalut godkend alle timer & uger - KLAR TIL TEST
                If CInt(erHellig) = 0 And CInt(firstDayInmonth) = 1 And CInt(weekDayTime) = 16 And objDR("rapporttype") = 11 And objDR("lto") = "tia" Then '16
                    Call defaultApprove(objDR("lto"), 111)
                End If



                'Pkt 10 Sender godkendte til NAV
                'AFVENTER


                antalloops = antalloops + 1
            End While
            objDR.Close()




        End If




    End Sub


    Private Sub btnOverRide_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bt.Click

        meMTxt.Text = "Mail sendt"
        'Dim retval = SomeFunction(value)
        'lto = "outz"
        'emailto = "sk@outzource.dk"

        'Dim thisuser As String = Request("user")
        'lto = Request("lto")
        'fm_lto.Text = lto
        'fm_thisuser.Text = thisuser

        Call hentData(lto, emailto, 1, 0, 0)


        '' etc...
    End Sub


    Public mthNameThis = ""
    Sub mthName(mth, lto)


        Select Case mth
            Case 1
                mthNameThis = "January"
            Case 2
                mthNameThis = "February"
            Case 3
                mthNameThis = "March"
            Case 4
                mthNameThis = "April"
            Case 5
                mthNameThis = "May"
            Case 6
                mthNameThis = "June"
            Case 7
                mthNameThis = "Juli"
            Case 8
                mthNameThis = "August"
            Case 9
                mthNameThis = "September"
            Case 10
                mthNameThis = "October"
            Case 11
                mthNameThis = "November"
            Case 12
                mthNameThis = "December"
        End Select


        'Response.Write mthNameThis

    End Sub


    Public strSQLfilenames As String
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

        '*** FEJL 
        'Denne fejler ved omskift mellen TIA / Rapporter
        'Uge 53 i denne og ozreport_lto
        'WWF og Dencker modtagere i ADMIn rapport_abo
        emailto = "all"
        'emailto = Request("sendtospecificemail")

        'meMTxt.Text = "HER 1"

        'Response.Write("AMEAIL;:" & emailto)
        'Response.Flush()



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

        '** �bner Connection ***'
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



            strSQLfilenames = "SELECT afe_id, afe_email, afe_file, mnavn, init, afe_progrp FROM abonner_file_email"
            Select Case rapporttype
                Case 0, 1, 2, 3, 4, 21
                    strSQLfilenames += " LEFT JOIN medarbejdere ON (email = afe_email AND mansat = 1)"
                Case 5, 6, 9, 10, 12, 13, 14, 105, 106, 116, 200
                    strSQLfilenames += " LEFT JOIN medarbejdere ON (init = afe_init AND mansat = 1)"
                Case Else
                    strSQLfilenames += " LEFT JOIN medarbejdere ON (email = afe_email AND mansat = 1)"

            End Select


            'GROUP BY init fjernet p� rapporter 0,1,2,3,4 
            'Da en medarbejder godt kan abonnere p� flere rapporter
            '20191120
            Select Case rapporttype
                Case 0, 1, 2, 3, 4

                    If lto = "wwf" And rapporttype = 200 Then
                        strSQLfilenames += " WHERE afe_sent = 0 AND mansat = 1 And m.medarbejdertype <> 20 AND afe_email <> '' ORDER BY afe_id" 'GROUP BY init 
                    Else
                        strSQLfilenames += " WHERE afe_sent = 0 AND mansat = 1 ORDER BY afe_id" 'GROUP BY init 
                    End If

                Case 5, 6, 105, 106, 116, 200

                    strSQLfilenames += " WHERE afe_sent = 0 AND mansat = 1 GROUP BY init ORDER BY afe_id"

                Case 9, 10, 12, 13, 14
                    strSQLfilenames += " WHERE afe_sent = 0 AND mansat = 1 ORDER BY afe_id"
                Case Else

                    strSQLfilenames += " WHERE afe_sent = 0 AND mansat = 1 ORDER BY afe_id" 'GROUP BY init

            End Select

            'Response.Write(strSQLfilenames)



            Dim emailSendTo As String = ""
            objCmd = New OdbcCommand(strSQLfilenames, objConn)
            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)



            While objDR.Read() = True

                Select Case this_rapporttype
                    Case 5, 6, 105, 106 '+ 105 Monthly TIA WEEKLY 

                        thisuser = objDR("mnavn") & " [" & objDR("init") & "]"

                        Call sendmailWeeklyNotify(objDR("afe_email"), objDR("afe_file"), lto, ltofolder, thisuser, reminder_rule_no)

                    Case 9 ' + Monthly: 109 'TIA WEEKLY LM - NOT Closed Weeks

                        thisuser = objDR("mnavn") & " [" & objDR("init") & "]"

                        Call sendmailWeeklyNotifyLM(objDR("afe_email"), objDR("afe_file"), lto, ltofolder, thisuser, reminder_rule_no, objDR("afe_progrp"))

                    Case 10, 12, 13, 14 '+ Monthly 110 'TIA WEEKLY LM (10,12) +PM (13,14) NOT Approved hours 13:00 + 15:00 Sns 300

                        thisuser = objDR("mnavn") & " [" & objDR("init") & "]"

                        Call sendmailWeeklyNotifyLM(objDR("afe_email"), objDR("afe_file"), lto, ltofolder, thisuser, reminder_rule_no, objDR("afe_progrp"))

                    Case 116

                        thisuser = objDR("mnavn") & " [" & objDR("init") & "]"

                        Call sendmailWeeklyNotifyLM_joblog(objDR("afe_email"), objDR("afe_file"), lto, ltofolder, thisuser, reminder_rule_no, objDR("afe_progrp"))


                    Case 200 ' + Weekly/Monthly: If flex > 22

                        thisuser = objDR("mnavn") & " [" & objDR("init") & "]"

                        Call sendmailWeeklyNotifyTeamleader(objDR("afe_email"), objDR("afe_file"), lto, ltofolder, thisuser, reminder_rule_no, objDR("afe_progrp"))


                    Case Else 'Standard rapporter 0,1,2,3,4, 21

                        Call sendmail(objDR("afe_email"), objDR("afe_file"), lto, ltofolder, thisuser, this_rapporttype)
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
            meMTxt.Text = "Rapport ready - no mails sent"
        End If

        'meMTxt.Text += strSQLfilenames


    End Sub










    Function sendmail(email, file, lto, ltofolder, thisuser, this_rapporttype) As String


        Dim reporttype As String = this_rapporttype
        Dim Smtp_Server As New SmtpClient
        Dim e_mail As New MailMessage()
        'Smtp_Server.UseDefaultCredentials = True
        'Smtp_Server.Credentials = New Net.NetworkCredential("administrator", "Sok!2637")
        Smtp_Server.Port = 25 '587
        'Smtp_Server.EnableSsl = True
        Smtp_Server.Host = "formrelay.rackhosting.com"

        'Test
        'email = "sk@outzource.dk"

        e_mail = New MailMessage()
        e_mail.From = New MailAddress("timeout_no_reply@outzource.dk")
        e_mail.To.Add(email)

        Dim emailrapportname As String = ""
        'Select Case this_rapporttype
        ' Case "1"
        ' emailrapportname = "HR report - weekly"
        ' Case "21"
        ' emailrapportname = "HR report - monthly"
        ' Case "2"
        ' emailrapportname = "Projectleader report - by activity"
        ' Case "22"
        ' emailrapportname = "Projectleader report - by job"
        ' Case "3"
        ' emailrapportname = "Your Project report"
        ' Case "4"
        ' emailrapportname = "Resource report"
        ' End Select

        '** Lidt m�rkelig l�sning her, mend rapporttype bliver hardcoded = 1 p� sendmail funktionen.

        emailrapportname = "Weekly Report"

        If InStr(file, "_1_") > 0 Then
            emailrapportname = "HR report - weekly"
        End If

        If InStr(file, "_21_") > 0 Then
            emailrapportname = "HR report - monthly"
        End If

        If InStr(file, "_2_") > 0 Then
            emailrapportname = "Projectleader report - by activity"
        End If

        If InStr(file, "_22_") > 0 Then
            emailrapportname = "Projectleader report - by job"
        End If

        If InStr(file, "_3_") > 0 Then
            emailrapportname = "Your Project report"
        End If

        If InStr(file, "_4_") > 0 Then
            emailrapportname = "Resource report"
        End If

        e_mail.Subject = "TimeOut - Reportservice " + emailrapportname
        e_mail.IsBodyHtml = True
        e_mail.Body = "Hi TimeOut User.<br><br>"
        e_mail.Body += "This mail Is sent from the TimeOut email service. You can change the settings For your subscriptions In TimeOut under TSA - administration - Subscribe." '+ thisuser


        'Dim rdir As String = ""
        If lto = "dencker" Then
            'rdir = ""
            e_mail.Body += "<br><br>Login To TimeOut here: <br>https://timeout.cloud/timeout_xp/wwwroot/ver2_14/login.asp?lto=dencker&rdir=igv&key=2.084-0311-B033"
        Else
            'rdir = lto
            e_mail.Body += "<br><br>Login to TimeOut here: <br>https://timeout.cloud/" + lto '+ thisuser
        End If



        'End If

        Dim FileName As String = "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_14\inc\upload\" & ltofolder & "\" & file

        If System.IO.File.Exists(FileName) Then e_mail.Attachments.Add(New System.Net.Mail.Attachment(FileName))


        'e_mail.Attachments.Add()

        Smtp_Server.Send(e_mail)
        'MsgBox("Mail Sent")

        Return "Mails Sent"

    End Function




    Function sendmailDailyNotifyNoHours(weekday, lto, reminder_rule_no) As String

        Dim thisDateDatformat As String = ""
        Dim ddnow As Date = Date.Now
        Dim dd As Date = Date.Now
        Dim ddw As Integer = DatePart("w", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)


        If weekday = 1 Then
            dd = DateAdd("d", -3, dd)
        Else
            dd = DateAdd("d", -1, dd)
        End If

        '** Er ig�r Hellig? 

        Dim thisDateEndSQL As String = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)
        Dim thisWeek As String = DatePart("ww", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        'If thisWeek = 53 And DatePart("ww", DateAdd("d", 7, thisWeek), Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays) = 2 Then
        'thisWeek = 1
        'End If

        Dim thisMonth As String = DatePart("m", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        Dim lastMonth As String = DatePart("m", DateAdd("m", -1, ddnow), Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)

        Dim ddw_tjknorm As Integer = DatePart("w", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        Dim normtimerDagstr = "man"

        Select Case ddw_tjknorm
            Case 1
                normtimerDagstr = "man"
            Case 2
                normtimerDagstr = "tir"
            Case 3
                normtimerDagstr = "ons"
            Case 4
                normtimerDagstr = "tor"
            Case 5
                normtimerDagstr = "fre"
            Case 6
                normtimerDagstr = "lor"
            Case 7
                normtimerDagstr = "son"
        End Select

        Dim email As String




        Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" & lto & ";user=to_outzource2;Password=SKba200473;"

        '** �bner Connection ***'
        Dim objConn As OdbcConnection
        objConn = New OdbcConnection(strConn)
        objConn.Open()

        Dim objCmd As OdbcCommand
        Dim objDR As OdbcDataReader
        Dim objDR3 As OdbcDataReader
        Dim objDR4 As OdbcDataReader
        'Dim objDataSet As New DataSet

        Dim thisMid As String = 0
        Dim rdir As String = "frommaillink"
        Dim varTjDatoUS_man As String = ""



        '**** Henter alle aktive medarbjedere
        Dim strSQLprogrpids As String = ""
        Dim strSQLmtimer As String = ""
        Dim timer_real As Double = 0
        Dim timer_norm As Double = 0

        Dim antalEmails As Integer = 0
        Dim empMaxFlexStr As String = ""
        Dim strSQLtimer As String = ""
        Dim antalM As Integer = 0

        strSQLprogrpids = "SELECT mnavn, init, m.email, mid As medid, med_cal, ansatdato, medarbejdertype FROM medarbejdere m"
        strSQLprogrpids += " WHERE mansat = 1"
        strSQLprogrpids += " ORDER BY mid LIMIT 2000"

        objCmd = New OdbcCommand(strSQLprogrpids, objConn)
        objDR3 = objCmd.ExecuteReader
        While objDR3.Read()




            timer_norm = 0
            'Tager alle typer med. Tfaktim
            strSQLtimer = "Select normtimer_" & normtimerDagstr & " AS normtimer FROM medarbejdertyper WHERE id = " & objDR3("medarbejdertype") & ""
            objCmd = New OdbcCommand(strSQLtimer, objConn)
            objDR4 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            If objDR4.Read() Then

                timer_norm = objDR4("normtimer")

            End If
            objDR4.Close()

            timer_real = 0
            'Tager alle typer med. Tfaktim
            strSQLtimer = "Select SUM(timer) AS sumtimer FROM timer WHERE tmnr = " & objDR3("medid") & " AND tdato = '" & thisDateEndSQL & "' GROUP BY tmnr"
            objCmd = New OdbcCommand(strSQLtimer, objConn)
            objDR4 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            If objDR4.Read() Then

                timer_real = objDR4("sumtimer")

            End If
            objDR4.Close()


            If CDbl(timer_real) = 0 And CDbl(timer_norm) <> 0 Then


                Dim Smtp_Server As New SmtpClient
                Dim e_mail As New MailMessage()
                'Smtp_Server.UseDefaultCredentials = True
                'Smtp_Server.Credentials = New Net.NetworkCredential("administrator", "Sok!2637")
                Smtp_Server.Port = 25 '587
                'Smtp_Server.EnableSsl = True
                Smtp_Server.Host = "formrelay.rackhosting.com"

                email = objDR3("email")

                'TEST
                If email = "" Or Len(Trim(email)) = 0 Or InStr(email, "@") = 0 Then
                    email = "support@outzource.dk"
                End If

                'email = "sk@outzource.dk"

                e_mail = New MailMessage()
                e_mail.From = New MailAddress("timeout_no_reply@outzource.dk")
                e_mail.To.Add(email)



                e_mail.Subject = "TimeOut - Du har glemt at taste timer"
                e_mail.IsBodyHtml = True
                e_mail.Body = "K�re " & objDR3("mnavn") & "<br><br>"

                e_mail.Body += "Du mangler at taste timer d. "

                e_mail.Body += thisDateEndSQL & "<br><br>"

                e_mail.Body += "Med venlig hilsen TimeOut Reminder Service<br><br>"

                e_mail.Body += "Login to TimeOut: <a href=""https://timeout.cloud/" & lto & """>timeout.cloud/" & lto & "</a><br><br><br><br>"
                e_mail.Body += "This mail is sent from the TimeOut email service. " & ddnow & "" '+ thisuser

                Smtp_Server.Send(e_mail)

                antalEmails = antalEmails + 1
                antalM = antalM + 1

            End If


        End While
        objDR3.Close()



        Return "Mails Sent"

    End Function


    Function sendmailWeeklyNotifyTeamleader(email, file, lto, ltofolder, thisuser, reminder_rule_no, progrpid) As String

        Dim thisDateDatformat As String = ""
        Dim ddnow As Date = Date.Now
        Dim dd As Date = Date.Now
        Dim ddw As Integer = DatePart("w", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)



        dd = DateAdd("d", -1, dd)


        Dim thisDateEndSQL As String = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)
        Dim thisWeek As String = DatePart("ww", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        'If thisWeek = 53 And DatePart("ww", DateAdd("d", 7, thisWeek), Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays) = 2 Then
        'thisWeek = 1
        'End If
        Dim thisMonth As String = DatePart("m", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        Dim lastMonth As String = DatePart("m", DateAdd("m", -1, ddnow), Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)


        Dim Smtp_Server As New SmtpClient
        Dim e_mail As New MailMessage()
        'Smtp_Server.UseDefaultCredentials = True
        'Smtp_Server.Credentials = New Net.NetworkCredential("administrator", "Sok!2637")
        Smtp_Server.Port = 25 '587
        'Smtp_Server.EnableSsl = True
        Smtp_Server.Host = "formrelay.rackhosting.com"

        'TEST
        If email = "" Or Len(Trim(email)) = 0 Or InStr(email, "@") = 0 Then
            email = "support@outzource.dk"
        End If

        'email = "sk@outzource.dk"

        e_mail = New MailMessage()
        e_mail.From = New MailAddress("timeout_no_reply@outzource.dk")
        e_mail.To.Add(email)



        e_mail.Subject = "TimeOut - Max Flex Reminder"
        e_mail.IsBodyHtml = True
        e_mail.Body = "Hi " & thisuser & "<br><br>"

        e_mail.Body += "Max Flex has exceeded 15 hours for the following employees in your team:<br><br>"



        Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" & lto & ";user=to_outzource2;Password=SKba200473;"

        '** �bner Connection ***'
        Dim objConn As OdbcConnection
        objConn = New OdbcConnection(strConn)
        objConn.Open()

        Dim objCmd As OdbcCommand
        Dim objDR As OdbcDataReader
        Dim objDR3 As OdbcDataReader
        Dim objDR4 As OdbcDataReader
        'Dim objDataSet As New DataSet

        Dim thisMid As String = 0
        Dim rdir As String = "frommaillink"
        Dim varTjDatoUS_man As String = ""



        '**** Henter alle aktive medarbjedere
        Dim strSQLprogrpids As String = ""
        Dim strSQLmtimer As String = ""
        Dim mf_flex_norm_real As Double = 0
        Dim mf_flex_limit As Double = 15
        Dim antalEmails As Integer = 0
        Dim empMaxFlexStr As String = ""
        Dim strSQLflex As String = ""
        Dim antalM As Integer = 0

        strSQLprogrpids = "SELECT pr.id AS pid, m.mnavn, init, m.email, mid As medid, teamleder, med_cal, ansatdato FROM progrupperelationer pr "
        strSQLprogrpids += " Left Join medarbejdere m ON (m.mid = pr.medarbejderid And mansat = 1) "
        strSQLprogrpids += " WHERE pr.projektgruppeid = " & progrpid & " And teamleder = 0 And mansat = 1 "
        strSQLprogrpids += " ORDER BY pr.id LIMIT 2000"

        objCmd = New OdbcCommand(strSQLprogrpids, objConn)
        objDR3 = objCmd.ExecuteReader
        While objDR3.Read()



            strSQLflex = "Select mf_flex_norm_real FROM medarbejder_flexsaldo WHERE mf_mid = " & objDR3("medid")
            objCmd = New OdbcCommand(strSQLflex, objConn)
            objDR4 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            If objDR4.Read() Then

                mf_flex_norm_real = objDR4("mf_flex_norm_real")
                'empMaxFlexStr = 15

                If mf_flex_norm_real > mf_flex_limit Then

                    empMaxFlexStr = empMaxFlexStr & "<br>" & objDR3("mnavn") & ": "
                    empMaxFlexStr = empMaxFlexStr & mf_flex_norm_real  ' Employee: " & objDR2.Item("mnavn") & " [" & objDR2.Item("init") & "] hours: " & objDR4("timer") & "<br>"
                    antalEmails = antalEmails + 1
                    antalM = antalM + 1

                End If

            End If
            objDR4.Close()


        End While
        objDR3.Close()

        empMaxFlexStr = empMaxFlexStr & "<br><br>"

        e_mail.Body += empMaxFlexStr

        If antalM = 0 Then

            e_mail.Body += "<br><br>No employees on the list"

        End If


        e_mail.Body += "Login to TimeOut: <a href=""https://timeout.cloud/" & lto & """>timeout.cloud/" & lto & "</a><br><br><br><br>"
        'WEEK  CLOSE Button



        e_mail.Body += "This mail is sent from the TimeOut email service. " & ddnow & "" '+ thisuser

        'Dim FileName As String = "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_14\inc\upload\" & ltofolder & "\" & file

        'If System.IO.File.Exists(FileName) Then e_mail.Attachments.Add(New System.Net.Mail.Attachment(FileName))

        'If email = "sk@outzource.dk" Then
        'If thisuser = "TimeOut - Support [TSU]" Then
        Smtp_Server.Send(e_mail)
        'End If
        'End If
        'MsgBox("Mail Sent")

        Return "Mails Sent"

    End Function








    Function sendmailWeeklyNotify(email, file, lto, ltofolder, thisuser, reminder_rule_no) As String

        Dim thisDateDatformat As String = ""
        Dim ddnow As Date = Date.Now
        Dim dd As Date = Date.Now
        Dim ddw As Integer = DatePart("w", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)



        Dim ltodb As String = lto
        If lto = "intranet" Then
            ltodb = lto
            lto = "outz"
        Else
            ltodb = lto
            lto = lto
        End If

        If CInt(reminder_rule_no) = 105 Or CInt(reminder_rule_no) = 106 Then

            dd = DateAdd("d", -1, dd)

        Else

            Select Case lto
                Case "tia"

                    If ddw = 1 Then
                        dd = DateAdd("d", -1, dd)
                    Else
                        dd = DateAdd("d", -(ddw), dd)
                    End If

                Case "plan", "outz"

                    dd = dd 'DateAdd("d", -1, dd)

                Case Else

                    If ddw = 1 Then
                        dd = DateAdd("d", -1, dd)
                    Else
                        dd = DateAdd("d", -(ddw), dd)
                    End If

            End Select

        End If


        Dim thisDate As String = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)
        Dim thisWeek As String = DatePart("ww", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        'If thisWeek = 53 And DatePart("ww", DateAdd("d", 7, thisWeek), Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays) = 2 Then
        'thisWeek = 1
        'End If
        Dim thisMonth As String = DatePart("m", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        Dim lastMonth As String = DatePart("m", DateAdd("m", -1, ddnow), Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)

        Dim lastWeek As String = DatePart("ww", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        Dim writeWeek As String = ""

        Dim deadline As String = ""

        If CInt(reminder_rule_no) = 105 Or CInt(reminder_rule_no) = 106 Then
            deadline = "23:59"
        Else
            Select Case lto
                Case "tia"
                    deadline = "Sunday 23:59"
                Case "outz", "plan"
                    deadline = "Monday 10:00"
                Case Else
                    deadline = "Sunday 23:59"
            End Select

        End If




        If CInt(reminder_rule_no) = 1 Or CInt(reminder_rule_no) = 2 Then


            dd = dd 'DateAdd("d", -1, dd)
            thisDateDatformat = DatePart("d", dd) & "." & DatePart("m", dd) & "." & DatePart("yyyy", dd)
            thisDate = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)

            thisWeek = DatePart("ww", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)

            lastWeek = DateAdd("d", -7, dd)
            lastWeek = DatePart("ww", lastWeek, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)



            Select Case lto
                Case "tia"
                    writeWeek = thisWeek
                Case "plan", "outz"
                    writeWeek = lastWeek
                Case Else
                    writeWeek = thisWeek
            End Select



            'If thisWeek = 53 And DatePart("ww", DateAdd("d", 7, thisWeek), Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays) = 2 Then
            'thisWeek = 1
            'End If

            If CInt(reminder_rule_no) = 105 Or CInt(reminder_rule_no) = 106 Then
                deadline = "Last day in month at 23:59 (11:59 PM)"
            Else
                Select Case lto
                    Case "tia"
                        deadline = "Sunday " & thisDateDatformat & " at 23:59 (11:59 PM)"
                    Case "outz", "plan"
                        deadline = "Monday " & thisDateDatformat & " at 10:00 (10:00 AM)"
                    Case Else
                        deadline = "Sunday " & thisDateDatformat & " at 23:59 (11:59 PM)"
                End Select

            End If


        End If


        Dim Smtp_Server As New SmtpClient
        Dim e_mail As New MailMessage()
        'Smtp_Server.UseDefaultCredentials = True
        'Smtp_Server.Credentials = New Net.NetworkCredential("administrator", "Sok!2637")
        Smtp_Server.Port = 25 '587
        'Smtp_Server.EnableSsl = True
        Smtp_Server.Host = "formrelay.rackhosting.com"

        'TEST
        If email = "" Or Len(Trim(email)) = 0 Or InStr(email, "@") = 0 Then 'Or email = "raminta.ivoskute@nordgain.com"
            email = "timeout@outzource.dk"
        End If

        'email = "sk@outzource.dk"

        e_mail = New MailMessage()
        e_mail.From = New MailAddress("timeout_no_reply@outzource.dk")
        e_mail.To.Add(email)


        If CInt(reminder_rule_no) = 1 Then
            Select Case lto
                Case "tia"
                    e_mail.Subject = "TimeOut - Complete Week First Reminder"
                Case "outz", "plan"
                    e_mail.Subject = "TimeOut - Complete Week Reminder"
                Case Else
                    e_mail.Subject = "TimeOut - Complete Week First Reminder"
            End Select

        End If

        If CInt(reminder_rule_no) = 2 Then
            e_mail.Subject = "TimeOut - Complete Week Second Reminder"
        End If

        If CInt(reminder_rule_no) = 105 Then
            e_mail.Subject = "TimeOut - Complete Month First Reminder"
        End If

        If CInt(reminder_rule_no) = 106 Then
            e_mail.Subject = "TimeOut - Complete Month Second Reminder"
        End If

        e_mail.IsBodyHtml = True
        e_mail.Body = "Hi " & thisuser & "<br>"

        If CInt(reminder_rule_no) = 1 Or CInt(reminder_rule_no) = 2 Then
            e_mail.Body += "You have not completed your week no. " & writeWeek & " before deadline: " & deadline & ", please do it immediately.<br><br>"
        End If

        If CInt(reminder_rule_no) = 105 Or CInt(reminder_rule_no) = 106 Then
            Call mthName(lastMonth, lto)
            e_mail.Body += "You have not completed " & mthNameThis & " before deadline last day in month at: " & deadline & ", please do it immediately.<br><br>"
        End If

        If CInt(reminder_rule_no) = 2 Then
            e_mail.Body += "Notice: This Is the second reminder; all time registration will be autocompleted at 10.00 AM DK Time!<br><br>"
        End If



        Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" & ltodb & ";user=to_outzource2;Password=SKba200473;"

        '** �bner Connection ***'
        Dim objConn As OdbcConnection
        objConn = New OdbcConnection(strConn)
        objConn.Open()

        Dim objCmd As OdbcCommand
        Dim objDR As OdbcDataReader
        'Dim objDataSet As New DataSet

        Dim thisMid As String = 0
        Dim rdir As String = "frommaillink"
        Dim varTjDatoUS_man As String = ""

        If ddw = 1 Then
            varTjDatoUS_man = DateAdd("d", -7, dd)
        Else
            varTjDatoUS_man = DateAdd("d", -(7 + ddw), dd)
        End If

        varTjDatoUS_man = Year(varTjDatoUS_man) & "-" & Month(varTjDatoUS_man) & "-" & Day(varTjDatoUS_man)



        If CInt(reminder_rule_no) = 105 Or CInt(reminder_rule_no) = 106 Then


            '*** Complete Month function
            Dim strSQLm As String = "SELECT mid FROM medarbejdere WHERE email = '" & email & "' AND mansat = 1"
            objCmd = New OdbcCommand(strSQLm, objConn)
            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

            If objDR.Read() = True Then

                thisMid = objDR("mid")

            End If
            objDR.Close()



            e_mail.Body += "You can complete your month by logging into TimeOut: <a href=""https://timeout.cloud/" & lto & """>timeout.cloud/" & lto & "</a><br><br>"
            'e_mail.Body += "<a href = ""https://timeout.cloud/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?func=opdatersmiley&rdir=" & rdir & "&varTjDatoUS_man=" & varTjDatoUS_man & "&usemrn=" & thisMid & """>Complete Month now >></a><br><br>"
            'MONTH CLOSE Button
        Else
            e_mail.Body += "You can complete your week by logging into TimeOut: <a href=""https://timeout.cloud/" & lto & """>timeout.cloud/" & lto & "</a><br><br><br><br>"
            'WEEK  CLOSE Button
        End If



        e_mail.Body += "This mail is sent from the TimeOut email service. " & ddnow & "" '+ thisuser

        'Dim FileName As String = "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_14\inc\upload\" & ltofolder & "\" & file

        'If System.IO.File.Exists(FileName) Then e_mail.Attachments.Add(New System.Net.Mail.Attachment(FileName))

        'If email = "sk@outzource.dk" Then
        'If thisuser = "TimeOut - Support [TSU]" Then
        Smtp_Server.Send(e_mail)
        'End If
        'End If
        'MsgBox("Mail Sent")

        Return "Mails Sent"

    End Function



    Function sendmailWeeklyNotifyLM(email, file, lto, ltofolder, thisuser, reminder_rule_no, progrpId) As String

        Dim ltodb As String = lto
        If lto = "intranet" Then
            ltodb = lto
            lto = "outz"
        Else
            ltodb = lto
            lto = lto
        End If

        Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" & ltodb & ";user=to_outzource2;Password=SKba200473;"

        '** �bner Connection ***'
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



        '*** WEEKLY - MONTHLY
        Select Case CInt(reminder_rule_no)
            Case 109, 110, 113, 114, 112
                dd = DateAdd("m", -1, dd) 'V�r sikker p� at komme tilbage til forrige m�ned
            Case Else
                If ddw = 1 Then
                    dd = DateAdd("d", -1, dd)
                Else
                    dd = DateAdd("d", -(ddw), dd)
                End If
        End Select




        'End If


        Dim thisDate As String = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)
        Dim thisWeek As String = DatePart("ww", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        'If thisWeek = 53 And DatePart("ww", DateAdd("d", 7, thisWeek), Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays) = 2 Then
        'thisWeek = 1
        'End If

        Dim thisMonth As String = DatePart("m", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        Dim lastMonth As String = DatePart("m", DateAdd("m", -1, ddnow), Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)

        Dim deadline As String = "Sunday 23:59"
        Dim deadline_9_109 As String = "Sunday 23:59 / 11:59 PM"
        Dim deadline_10 As String = "1.00 PM"
        Dim deadline_12 As String = "3.00 PM"
        Select Case lto
            Case "tia"
                deadline = "Sunday 23:59"
                deadline_9_109 = "Sunday 23:59 / 11:59 PM"
                deadline_10 = "1.00 PM"
                deadline_12 = "3.00 PM"
            Case "outz", "plan"
                deadline = "Monday 10:00"
                deadline_9_109 = "Monday 10:00"
                deadline_10 = "13.00"
                deadline_12 = "13.00"
            Case Else
                deadline = "Sunday 23:59"
                deadline_9_109 = "Sunday 23:59 / 11:59 PM"
                deadline_10 = "1.00 PM"
                deadline_12 = "3.00 PM"
        End Select


        'Dim thisDate As String = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)
        'Dim thisWeek As String = DatePart("ww", dd)


        'Dim dgsondag As Date = Date.Now
        'dgsondag = dd 'DateAdd("d", -1, dd)
        'Dim slutDatoWeekSQL As String = DatePart("yyyy", dgsondag) & "-" & DatePart("m", dgsondag) & "-" & DatePart("d", dgsondag)

        Dim empNotCompleted As String = ""
        If CInt(reminder_rule_no) = 9 Then
            'empNotCompleted = "Employees not completed week no. " & thisWeek & ", it should be closed before: " & deadline & ".<br>"
            empNotCompleted = "The following employee/s from your team has Not completed their week no. " & thisWeek & " before deadline: " & deadline_9_109 & " DK Time.<br>"
        End If

        If CInt(reminder_rule_no) = 10 Then
            empNotCompleted = "You have not approved all weeks for the employees in your team before deadline: today " & deadline_10 & " DK Time.<br><br>" 'Employees hours/weeks not Approved:<br>"
            'empNotCompleted += "The following employee/s from your team still has non-approved weeks:<br><br>"
        End If

        If CInt(reminder_rule_no) = 12 Then
            empNotCompleted = "You have Not approved all weeks for the employees in your team before deadline: Today " & deadline_12 & " DK Time.<br><br>"
            'empNotCompleted += "The following employee/s from your team still has non-approved weeks:<br><br>"
        End If


        If CInt(reminder_rule_no) = 109 Then
            Call mthName(lastMonth, lto)
            'Dim mthNameThis As String = "June"
            'empNotCompleted = "Employees not completed week no. " & thisWeek & ", it should be closed before: " & deadline & ".<br>"
            empNotCompleted = "The following employee/s from your team has Not completed " & mthNameThis & " before deadline : last day in month at " & deadline_9_109 & " DK Time.<br>"
        End If

        If CInt(reminder_rule_no) = 110 Then
            empNotCompleted = "You have not approved last month for the employees in your team before deadline: today " & deadline_10 & " DK Time.<br><br>" 'Employees hours/weeks not Approved:<br>"
            'empNotCompleted += "The following employee/s from your team still has non-approved weeks:<br><br>"
        End If


        If CInt(reminder_rule_no) = 112 Then
            empNotCompleted = "You have Not approved last month for the employees in your team before deadline: Today " & deadline_12 & " DK Time.<br><br>"
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
        Dim ddSondayLastWeekSQL As String = ""

        Dim ddMondayLastWeekSQL As String = ""
        Dim ddMondayLastWeek As Date = Date.Now


        Dim ddSondayLastWeekMius7 As Date = Date.Now
        Dim ddSondayLastWeekMius7SQL As String = ""

        '*** WEEKLY - MONTHLY
        Select Case CInt(reminder_rule_no)
            Case 109, 110, 113, 114, 112

                Dim lastDay As String = "30"
                Select Case Month(dd)
                    Case 1, 3, 5, 7, 8, 10, 12
                        lastDay = "31"

                    Case 2
                        Select Case Year(dd)
                            Case "2000", "2004", "2008", "2012", "2016", "2020", "2024", "2028", "2032", "2036", "2040", "2044"
                                lastDay = "29"
                            Case Else
                                lastDay = "28"

                        End Select


                    Case Else

                        lastDay = "30"
                End Select

                ddSondayLastWeek = lastDay & "-" & Month(dd) & "-" & Year(dd) 'DateAdd("d", -(ddw), ddSondayLastWeek)
                ddSondayLastWeekSQL = DatePart("yyyy", ddSondayLastWeek) & "-" & DatePart("m", ddSondayLastWeek) & "-" & DatePart("d", ddSondayLastWeek)

                ddSondayLastWeekMius7 = DateAdd("d", -7, ddSondayLastWeek)
                ddSondayLastWeekMius7SQL = DatePart("yyyy", ddSondayLastWeekMius7) & "-" & DatePart("m", ddSondayLastWeekMius7) & "-" & DatePart("d", ddSondayLastWeekMius7)

            Case Else

                ddSondayLastWeek = dd 'DateAdd("d", -(ddw), ddSondayLastWeek)
                ddSondayLastWeekSQL = DatePart("yyyy", ddSondayLastWeek) & "-" & DatePart("m", ddSondayLastWeek) & "-" & DatePart("d", ddSondayLastWeek)

        End Select


        '*** WEEKLY - MONTHLY
        Select Case CInt(reminder_rule_no)
            Case 109, 110, 113, 114, 112

                ddMondayLastWeek = "1-" & Month(dd) & "-" & Year(dd)
                ddMondayLastWeekSQL = DatePart("yyyy", ddMondayLastWeek) & "-" & DatePart("m", ddMondayLastWeek) & "-" & DatePart("d", ddMondayLastWeek)

            Case Else

                Select Case lto
                    Case "tia"
                        ddMondayLastWeek = DateAdd("d", -(7), dd) 'DateAdd("d", -(ddw + 6), ddMondayLastWeek)
                    Case "plan", "outz"
                        ddMondayLastWeek = DateAdd("d", -(6), dd)
                    Case Else
                        ddMondayLastWeek = DateAdd("d", -(7), dd)
                End Select


                ddMondayLastWeekSQL = DatePart("yyyy", ddMondayLastWeek) & "-" & DatePart("m", ddMondayLastWeek) & "-" & DatePart("d", ddMondayLastWeek)

        End Select






        If CInt(reminder_rule_no) = 9 Or CInt(reminder_rule_no) = 10 Or CInt(reminder_rule_no) = 12 Or CInt(reminder_rule_no) = 13 Or CInt(reminder_rule_no) = 14 Then
            empNotCompleted += "Period: " & DatePart("yyyy", ddMondayLastWeek) & "-" & DatePart("m", ddMondayLastWeek) & "-" & DatePart("d", ddMondayLastWeek) & " And " & DatePart("yyyy", ddSondayLastWeek) & "-" & DatePart("m", ddSondayLastWeek) & "-" & DatePart("d", ddSondayLastWeek) & "<br><br>"
        End If


        If CInt(reminder_rule_no) = 109 Or CInt(reminder_rule_no) = 110 Or CInt(reminder_rule_no) = 113 Or CInt(reminder_rule_no) = 114 Or CInt(reminder_rule_no) = 112 Then
            'empNotCompleted += "Period: " & MonthName(DatePart("m", ddSondayLastWeek)) & "<br><br>"
            empNotCompleted += "<br><br>"
        End If


        Dim hoursAppStatusSQL As String = ""
        Dim weekAppStatusSQL As String = ""
        Dim strSQLugeafsluttet As String = ""

        If CInt(reminder_rule_no) = 10 Or CInt(reminder_rule_no) = 12 Or CInt(reminder_rule_no) = 13 Or CInt(reminder_rule_no) = 14 Then
            hoursAppStatusSQL = " And (godkendtstatus = 0 Or godkendtstatus = 3)" 'Ikke taget stilling / Tentative
            weekAppStatusSQL = " And ugegodkendt = 0 AND splithr = 0"
        End If

        If CInt(reminder_rule_no) = 109 Or CInt(reminder_rule_no) = 110 Or CInt(reminder_rule_no) = 113 Or CInt(reminder_rule_no) = 114 Or CInt(reminder_rule_no) = 112 Then
            'hoursAppStatusSQL = " And (godkendtstatus = 0 Or godkendtstatus = 3)" 'Ikke taget stilling / Tentative
            weekAppStatusSQL = " And splithr = 1"
        End If


        Select Case CInt(reminder_rule_no)
            Case 9, 10, 12, 109, 110, 112

                strSQLprogrpids = "SELECT pr.id AS pid, m.mnavn, init, m.email, mid As medid, teamleder, med_cal FROM progrupperelationer pr "
                strSQLprogrpids += " Left Join medarbejdere m ON (m.mid = pr.medarbejderid And mansat = 1) "
                strSQLprogrpids += " WHERE pr.projektgruppeid = " & progrpId & " And teamleder = 0 And mansat = 1 "
                strSQLprogrpids += " ORDER BY pr.id LIMIT 2000"

            Case 13, 14, 113, 114

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
            If ((CInt(reminder_rule_no) = 13 Or CInt(reminder_rule_no) = 14 Or CInt(reminder_rule_no) = 113 Or CInt(reminder_rule_no) = 114) And progrpId <> 16) Then
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



            'LM 9,10,12 - 109, 110
            If ((CInt(reminder_rule_no) = 9 Or CInt(reminder_rule_no) = 109) And objDR2("med_cal") <> "LT" And (progrpId = 11 Or progrpId = 15 Or progrpId = 17)) _
             Or ((CInt(reminder_rule_no) = 10 Or CInt(reminder_rule_no) = 110) And progrpId <> 16) _
             Or ((CInt(reminder_rule_no) = 12 Or CInt(reminder_rule_no) = 112) And (progrpId = 12 Or progrpId = 13 Or progrpId = 14)) Then
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



            '*** Henter medarbejdere der ikke har afsluttet WEEK i grupppen
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

            If CInt(reminder_rule_no) = 109 Or CInt(reminder_rule_no) = 110 Or CInt(reminder_rule_no) = 112 Then

                strSQLugeafsluttet = "Select mid FROM ugestatus WHERE (mid = " & objDR2("medid") & " And ( MONTH(uge) = MONTH('" & ddSondayLastWeekSQL & "') AND YEAR(uge) = YEAR('" & ddSondayLastWeekSQL & "') )  " & weekAppStatusSQL & ") OR (mid = " & objDR2("medid") & " And uge >= '" & ddSondayLastWeekMius7SQL & "' AND splithr = 0)"
                objCmd = New OdbcCommand(strSQLugeafsluttet, objConn)
                objDR4 = objCmd.ExecuteReader
                weekclosedbyuser = 0
                If objDR4.Read() = True Then
                    weekclosedbyuser = 1

                End If
                objDR4.Close()

            End If


            If CInt(reminder_rule_no) = 9 Or CInt(reminder_rule_no) = 109 Then

                'empNotCompleted += "tdato BETWEEN '" & ddMondayLastWeekSQL & "' AND '" & ddSondayLastWeekSQL & "<br><br>"

                If (CInt(weekclosedbyuser) = 0 And CInt(medidMember100400700) = 0) Or (CInt(weekclosedbyuser) = 0 And CInt(medidMember100400700) = 1 And medidHours100400700 > 0) Then

                    antalEmails = antalEmails + 1
                    empNotCompleted += objDR2.Item("mnavn") & " [" & objDR2.Item("init") & "] " & objDR2.Item("email") & "<br>"

                End If


            End If


            If CInt(reminder_rule_no) = 10 Or CInt(reminder_rule_no) = 110 Then 'LM not approved

                If (CInt(weekclosedbyuser) = 1 And medidMember100400700 = 1) Then

                    antalEmails = antalEmails + 1
                    empNotCompleted += objDR2.Item("mnavn") & " [" & objDR2.Item("init") & "] " & objDR2.Item("email") & " <br>" '" & strSQLugeafsluttet & "

                End If


            End If


            If CInt(reminder_rule_no) = 12 Or CInt(reminder_rule_no) = 112 Then 'LM not approved Second SnS

                If (CInt(weekclosedbyuser) = 1 And medidMember100400700 = 1) Then

                    antalEmails = antalEmails + 1
                    empNotCompleted += objDR2.Item("mnavn") & " [" & objDR2.Item("init") & "] " & objDR2.Item("email") & "<br>"

                End If


            End If



        End While
        objDR2.Close()


        Dim sendtest As Integer = 0
        If antalEmails > 0 Or sendtest = 1 Then

            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            'Smtp_Server.UseDefaultCredentials = True
            'Smtp_Server.Credentials = New Net.NetworkCredential("administrator", "Sok!2637")
            Smtp_Server.Port = 25 '587
            'Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "formrelay.rackhosting.com"


            If email = "" Or Len(Trim(email)) = 0 Or InStr(email, "@") = 0 Then
                email = "timeout@outzource.dk"
            End If

            'TEST
            'email = "sk@outzource.dk"

            e_mail = New MailMessage()
            e_mail.From = New MailAddress("timeout_no_reply@outzource.dk")
            e_mail.To.Add(email)


            '** Weekly
            If CInt(reminder_rule_no) = 9 Then
                Select Case lto
                    Case "tia"
                        e_mail.Subject = "TimeOut - LM - costcenter members complete Week reminder"
                    Case "outz", "plan"
                        e_mail.Subject = "TimeOut - Teamleder - Complete Week reminder"
                    Case Else
                        e_mail.Subject = "TimeOut - LM - costcenter members complete Week reminder"
                End Select

            End If

            If CInt(reminder_rule_no) = 10 Then
                Select Case lto
                    Case "tia"
                        e_mail.Subject = "TimeOut - LM - Weeks not Approved"
                    Case "outz", "plan"
                        e_mail.Subject = "TimeOut - Teamleder - Weeks not Approved"
                    Case Else
                        e_mail.Subject = "TimeOut - LM - Weeks not Approved"
                End Select

            End If

            If CInt(reminder_rule_no) = 12 Then
                e_mail.Subject = "TimeOut - LM - weeks not Approved second reminder"
            End If


            If CInt(reminder_rule_no) = 13 Or CInt(reminder_rule_no) = 113 Then
                e_mail.Subject = "TimeOut - PM - projecthours not Approved"
            End If

            If CInt(reminder_rule_no) = 14 Or CInt(reminder_rule_no) = 114 Then
                e_mail.Subject = "TimeOut - PM - projecthours not Approved - second reminder"
            End If

            '** Monthly
            If CInt(reminder_rule_no) = 109 Then
                e_mail.Subject = "TimeOut - LM - costcenter members complete Month reminder"
            End If

            If CInt(reminder_rule_no) = 110 Then
                e_mail.Subject = "TimeOut - LM - last month not Approved"
            End If

            If CInt(reminder_rule_no) = 112 Then
                e_mail.Subject = "TimeOut - LM - last month not Approved second reminder"
            End If

            e_mail.IsBodyHtml = True
            e_mail.Body = "Hi " & thisuser & "<br><br>"




            Select Case CInt(reminder_rule_no)
                Case 10, 12
                    e_mail.Body += "The following employee/s from your team still has non-approved weeks:<br><br>"
                Case 13, 14, 113, 114
                    e_mail.Body += "The following projects still has non-approved project hours:<br><br>"
                Case 109
                    'e_mail.Body += strSQLugeafsluttet
                Case 110, 112
                    e_mail.Body += "The following employee/s from your team still has non-approved month:<br><br>"
            End Select



            e_mail.Body += empNotCompleted


            Select Case CInt(reminder_rule_no)
                Case 999
                    e_mail.Body += "<br><br>Please make sure to complete all weeks immediately.<br><br>"
                Case 10, 12
                    e_mail.Body += "<br><br>Please make sure to approve all weeks immediately.<br><br>"
                Case 13, 14, 113, 114
                    e_mail.Body += "<br><br>Please make sure to approve all project hours immediately.<br><br>"
                Case 110, 112
                    e_mail.Body += "<br><br>Please make sure to approve month immediately.<br><br>"
            End Select

            If CInt(reminder_rule_no) = 10 Or CInt(reminder_rule_no) = 13 Or CInt(reminder_rule_no) = 110 Or CInt(reminder_rule_no) = 113 Then
                Select Case lto
                    Case "tia"
                        e_mail.Body += "<b>Notice : All time registration will be automatically approved by 4.00 PM DK Time!</b>"
                    Case "outz", "plan"
                        e_mail.Body += "<b>Notice : No time registration will be automatically approved.</b>"
                    Case Else
                        e_mail.Body += "<b>Notice : All time registration will be automatically approved by 4.00 PM DK Time!</b>"
                End Select

            End If


            If CInt(reminder_rule_no) = 12 Or CInt(reminder_rule_no) = 14 Or CInt(reminder_rule_no) = 114 Or CInt(reminder_rule_no) = 112 Then
                e_mail.Body += "<b>Notice: This is your second reminder and all-time registration will be automatically approved by 4.00 PM DK Time!</b><br><br>"
            End If

            Select Case CInt(reminder_rule_no)
                Case 9
                    e_mail.Body += "<br><br>You can complete weeks by logging into TimeOut: <a href=""https://timeout.cloud/" & lto & """>timeout.cloud/" & lto & "</a><br><br><br><br>"
                Case 10, 12
                    e_mail.Body += "<br><br>You can approve weeks by logging into TimeOut: <a href=""https://timeout.cloud/" & lto & """>timeout.cloud/" & lto & "</a><br><br><br><br>"
                Case 13, 14, 113, 114
                    e_mail.Body += "<br><br>You can approve hours by logging into TimeOut: <a href=""https://timeout.cloud/" & lto & """>timeout.cloud/" & lto & "</a><br><br><br><br>"
                Case 109, 110, 112
                    e_mail.Body += "<br><br>You can complete month by logging into TimeOut: <a href=""https://timeout.cloud/" & lto & """>timeout.cloud/" & lto & "</a><br><br><br><br>"

            End Select


            e_mail.Body += "This mail is sent from the TimeOut email service. " & ddnow & "" '+ thisuser

            'Dim FileName As String = "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_14\inc\upload\" & ltofolder & "\" & file

            'If System.IO.File.Exists(FileName) Then e_mail.Attachments.Add(New System.Net.Mail.Attachment(FileName))





            'If CInt(reminder_rule_no) <> 109 Then

            Smtp_Server.Send(e_mail)

            'Else

            'If thisuser = "Laima Stoniene [LST]" Then
            ' Smtp_Server.Send(e_mail)
            'End If

            'End If


            'MsgBox("Mail Sent")



        End If

        meMTxt.Text = "All Mails Sent"
        'Return "Mails Sent 2"

    End Function




    Function sendmailWeeklyNotifyLM_joblog(email, file, lto, ltofolder, thisuser, reminder_rule_no, progrpId) As String



        Dim ddnow As Date = Date.Now
        Dim dd As Date = Date.Now
        Dim ddw As Integer = DatePart("w", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)



        '*** WEEKLY - MONTHLY
        Select Case CInt(reminder_rule_no)
            Case 116
                dd = DateAdd("d", -1, dd)

        End Select



        Dim thisDate As String = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)
        Dim thisWeek As String = DatePart("ww", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        'If thisWeek = 53 And DatePart("ww", DateAdd("d", 7, thisWeek), Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays) = 2 Then
        'thisWeek = 1
        'End If
        Dim thisMonth As String = DatePart("m", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        Dim lastMonth As String = DatePart("m", DateAdd("m", -1, ddnow), Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)

        'Dim deadline As String = "Sunday 23:59"

        Dim empNotCompleted As String = ""
        empNotCompleted = "Joblog, hours stated on projects. Last month<br><br>"


        Dim antalEmails As Integer
        'Dim weekclosedbyuser As Integer = 0

        Dim strSQLprogrpids As String = ""
        'Dim strSQLjob As String = ""
        'Dim strSQL100400700 As String = ""
        'Dim medidMember100400700 As Integer = 0
        'Dim medidHours100400700 As Integer = 0

        Dim ddSondayLastWeek As Date = Date.Now
        ddSondayLastWeek = dd 'DateAdd("d", -(ddw), ddSondayLastWeek)
        'Dim ddSondayLastWeekSQL As String = DatePart("yyyy", ddSondayLastWeek) & "-" & DatePart("m", ddSondayLastWeek) & "-" & DatePart("d", ddSondayLastWeek)

        ' Dim ddMondayLastWeek As Date = Date.Now
        'ddMondayLastWeek = DateAdd("d", -(7), dd) 'DateAdd("d", -(ddw + 6), ddMondayLastWeek)
        'Dim ddMondayLastWeekSQL As String = DatePart("yyyy", ddMondayLastWeek) & "-" & DatePart("m", ddMondayLastWeek) & "-" & DatePart("d", ddMondayLastWeek)

        Call mthName(DatePart("m", ddSondayLastWeek), lto)
        'empNotCompleted += "Period: " & MonthName(DatePart("m", ddSondayLastWeek)) & "<br><br>"
        empNotCompleted += "Period: " & mthNameThis & "<br><br>"


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

        e_mail.Subject = "TimeOut - LM costcenter 600 - project joblog report"

        e_mail.IsBodyHtml = True
        e_mail.Body = "Hi " & thisuser & "<br><br>"
        e_mail.Body += "The following projects has project hours last month:<br><br>"

        e_mail.Body += empNotCompleted
        e_mail.Body += "<br><br>You can check hours by logging into TimeOut: <a href=""https://timeout.cloud/tia"">timeout.cloud/tia</a><br><br><br><br>"
        e_mail.Body += "This mail is sent from the TimeOut email service. " & ddnow & "" '+ thisuser


        Dim FileName As String = "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_14\inc\upload\" & ltofolder & "\" & file

        If System.IO.File.Exists(FileName) Then e_mail.Attachments.Add(New System.Net.Mail.Attachment(FileName))


        Smtp_Server.Send(e_mail)



        Return "Mails Sent"

    End Function




    Function defaultComplete(lto, reminder_rule_no) As String


        'lto = "tia_20190311"





        Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" & lto & ";user=to_outzource2;Password=SKba200473;"

        '** �bner Connection ***'
        Dim objConn As OdbcConnection
        objConn = New OdbcConnection(strConn)
        objConn.Open()

        Dim objCmd As OdbcCommand
        'Dim objDataSet As New DataSet

        Dim objDR2 As OdbcDataReader
        Dim objDR4 As OdbcDataReader

        Dim dd As Date = Date.Now
        Dim thisDateSQL As String = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd) & " " & DatePart("h", dd) & ":" & DatePart("n", dd) & ":02"

        Dim DayOfWeekNo As Integer = dd.DayOfWeek
        Dim thisUgeDag As Integer = 0

        Select Case DayOfWeekNo
            Case 0
                thisUgeDag = 7
            Case 1
                thisUgeDag = 1
            Case 2
                thisUgeDag = 2
            Case 3
                thisUgeDag = 3
            Case 4
                thisUgeDag = 4
            Case 5
                thisUgeDag = 5
            Case 6
                thisUgeDag = 6
        End Select


        dd = DateAdd("d", -thisUgeDag, dd)
        Dim thisWeek As String = DatePart("ww", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        'If thisWeek = 53 And DatePart("ww", DateAdd("d", 7, thisWeek), Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays) = 2 Then
        'thisWeek = 1
        'End If

        Dim lastWeekSundaySQL As String = "" 'DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)
        Dim lastWeekMondaySQL As String = ""


        If CInt(reminder_rule_no) = 108 Then 'Monthly
            Dim ddm As Date = Date.Now
            ddm = DateAdd("m", -1, ddm)
            'dd = DateAdd("d", 1, dd)

            'lastWeekMondaySQL = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-1" '& DatePart("d", dd)

            Dim lastDay As String = "30"
            Select Case Month(ddm)
                Case 1, 3, 5, 7, 8, 10, 12
                    lastDay = "31"

                Case 2
                    Select Case Year(ddm)
                        Case "2000", "2004", "2008", "2012", "2016", "2020", "2024", "2028", "2032", "2036", "2040", "2044"
                            lastDay = "29"
                        Case Else
                            lastDay = "28"

                    End Select


                Case Else

                    lastDay = "30"
            End Select

            Dim ddLastMonth As Date = lastDay & "-" & Month(ddm) & "-" & Year(ddm) 'DateAdd("d", -(ddw), ddSondayLastWeek)
            lastWeekSundaySQL = DatePart("yyyy", ddLastMonth) & "-" & DatePart("m", ddLastMonth) & "-" & lastDay
            lastWeekMondaySQL = DatePart("yyyy", ddLastMonth) & "-" & DatePart("m", ddLastMonth) & "-" & DatePart("d", ddLastMonth)




        Else
            lastWeekSundaySQL = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)
            dd = DateAdd("d", -6, dd)
            lastWeekMondaySQL = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)
        End If


        Dim deadline As String = "Monday 09:00"
        If CInt(reminder_rule_no) = 108 Then
            deadline = "First day in month at 09:00"
        Else
            deadline = deadline
        End If




        If CInt(reminder_rule_no) = 8 Then 'S�tter specille for rapport 8. Autocomplte alle medarbejderes uge + timer mandag 10.00. pkt 5.



        End If

        Dim weekclosedbyuser As Integer = 0

        Dim strSQLmedids As String = "SELECT mid AS medid, email, mnavn AS navn FROM medarbejdere WHERE mansat = 1 ORDER BY mid LIMIT 3000"
        objCmd = New OdbcCommand(strSQLmedids, objConn)
        objDR2 = objCmd.ExecuteReader
        While objDR2.Read()


            weekclosedbyuser = 0

            If CInt(reminder_rule_no) = 8 Then '** Weekly

                '** AND splithr = 1??? 
                '*** AND splithr = 0 tilf�je 20190911
                Dim strSQLugeafsluttet As String = "Select mid FROM ugestatus WHERE mid = " & objDR2.Item("medid") & " AND uge = '" & lastWeekSundaySQL & "' AND splithr = 0 GROUP BY mid"
                objCmd = New OdbcCommand(strSQLugeafsluttet, objConn)




                objDR4 = objCmd.ExecuteReader

                If objDR4.Read() = True Then

                    If IsDBNull(objDR4("mid")) <> True Then
                        weekclosedbyuser = 1
                    End If

                End If
                objDR4.Close()



                'If CInt(objDR2("medid")) = 1630 Then
                ' meMTxt.Text = "strSQLugeafsluttet: " & strSQLugeafsluttet & " weekclosedbyuser: " & weekclosedbyuser
                'End If

            Else '** Monthly



                Dim strSQLugeafsluttet As String = "Select mid FROM ugestatus WHERE mid = " & objDR2("medid") & " And ( MONTH(uge) = MONTH('" & lastWeekSundaySQL & "') AND YEAR(uge) = YEAR('" & lastWeekSundaySQL & "') ) AND splithr = 1"
                objCmd = New OdbcCommand(strSQLugeafsluttet, objConn)
                objDR4 = objCmd.ExecuteReader
                weekclosedbyuser = 0
                If objDR4.Read() = True Then

                    If IsDBNull(objDR4("mid")) <> True Then
                        weekclosedbyuser = 1
                    End If


                End If
                objDR4.Close()

            End If




            If CInt(weekclosedbyuser) = 0 Then

                Dim intStatusAfs As Integer = 2 '** Afsl. forsent




                If CInt(reminder_rule_no) = 8 Then
                    '**** SLETTER (Opdaterer for revisionsspor) EVT HRsplit midlertidig godkendt WEEK ***
                    '**** SPLIT HR bruger altid datoen fra den sidste uge i m�neden. 
                    '**** FULD uge godkendes igen fra WEEK approval

                    '**** Den sl�es fra 20190602, da LUK m�ned IKKE m� �bnes igen ****
                    'Dim strSQLafslutDelHRsplit As String = "UPDATE ugestatus SET uge = '2002-01-01' WHERE WEEK(uge, 3) = WEEK('" & lastWeekSundaySQL & "', 3) AND year(uge) = '" & Year(lastWeekSundaySQL) & "' AND mid = " & objDR2.Item("medid") & " AND splithr = 1"
                    'objCmd = New OdbcCommand(strSQLafslutDelHRsplit, objConn)
                    'objDR4 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    'objDR4.Close()

                End If

                Dim sglKriHrsplitflt = ""
                Dim sglKriHrsplitval = ""
                If CInt(reminder_rule_no) = 108 Then
                    sglKriHrsplitflt = ", splithr"
                    sglKriHrsplitval = ", 1"
                Else
                    sglKriHrsplitflt = ""
                    sglKriHrsplitval = ""
                End If

                Dim dgs As Date = Date.Now
                Dim strSQLupdateugestatus As String = "INSERT INTO ugestatus (uge, afsluttet, mid, status " & sglKriHrsplitflt & ") VALUES ('" & lastWeekSundaySQL & "', '" & thisDateSQL & "', " & objDR2.Item("medid") & ", " & intStatusAfs & "" & sglKriHrsplitval & ")"
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

        '** �bner Connection ***'
        Dim objConn As OdbcConnection
        objConn = New OdbcConnection(strConn)
        objConn.Open()

        Dim objCmd As OdbcCommand
        'Dim objDataSet As New DataSet

        Dim objDR2 As OdbcDataReader
        Dim objDR3 As OdbcDataReader
        Dim objDR4 As OdbcDataReader

        Dim dd As Date = Date.Now
        Dim ddm As Date = Date.Now
        Dim thisDateSQL As String = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)
        Dim thisDateSQLmm As String = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd) & " " & DatePart("h", dd) & ":" & DatePart("n", dd) & ":04"

        Dim DayOfWeekNo As Integer = dd.DayOfWeek
        Dim thisUgeDag As Integer = 0

        Select Case DayOfWeekNo
            Case 0
                thisUgeDag = 7
            Case 1
                thisUgeDag = 1
            Case 2
                thisUgeDag = 2
            Case 3
                thisUgeDag = 3
            Case 4
                thisUgeDag = 4
            Case 5
                thisUgeDag = 5
            Case 6
                thisUgeDag = 6
        End Select


        dd = DateAdd("d", -thisUgeDag, dd)
        'Dim thisWeek As String = DatePart("ww", dd)
        Dim thisWeek As String = DatePart("ww", dd, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        'If thisWeek = 53 And DatePart("ww", DateAdd("d", 7, thisWeek), Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays) = 2 Then
        'thisWeek = 1
        'End If

        Dim lastWeekSundaySQL As String = "" 'DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)
        Dim lastWeekMondaySQL As String = ""
        Dim ddSondayLastWeek As Date = Date.Now

        '*** WEEKLY - MONTHLY
        Select Case CInt(reminder_rule_no)
            Case 111

                ddm = DateAdd("m", -1, Date.Now)

                Dim lastDay As String = "30"
                Select Case Month(ddm)
                    Case 1, 3, 5, 7, 8, 10, 12
                        lastDay = "31"

                    Case 2
                        Select Case Year(ddm)
                            Case "2000", "2004", "2008", "2012", "2016", "2020", "2024", "2028", "2032", "2036", "2040", "2044"
                                lastDay = "29"
                            Case Else
                                lastDay = "28"

                        End Select


                    Case Else

                        lastDay = "30"
                End Select

                ddSondayLastWeek = lastDay & "-" & Month(ddm) & "-" & Year(ddm) 'DateAdd("d", -(ddw), ddSondayLastWeek)
                lastWeekSundaySQL = DatePart("yyyy", ddSondayLastWeek) & "-" & DatePart("m", ddSondayLastWeek) & "-" & DatePart("d", ddSondayLastWeek)


            Case Else

                lastWeekSundaySQL = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)

        End Select



        If CInt(reminder_rule_no) = 111 Then


            lastWeekMondaySQL = DatePart("yyyy", ddm) & "-" & DatePart("m", ddm) & "-1"
            Dim deadline As String = "Last day in month at 16:00"

        Else

            dd = DateAdd("d", -7, dd)
            lastWeekMondaySQL = DatePart("yyyy", dd) & "-" & DatePart("m", dd) & "-" & DatePart("d", dd)
            Dim deadline As String = "Monday 16:00"

        End If




        If CInt(reminder_rule_no) = 11 Then 'S�tter specille for rapport 11. 



        End If



        Dim weekclosedbyuser As Integer = 0

        Dim strSQLmedids As String = "SELECT mid AS medid, email, mnavn AS navn FROM medarbejdere WHERE mansat = 1 ORDER BY mid LIMIT 2000"
        objCmd = New OdbcCommand(strSQLmedids, objConn)
        objDR2 = objCmd.ExecuteReader
        While objDR2.Read()

            weekclosedbyuser = 0

            Dim strSQLugeafsluttet As String = ""
            If CInt(reminder_rule_no) = 111 Then
                strSQLugeafsluttet = "Select id, mid, ugegodkendt FROM ugestatus WHERE mid = " & objDR2.Item("medid") & " AND MONTH(uge) = MONTH('" & lastWeekSundaySQL & "') AND YEAR(uge) = YEAR('" & lastWeekSundaySQL & "') GROUP BY mid ORDER BY id DESC" 'GODKENDER ALT UANSET om DET ER HR
            Else
                strSQLugeafsluttet = "Select id, mid, ugegodkendt FROM ugestatus WHERE mid = " & objDR2.Item("medid") & " AND uge = '" & lastWeekSundaySQL & "' GROUP BY mid"
            End If


            'Response.Write(strSQLugeafsluttet & "<br><br>")
            'Response.Flush()



            objCmd = New OdbcCommand(strSQLugeafsluttet, objConn)
            objDR4 = objCmd.ExecuteReader

            If objDR4.Read() = True Then

                If CInt(reminder_rule_no) = 11 Then

                    '**** SLETTER (Opdaterer for revisionsspor) EVT HRsplit midlertidig godkendt WEEK ***
                    '**** SPLIT HR bruger altid datoen fra den sidste uge i m�neden. 
                    '**** FULD uge godkendes igen fra WEEK approval

                    '*** 20190911 Fjernet ligesom  ed complete. M�neden skal beholdes lukket.
                    'Dim strSQLafslutDelHRsplit As String = "UPDATE ugestatus SET uge = '2002-01-01' WHERE WEEK(uge, 3) = WEEK('" & lastWeekSundaySQL & "', 3) AND year(uge) = '" & Year(lastWeekSundaySQL) & "' AND mid = " & objDR4("mid") & " AND splithr = 1"
                    'objCmd = New OdbcCommand(strSQLafslutDelHRsplit, objConn)
                    'objDR3 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    'objDR3.Close()

                End If


                '***** Hvis aflsut m�ned / afslut uge er samme dag / weekend. Opdater Splithr = 0
                If CInt(reminder_rule_no) = 111 Then


                End If


                'If objDR4("id") <> "" Then
                If objDR4("ugegodkendt") = 0 Then

                    Dim dgs As Date = Date.Now
                    Dim strSQLupdateugestatus As String = "UPDATE ugestatus SET ugegodkendt = 1, ugegodkendtaf = 1, ugegodkendtdt = '" & thisDateSQLmm & "' WHERE id = " & objDR4("id")
                    objCmd = New OdbcCommand(strSQLupdateugestatus, objConn)
                    objDR3 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    objDR3.Close()

                    'End If

                    'End If
                End If

                weekclosedbyuser = 1

            End If
            objDR4.Close()



            '*** Hvis uge completed er slettet siden kl. 11.00
            If CInt(weekclosedbyuser) = 0 Then


                Dim sglKriHrsplitflt = ""
                Dim sglKriHrsplitval = ""
                If CInt(reminder_rule_no) = 111 Then
                    sglKriHrsplitflt = ", splithr"
                    sglKriHrsplitval = ", 1"
                Else
                    sglKriHrsplitflt = ""
                    sglKriHrsplitval = ""
                End If

                Dim intStatusAfs As Integer = 2 '** Afsl. forsent

                Dim dgs As Date = Date.Now
                Dim strSQLupdateugestatus As String = "INSERT INTO ugestatus (uge, afsluttet, mid, status " & sglKriHrsplitflt & ", ugegodkendt, ugegodkendtaf, ugegodkendtdt) VALUES ('" & lastWeekSundaySQL & "', '" & thisDateSQLmm & "', " & objDR2.Item("medid") & ", " & intStatusAfs & "" & sglKriHrsplitval & ", 1,1,'" & thisDateSQLmm & "')"
                objCmd = New OdbcCommand(strSQLupdateugestatus, objConn)
                objDR4 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                objDR4.Close()

            End If
            '****** END *****************************************

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

            '** �bner Connection ***'
            Dim objConn As OdbcConnection
            objConn = New OdbcConnection(strConn)
            objConn.Open()

            Dim objCmd As OdbcCommand
            Dim objDR As OdbcDataReader

            Dim strSQLfilenames As String = "SELECT email, navn, rapporttype FROM rapport_abo WHERE lto = '" & lto & "' GROUP BY email ORDER BY navn, email"
            objCmd = New OdbcCommand(strSQLfilenames, objConn)
            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

            '*** Sender alle de rapporter en medarbejder abonerer p� hvis man v�lger en medarb.


            While objDR.Read() = True



                strOptions += "<option value=" & objDR("email") & ">" & objDR("navn") & " - " & objDR("email") & "</option>"



            End While
            objDR.Close()
            ' Rap: " & objDR("rapporttype") & "
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

    <asp:Label runat="server" ID="meMTxt" Style="position:relative; width:300px; height:400px; vertical-align:top; border:1px #999999 solid;">Sys. besked:..</asp:Label>


<br /><br /><a href="#" onclick="Javascript:window.close();" style="color:red; font-size:12px; float:right;">[Close window X]</a>

        
    </div>
   
    
    
    </form>
</body>
</html>






        




