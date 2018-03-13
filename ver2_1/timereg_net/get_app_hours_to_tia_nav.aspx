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
<%@  Import Namespace="System.Web.Mail" %>


<!--@ Page Language="C#" AutoEventWireup="true" CodeFile="importer_job_wilke.cs" Inherits="importer_job_wilke" 

    <!--@  Import Namespace="System.Globalization" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<script runat="server" language="vbscript">


    Sub Page_load()



        Dim jobno = Request("jobno")
        Dim jobtaskno = Request("jobno")
        Dim action = Request("action")


        'meMTxt.Text = "HENT DATA"
        If (action = "check") Then ''Tjekker om der er timer på et job/task
            Call tjekTimer(jobno, jobtaskno)
        Else 'Henter timer til NAV
            Call hentData()
        End If


        'Response.Write("<br>No (INIT): " + Request("no"))





    End Sub


    Sub tjekTimer(ByRef jobno, ByRef jobtaskno)

        'Dim kategori As Integer
        Dim lto As String
        Dim t As Integer = 0

        lto = "tia"
        'meMTxt.Text = "HEJ SØREN"

        'Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_intranet;User=root;Password=;"
        Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" & lto & ";User=to_outzource2;Password=SKba200473;"

        Dim objConn As OdbcConnection
        objConn = New OdbcConnection(strConn)
        objConn.Open()

        Dim objCmd As OdbcCommand
        Dim objDR As OdbcDataReader

        If (jobtaskno <> "") Then

            Dim avarenr = "" + jobno + "" + jobtaskno + ""

            Dim strSQLmedins As String = "SELECT a.id "
            strSQLmedins += "FROM aktiviteter as a LEFT JOIN timer AS t ON (taktivitetid = a.id)"
            strSQLmedins += "WHERE a.avarenr = '" + avarenr + "' AND fakturerbar = 1"

            objCmd = New OdbcCommand(strSQLmedins, objConn)
            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

            If objDR.Read() = True Then

                t = 1
            End If
            objDR.Close()
            objConn.Close()

        Else



            Dim strSQLmedins As String = "SELECT tid "
            strSQLmedins += "FROM timer AS t "
            strSQLmedins += "WHERE tjobnr = '" + jobno + "'"

            objCmd = New OdbcCommand(strSQLmedins, objConn)
            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

            If objDR.Read() = True Then

                t = 1
            End If
            objDR.Close()
            objConn.Close()




        End If

        If (t > 0) Then 'Job / TASK spærret for redigering

            Response.Write("2")

        Else

            Response.Write("1")

        End If

    End Sub




    Sub hentData()




        'Dim kategori As Integer
        Dim lto As String

        lto = "tia"
        'meMTxt.Text = "HEJ SØREN"

        'Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_intranet;User=root;Password=;"
        Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" & lto & ";User=to_outzource2;Password=SKba200473;"

        Dim objConn As OdbcConnection
        objConn = New OdbcConnection(strConn)
        objConn.Open()

        Dim objCmd As OdbcCommand
        Dim objCmd2 As OdbcCommand
        'Dim objDataSet As New DataSet
        Dim objDR As OdbcDataReader
        Dim objDR2 As OdbcDataReader
        'Dim objDR3 As OdbcDataReader
        'Dim dt As DataTable
        'Dim dr As DataRow

        Dim errorFound As Integer = 0
        Dim errorFoundTxt As String = ""


        '*** MAIL *****
        Dim mailmodtagerOK As Integer = 0
        Dim strSQLlastJobIDMail As String = ""



        '***************************************************************************************
        '**** VIA recordset
        '***************************************************************************************

        'Dim row As DataRow
        Dim t As Integer = 0
        Dim avarenr As String = ""
        Dim init As String = ""
        Dim timerKom As String = ""
        Dim tjobnr As String = ""
        Dim Timerthis As String = "0"
        'Decimal = 0
        Dim tdato As String = ""

        Dim costcenter As String = ""


        'Dim Table1 As DataTable
        'Table1 = New DataTable("tb_to_var")

        'Dim column1 As DataColumn = New DataColumn("lto")
        'column1.DataType = System.Type.GetType("System.String")

        'Dim column2 As DataColumn = New DataColumn("importtype")
        'column2.DataType = System.Type.GetType("System.String")

        'Dim column3 As DataColumn = New DataColumn("init")
        'column3.DataType = System.Type.GetType("System.String")

        'Dim column4 As DataColumn = New DataColumn("mnavn")
        'column4.DataType = System.Type.GetType("System.String")

        'Dim column5 As DataColumn = New DataColumn("email")
        'column5.DataType = System.Type.GetType("System.String")

        '** TEST ENViRoNMENT
        'Dim CallWebService As New WebReferenceNAVTia_sendhours.TimeOut
        'CallWebService.Credentials = New System.Net.NetworkCredential(”tiademo”, ”Monday2017”, ”DEVX01”)

        Dim stDate As Date
        Dim slDate As Date
        '** Henter startdate
        Dim strSQLstDato As String = "Select lk_dato FROM lon_korsel WHERE lk_id > 0 ORDER BY lk_id DESC limit 1"
        objCmd = New OdbcCommand(strSQLstDato, objConn)
        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

        If objDR.Read() = True Then

            stDate = objDR("lk_dato")

        End If
        objDR.Close()

        stDate = DateAdd("d", 1, stDate)
        slDate = DateAdd("m", 1, stDate)
        slDate = DateAdd("d", -1, slDate)

        '** PROD ENViRoNMENT
        Dim CallWebService As New WebReferenceNAVTiaProd_sendhours.TimeOut
        CallWebService.Credentials = New System.Net.NetworkCredential(”Timeout”, ”Tia2017!”, ”tia.local”)

        Dim dDato As Date = Date.Now 'tirsdag morgen 07:00
        Dim datSQLformat As String = "yyyy-M-d"
        Dim overfortDatoSQL As String = dDato.ToString(datSQLformat)


        Dim stDatoSQLformat As String = "yyyy-M-d"
        stDatoSQLformat = stDate.ToString(stDatoSQLformat)

        Dim slDatoSQLformat As String = "yyyy-M-d"
        slDatoSQLformat = slDate.ToString(slDatoSQLformat)

        ''AND jobnr BETWEEN 7000 AND 72000

        Dim strSQLmedins As String = "SELECT tid, timer, m.init, tdato, tjobnr, a.avarenr, timerkom, tmnr "
        strSQLmedins += "FROM timer AS t "
        strSQLmedins += "LEFT JOIN medarbejdere AS m ON (m.mid = t.tmnr) "
        strSQLmedins += "LEFT JOIN aktiviteter AS a ON (a.id = t.taktivitetid) "
        strSQLmedins += "WHERE tdato BETWEEN '" & stDatoSQLformat & "' AND '" & slDatoSQLformat & "' AND (overfort = 0 OR overfort = 2) AND (godkendtstatus = 1) AND tmnr <> 1 AND timer <> 0 AND tfaktim <> '90' ORDER BY tid" ' AND avarenr <> '' AND avarenr IS NOT NULL ORDER BY tid "
        'AND tjobnr = 'PS1157'
        'timer_dobbel_20171103
        'strSQLmedins += "LIMIT 10"
        'AND init <> 'XXMAK'
        '_dobbel_20171004_3
        objCmd = New OdbcCommand(strSQLmedins, objConn)
        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

        While objDR.Read() = True





            If String.IsNullOrEmpty(objDR("avarenr")) <> True Then
                avarenr = objDR("avarenr")
            Else
                avarenr = "0"
            End If

            If String.IsNullOrEmpty(objDR("init")) <> True Then
                init = objDR("init")
            Else
                init = ""
            End If

            'If String.IsNullOrEmpty(objDR("timerkom")) <> True Then
            'timerKom = objDR("timerkom")
            'Else
            timerKom = ""
            'End If

            If String.IsNullOrEmpty(objDR("tjobnr")) <> True Then
                tjobnr = objDR("tjobnr")
            Else
                tjobnr = "0"
            End If

            If String.IsNullOrEmpty(objDR("tdato")) <> True Then
                tdato = objDR("tdato")
            Else
                tdato = "2002-01-01"
            End If

            If String.IsNullOrEmpty(objDR("timer")) <> True Then
                Timerthis = objDR("timer") '.ToString.Replace(",", ".")

            Else
                Timerthis = 0
            End If


            avarenr = UCase(Replace(avarenr, tjobnr, ""))


            'Set up structured error handling.
            Try
                ' Cause a "Divide by Zero" exception.


                If CallWebService.SendJournalData(tdato, tjobnr, avarenr, init, Timerthis, timerKom, objDR("tid")) = False Then
                    Dim strSQLtimerOverfort As String = "UPDATE timer SET overfort = 2, overfortdt = '" & overfortDatoSQL & "' WHERE tid = " & objDR("tid")
                    objCmd2 = New OdbcCommand(strSQLtimerOverfort, objConn)
                    objDR2 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)

                    errorFound = 1
                    errorFoundTxt += "<br>" & objDR("tjobnr") & "," & avarenr & ", " & init & ", " & objDR("tdato") & ", " & Timerthis & ""
                Else

                    'CallWebService.SendJournalData(tdato, tjobnr, avarenr, init, Timerthis, timerKom, objDR("tid"))
                    Dim strSQLtimerOverfort As String = "UPDATE timer SET overfort = 1, overfortdt = '" & overfortDatoSQL & "' WHERE tid = " & objDR("tid")
                    objCmd2 = New OdbcCommand(strSQLtimerOverfort, objConn)
                    objDR2 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)
                End If

                ' This statement does not execute because program
                ' control passes to the Catch block when the
                ' exception occurs.

            Catch ex As Exception

                ''Show the exception's message.
                'CallWebService.SendJournalError("+ ex.Message + " jobnr i TO:  " & objDR("tjobnr") & " ID:"& objDR("tid"))
                Response.Write("2 " + ex.Message + " jobnr i TO: " & objDR("tjobnr") & "<br>")


                Dim strSQLtimerOverfort As String = "UPDATE timer SET overfort = 2, overfortdt = '" & overfortDatoSQL & "' WHERE tid = " & objDR("tid")
                objCmd2 = New OdbcCommand(strSQLtimerOverfort, objConn)
                objDR2 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)

                errorFound = 1
                errorFoundTxt += "<br>" & objDR("tjobnr") & "," & avarenr & ", " & init & ", " & objDR("tdato") & ", " & Timerthis & ""

                'jobnr: " + err_jobnr + " errid:" + errThisTOnoStr



                ' Show the stack trace, which is a list of methods
                ' that are currently executing.
                'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)
                'Finally
                ' This line executes whether or not the exception occurs.
                'MessageBox.Show("in Finally block")
                'End Try



                'Catch ex As Exception

                'Response.Write("FEJL" & ex.ToString())



            End Try


            'objDR("avarenr")

            'row = Table1.NewRow()
            'row("lto") = lto
            'row("importtype") = "1"
            'row("init") = objDR("init")
            'row("tdato") = objDR("tdato")
            'row("timer") = objDR("timer")
            'row("timerkom") = objDR("timerkom")

            'row("expvendorno") = objDR("expvendorno")
            'row("costcenter") = objDR("costcenter")
            'row("linemanager") = objDR("linemanager")
            'row("countrycode") = objDR("countrycode")
            'row("weblang") = objDR("weblang")

            'taktivitetnavnLst = taktivitetnavnLst & objDR("init") & "<br>"

            'Add row
            'Table1.Rows.Add(row)



            t = t + 1




        End While


        objDR.Close()
        objConn.Close()

        If CInt(t) > 0 Then

            'datasrc.Text = taktivitetnavnLst

            ' Dim Table1 As DataTable
            'Table1 = New DataTable("tb_to_var")



            'Dim column1 As DataColumn = New DataColumn("lto")
            'column1.DataType = System.Type.GetType("System.String")

            'Table1.Columns.Add(column1)

            'Dim row As DataRow

            'row = Table1.NewRow()
            'row("lto") = "tia"

            ''Add row
            'Table1.Rows.Add(row)

            'Dim ds As DataSet = New DataSet("ds")
            'ds.Tables.Add(Table1)

            'Dim CallWebService As New dk_rack.outzource_timeout2_importmed_nav.oz_importmed_na()
            'CallWebService.addmed(ds)



            Response.Write("1")

        Else

            Response.Write("2")
            datasrc.Text = "Der opstod en fejl"


        End If


        'Response.Write("<br>" & strSQLmedins)




        '*** TIA error mail ***
        If CInt(errorFound) <> 0 Then



            Dim strBody As String
            mailmodtagerOK = 0
            'Return strSQLlastJobIDMail


            Dim myMail As Object
            myMail = CreateObject("CDO.Message")


            myMail.From = "timeout_no_reply@outzource.dk" 'TimeOut Email Service 

            myMail.To = "SNI@tiatechnology.com; jok@tiatechnology.com"
            myMail.CC = "sk@outzource.dk"
            'myMail.To = "sk@outzource.dk"
            myMail.Subject = "Record errors TimeOut >> NAV"


            strBody = "<br>Job -- Taks No. -- Init -- Date -- Hours <br><br>" & errorFoundTxt

            strBody = strBody & "<br><br><br><br><br><br>Best regards<br><i>"
            strBody = strBody & "TimeOut > NAV import service </i><br><br>&nbsp;"


            'Mailer.BodyText = strBody
            myMail.HTMLBody = "<html><head></head><body>" & strBody & "</body>"

            myMail.Configuration.Fields.Item _
            ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            'Name or IP of remote SMTP server


            Dim smtpServer As String = "formrelay.rackhosting.com"

            myMail.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpServer

            'Server port
            myMail.Configuration.Fields.Item _
            ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
            myMail.Configuration.Fields.Update

            myMail.Send
            mailmodtagerOK = mailmodtagerOK + 1
            myMail = Nothing




        End If






    End Sub





</script>
   


<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <asp:TextBox runat="server" ID="meid"></asp:TextBox>
    <asp:TextBox runat="server" ID="meMTxt" Style="width:600px; height:400px; vertical-align:top;">NAV Data:</asp:TextBox>
    </div>
    <asp:Button runat="server" Text="Hent data" ID="bt" OnClick="hentData"  />

    <h4>Reading Data from the connection
    <asp:Label ID="datasrc" runat="server"></asp:Label> to the DataGrid</h4>

    <asp:DataGrid ID="dgNameList" runat="server" /><br />
    </form>
</body>
</html>






        




