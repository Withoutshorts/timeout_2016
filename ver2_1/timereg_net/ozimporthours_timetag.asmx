<%@ WebService language="VB" class="to_import_timetag" %>
Imports System
Imports System.Web.Services
Imports System.Data
Imports System.Data.Odbc

Imports System.Globalization
Imports System.Threading

Public Class to_import_timetag

    Public testValue As String
    Public testmode As Integer = 0

    Public timerkom As String
    Public intMedarbId As String
    Public intTempImpId, intJobNr As String
    Public dlbTimer As String
    Public cdDato As Date
    Public importFrom As Integer

    Public errThisTOno As Integer = 0

    public editorIn As String
    Public meNavn As String
    Public meNr As String
    Public meID As Integer
    Public mTypeSQL As String

    Public jobNavn As String
    Public jobId As Double
    Public jobFastPris As Integer
    Public jobSeraft As Double
    Public jobKid As Double
    Public jobValuta As Integer

    Public aktNavn As String
    'aktnavnUse
    Public aktId As Double = 0
    Public aktFakturerbar As Integer

    Public kurs As String

    Public kNavn As String
    Public kNr As Double

    Public extsysid As double

    Public intJobNrTxt As String

    Public intTimepris As String = 0
    Public timeprisAlt As Double = 0
    Public intValuta As Double
    Public tpid As Double
    Public foundone As String = "n"
    Public timeprisalernativ, valutaAlt As String
    Public mtp As Integer
    Public aktIdUse As Double

    Public kostpris As String
    Public tprisGen As String
    Public valutaGen As Integer = 1



    Public emx As String = 100
    Public tempM As String

    Public errThisTOnoAll As String = "ErrorMSG: "
    Public len_tempM As String
    Public akttypeUse As String

    Public dbnavn As String = ""
    Public cdDatoSQL As String
    Public ugeErLukket As Integer = 0

    Public intCountInserted As Integer = 0

    Public err_jobnr As String = String.Empty

    Public lto As String

    Public strSQLerrTest As String
    Public t As Integer


    Public stTid As String = "00:00:00"
    Public slTid As String = "00:00:00"


    'Public errThisTOnoAll As Integer = 0

    Public errThisTOnoStr As String = "<br>Fejlkoder for linjer der blev forsøgt indlæst:<br>"
    Public errTxtThis As string


    Function EncodeUTF8(ByVal s)
        Dim i
        Dim c

        i = 1
        Do While i <= Len(s)
            c = Asc(Mid(s, i, 1))
            If c >= &H80 Then
                s = Left(s, i - 1) + Chr(&HC2 + ((c And &H40) / &H40)) + Chr(c And &HBF) + Mid(s, i + 1)
                i = i + 1
            End If
            i = i + 1
        Loop

        's = Replace(s, "Ã,", "&oslash;")
        EncodeUTF8 = s
    End Function

    Function DecodeUTF8(ByVal s)
        Dim i
        Dim c
        Dim n

        i = 1
        Do While i <= Len(s)
            c = Asc(Mid(s, i, 1))
            If c And &H80 Then
                n = 1
                Do While i + n < Len(s)
                    If (Asc(Mid(s, i + n, 1)) And &HC0) <> &H80 Then
                        Exit Do
                    End If
                    n = n + 1
                Loop
                If n = 2 And ((c And &HE0) = &HC0) Then
                    c = Asc(Mid(s, i + 1, 1)) + &H40 * (c And &H1)
                Else
                    c = 191
                End If
                s = Left(s, i - 1) + Chr(c) + Mid(s, i + n)
            End If
            i = i + 1
        Loop
        DecodeUTF8 = s
    End Function

    Public jq_formatTxt As String
    Function jq_format(ByVal jq_str As String) As String

        jq_str = Replace(jq_str, "ø", "&oslash;")
        jq_str = Replace(jq_str, "æ", "&aelig;")
        jq_str = Replace(jq_str, "å", "&aring;")
        jq_str = Replace(jq_str, "Ø", "&Oslash;")
        jq_str = Replace(jq_str, "Æ", "&AElig;")
        jq_str = Replace(jq_str, "Å", "&Aring;")
        jq_str = Replace(jq_str, "Ö", "&Ouml;")
        jq_str = Replace(jq_str, "ö", "&ouml;")
        jq_str = Replace(jq_str, "Ü", "&Uuml;")
        jq_str = Replace(jq_str, "ü", "&uuml;")
        jq_str = Replace(jq_str, "Ä", "&Auml;")
        jq_str = Replace(jq_str, "ä", "&auml;")
        jq_str = Replace(jq_str, "é", "&eacute;")
        jq_str = Replace(jq_str, "É", "&Eacute;")
        jq_str = Replace(jq_str, "á", "&aacute;")
        jq_str = Replace(jq_str, "Á", "&Aacute;")
        jq_str = Replace(jq_str, "Ã,", "&Oslash;")


        jq_formatTxt = jq_str

    End Function


    <WebMethod()> Public Function timeout_importTimer_timetag(ByVal ds As DataSet) As String





        '**** Oringin 10: Excel upload (forberedt lidt, men kører i sin egen fil ozimporthours.asmx)       
        '**** Origin 11: TimeTag 



        'On Error Resume Next


        'Dim con As New SqlConnection(Application("MyDatabaseConnectionString"))
        'Dim daCust As New SqlDataAdapter("Select * From tblCustomer", con)
        'Dim cbCust As New SqlCommandBuilder(daCust)
        'daCust.Update(ds, "Cust")
        'Return ds        

        Dim objConn As OdbcConnection
        Dim objCmd As OdbcCommand
        'Dim objDataSet As New DataSet
        Dim objDR As OdbcDataReader
        Dim objDR2 As OdbcDataReader
        Dim dt As DataTable
        Dim dr As DataRow

        Dim strConn As String

        'Dim objTable As DataTable

        'Dim t As Double = 0
        t = 0
        'Try


        'DataColumn column1 = new DataColumn("id", typeof(double)); 0

        'DataColumn column2 = new DataColumn("dato", typeof(DateTime)); 1

        'DataColumn column3 = new DataColumn("editor", typeof(string)); 2

        'DataColumn column4 = new DataColumn("origin", typeof(int)); 3

        'DataColumn column5 = new DataColumn("medarbejderid", typeof(string)); 4

        'DataColumn column6 = new DataColumn("jobid", typeof(int)); 5

        'DataColumn column7 = new DataColumn("aktnavn", typeof(string)); 6

        'DataColumn column8 = new DataColumn("timer", typeof(double)); 7

        'DataColumn column9 = new DataColumn("tdato", typeof(DateTime)); 8

        'DataColumn column10 = new DataColumn("lto", typeof(string)); 9

        'DataColumn column11 = new DataColumn("timerkom", typeof(string)); 10

        'DataColumn column12 = new DataColumn("aktid", typeof(bigint)); 11

        'DataColumn column13 = new DataColumn("jobnavn", typeof(string)); 12

        'DataColumn column14 = new DataColumn("sttid", typeof(string)); 13
        'DataColumn column15 = new DataColumn("sltid", typeof(string)); 14








        For Each dt In ds.Tables
            For Each dr In dt.Rows



                'Return "her"




                'Initialize
                errThisTOno = 0

                meID = 0
                meNavn = "-"
                meNr = "0"
                extsysid = 0

                intJobNr = 0
                'intjobId = 0
                intValuta = 0
                intMedarbId = ""
                jobKid = 0
                jobId = 0
                intJobNrTxt = ""
                foundone = "n"

                kNr = 0
                kNavn = ""

                intTimepris = 0
                timeprisAlt = 0
                kurs = 100
                kostpris = 0
                mTypeSQL = ""


                importFrom = ds.Tables(0).Rows(t).Item(3)


                dbnavn = ""

                Try
                    'If ds.Tables("timer_import_temp").Columns.Contains("lto") Then

                    If String.IsNullOrEmpty(ds.Tables(0).Rows(t).Item(9)) = False Then

                        'LTO
                        dbnavn = ds.Tables(0).Rows(t).Item(9)
                        'dbnavn = "outz"
                    Else
                        dbnavn = "demo"
                    End If

                    If dbnavn = "outz" Or dbnavn = "intranet - local" Then
                        dbnavn = "intranet"
                    Else
                        dbnavn = dbnavn
                    End If

                    'End If
                Catch ex As Exception
                    Throw New Exception("Get dbnavn error: " + ex.Message)
                End Try








                '**** Conn string ****'

                If t = 0 Then

                    'Try

                    'importFrom = "3" Then 'MMMI / ekstern
                    'dbnavn = "intranet"
                    '81.19.249.35
                    'strConn = "Driver={MySQL ODBC 3.51 Driver};Server=62.182.173.226;Database=timeout_intranet;User=outzource;Password=SKba200473;" '" & dbnavn & "


                    Select Case LCase(dbnavn)
                        Case "wilke", "intranet"
                            strConn = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154; Port=3306; Database=timeout_" & dbnavn & ";User=to_outzource2;Password=SKba200473; Option=32;"
                        Case Else
                            strConn = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_" & dbnavn & ";"
                    End Select

                    'strConn = "Driver={MySQL ODBC 3.51 Driver};Server=172.0.0.1;Database=timeout_intranet;User=outzource;Password=SKba200473;"


                    '** Åbner Connection ***'
                    objConn = New OdbcConnection(strConn)
                    objConn.Open()

                    'Catch ex As Exception
                    'Throw New Exception("If t=0 database connect and open error: " + ex.Message)
                    'End Try


                End If


                'Return "Er i loopet:  linjer: " + ds.Tables(0).Rows.Count.ToString() + " importFrom: " + importFrom



                'If ds.Tables("timer_import_temp").Columns('0') Then '.Contains("origin")
                'importFrom = ds.Tables("timer_import_temp").Rows(t).Item("origin")

                '** 10 Excelupload
                '** 11 TimeTag







                Try
                    'cdDato = "2012-10-10"
                    'If ds.Tables("timer_import_temp").Columns.Contains("tdato") Then
                    'cdDato = ds.Tables("timer_import_temp").Rows(t).Item("tdato")

                    ' Sets the CurrentCulture property to Danish Denmark
                    Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")

                    If String.IsNullOrEmpty(ds.Tables(0).Rows(t).Item(8)) = False Then
                        cdDato = ds.Tables(0).Rows(t).Item(8)
                    Else
                        cdDato = New Date(2012, 1, 1)
                        errThisTOno = 8
                    End If


                    'cdDatoSQL = Year(cdDato) & "/" & Month(cdDato) & "/" & Day(cdDato)

                Catch ex As Exception
                    Throw New Exception("Get cdDato error: " + ex.Message)
                End Try


                cdDatoSQL = Year(cdDato) & "/" & Month(cdDato) & "/" & Day(cdDato)















                'errThisTOno = 6
                '*** end **'

                If importFrom = 10 Then

                    Try
                        '*** if record findes i timereg i forvejen ****'
                        intTempImpId = ds.Tables(0).Rows(t).Item(0)
                    Catch ex As Exception
                        Throw New Exception("Get intTempImpId error:" + ex.Message)
                    End Try

                    intTempImpId = 99

                End If


                'If importFrom = 10 Then
                'editorIn = "ExcelUpload"
                '    End If

                'If importFrom = 11 Then
                editorIn = "TimeTag"
                'End If






                Try
                    '**** Jobnavn ****
                    jobNavn = ".."

                    If String.IsNullOrEmpty(ds.Tables(0).Rows(t).Item(12)) = False Then
                        jobNavn = ds.Tables(0).Rows(t).Item(12)
                        jobNavn = jobNavn.ToString()
                    Else
                        jobNavn = "NaN"
                    End If


                Catch ex As Exception
                    Throw New Exception("Get jobNavn error:" + ex.Message)
                End Try


                jobNavn = Replace(jobNavn, "'", "''")
                jobNavn = EncodeUTF8(jobNavn)
                jobNavn = DecodeUTF8(jobNavn)





                Try
                    '**** Aktivitets navn ****
                    aktNavn = ""
                    'If ds.Tables("timer_import_temp").Columns.Contains("aktnavn") Then
                    If String.IsNullOrEmpty(ds.Tables(0).Rows(t).Item(6)) = False Then
                        aktNavn = ds.Tables(0).Rows(t).Item(6)
                        aktNavn = aktNavn.ToString()
                    End If
                Catch ex As Exception
                    Throw New Exception("Get aktNavn error:" + ex.Message)
                End Try

                aktNavn = Replace(aktNavn, "'", "''")
                aktNavn = EncodeUTF8(aktNavn)
                aktNavn = DecodeUTF8(aktNavn)


                Try
                    '** Kommentar **' 11??
                    timerkom = ""
                    If String.IsNullOrEmpty(ds.Tables(0).Rows(t).Item(10)) = False Then
                        timerkom = ds.Tables(0).Rows(t).Item(10)
                        timerkom = timerkom.ToString()
                    End If

                Catch ex As Exception
                    Throw New Exception("Get timerkom error:" + ex.Message)
                End Try


                Try
                    '*** Medarbejer ID ****'
                    tempM = IsDBNull(ds.Tables(0).Rows(t).Item(4))
                    'tempM = "SK"
                    If tempM <> True Then
                        tempM = ds.Tables(0).Rows(t).Item(4)
                        tempM = tempM.ToString()
                    Else
                        tempM = ""
                    End If

                    intMedarbId = tempM
                Catch ex As Exception
                    Throw New Exception("Get intMedarbId error:" + ex.Message)
                End Try



                Try
                    '** Jobnr = jobid **'
                    If String.IsNullOrEmpty(ds.Tables(0).Rows(t).Item(5)) = False Then
                        intJobNr = ds.Tables(0).Rows(t).Item(5)
                    Else
                        intJobNr = 0
                    End If

                    intJobNrTxt = intJobNr
                Catch ex As Exception
                    Throw New Exception("Get intJobNr error:" + ex.Message)
                End Try


                Try
                    '** LTO **'
                    lto = ""
                    If String.IsNullOrEmpty(ds.Tables(0).Rows(t).Item(9)) = False Then
                        lto = ds.Tables(0).Rows(t).Item(9)
                    Else
                        lto = 0
                    End If


                Catch ex As Exception
                    Throw New Exception("Get intJobNr error:" + ex.Message)
                End Try


                Try
                    '** Timer ***'
                    If String.IsNullOrEmpty(ds.Tables(0).Rows(t).Item(7)) = False Then
                        dlbTimer = ds.Tables(0).Rows(t).Item(7)


                        If InStr(dlbTimer, ",") <> 0 OR InStr(dlbTimer, ".") <> 0 Then



                            '*

                            'dlbTimer = Replace(dlbTimer, ".", ",")
                            'If InStr(dlbTimer, ",") = 1 Then '< 10 timer
                            'trmhrs = Left(dlbTimer, 1)
                            'trmmin = (Mid(dlbTimer, 2, 2) * 100) / 60
                            'Else
                            'trmhrs = Left(dlbTimer, 2)
                            'trmmin = (Mid(dlbTimer, 3, 2) * 100) / 60
                            'End If









                            'trmmin = Left(trmmin, 2)

                            'dlbTimer = trmhrs & trmmin
                            ''& "." & trmmin trmhrs 
                            '*********




                            'If lto = "outz" Then

                            'timerkom = timerkom &"/"& dlbTimer &"/"& Mid(dlbTimer, 2, 1) 
                            'Dim timerOrg As string = dlbTimer

                            dlbTimer = Replace(dlbTimer, ".", ",")

                            Dim trmmin As String ', trmminTjk As Single


                            Dim trmhrs As String
                            Dim D As String, TB As Array, tmin As Single, trmminArr As Array

                            'As Array 'Result As Single

                            D = dlbTimer
                            TB = Split(D, ",")
                            trmhrs = TB(0)
                            'tmin = TB(1)/60
                            trmmin = ((TB(1) * 100) / 60) '/ 100

                            trmminArr = Split(trmmin, ",")

                            'trmminTjk = trmminArr(0) / 1

                            'If trmminTjk < 10 Then 'STRANGE: Noget med 05 min bliver = 5 og dermed forveksles med 50 min.
                            If ((Mid(dlbTimer, 3, 1) = "0" And Mid(dlbTimer, 2, 1) = ",") And (Mid(dlbTimer, 4, 1) = "0" Or Mid(dlbTimer, 4, 1) = "1" Or Mid(dlbTimer, 4, 1) = "2" Or Mid(dlbTimer, 4, 1) = "3" Or Mid(dlbTimer, 4, 1) = "4" Or Mid(dlbTimer, 4, 1) = "5" Or Mid(dlbTimer, 4, 1) = "6")) _
                            Or ((Mid(dlbTimer, 4, 1) = "0" And Mid(dlbTimer, 3, 1) = ",") And (Mid(dlbTimer, 4, 1) = "0" Or Mid(dlbTimer, 5, 1) = "1" Or Mid(dlbTimer, 5, 1) = "2" Or Mid(dlbTimer, 5, 1) = "3" Or Mid(dlbTimer, 5, 1) = "4" Or Mid(dlbTimer, 5, 1) = "5" Or Mid(dlbTimer, 5, 1) = "6")) Then
                                '< 10 timer

                                '
                                '



                                If Left(Replace(trmmin, ".", ""), 3) <> "999" Then '6 min

                                    dlbTimer = trmhrs & ",0"
                                    trmmin = Left(trmminArr(0), 3)

                                Else

                                    dlbTimer = trmhrs & ","
                                    trmmin = Left(trmminArr(0), 3)

                                End If






                            Else
                                dlbTimer = trmhrs & ","
                                trmmin = Left(trmminArr(0), 3)
                            End If

                            trmmin = Replace(trmmin, ".", "")
                            'trmmin = FormatNumber(trmmin, 0)
                            'trmmin = Left(trmmin, 2)
                            'dlbTimer & 

                            If (trmmin = 369) Then

                                trmmin = 37

                            End If

                            If (trmmin = 899) Then

                                trmmin = 90

                            End If

                            If (trmmin = 599) Then

                                trmmin = 60

                            End If

                            If (trmmin = 999) Then

                                trmmin = 10

                            End If

                            If (trmmin = 549) Then

                                trmmin = 55

                            End If

                            If (trmmin = 499) Then

                                trmmin = 50

                            End If

                            If (trmmin = 749) Then

                                trmmin = 75

                            End If

                            dlbTimer = dlbTimer & trmmin
                            ' timerkom = timerkom & "/" & dlbTimer
                            ' dlbTimer = 1 'dlbTimer & trmmin

                            'Else


                            'Dim trmmin As Double
                            'Dim trmhrs As String

                            'dlbTimer = Replace(dlbTimer, ".", ",")

                            'If InStr(dlbTimer, ",") = 2 Then '< 10 timer
                            'trmhrs = Left(dlbTimer, 2)



                            'trmmin = (Mid(dlbTimer, 3, 2) * 100) / 60

                            ' Else
                            '    trmhrs = Left(dlbTimer, 3)


                            'trmmin = (Mid(dlbTimer, 4, 2) * 100) / 60
                            'End If

                            'trmmin = FormatNumber(trmmin, 0)

                            'If FormatNumber(trmmin, 0) < 10 Then
                            'dlbTimer = trmhrs & "0" & trmmin
                            'Else
                            'dlbTimer = trmhrs & trmmin
                            'End If

                            'End If




                            ' If CInt(intMedarbId) = 10 Then
                            '    timerkom = timerkom & "/X dlbTimer:  " & dlbTimer
                            'End If

                            '& "." & trmmin trmhrs 

                            'response.write "dlbTimer: " & dlbTimer
                            'response.end()


                        End If

                    Else
                        dlbTimer = 0
                    End If
                Catch ex As Exception
                    Throw New Exception("Get dlbTimer error:" + ex.Message)
                End Try



                Try
                    '** sttid ***'
                    If String.IsNullOrEmpty(ds.Tables(0).Rows(t).Item(13)) = False Then
                        stTid = ds.Tables(0).Rows(t).Item(13)
                    Else
                        stTid = "00:00:00"
                    End If
                Catch ex As Exception
                    Throw New Exception("Get St.tid error:" + ex.Message)
                End Try


                Try
                    '** sltid ***'
                    If String.IsNullOrEmpty(ds.Tables(0).Rows(t).Item(14)) = False Then
                        slTid = ds.Tables(0).Rows(t).Item(14)
                    Else
                        slTid = "00:00:00"
                    End If
                Catch ex As Exception
                    Throw New Exception("Get Sl.tid error:" + ex.Message)
                End Try










                If importFrom = 10 Then 'Excel Upload // 7.7.2014 Bruger stadigvæk den gamle webservice

                    Try
                        If CInt(errThisTOno) = 0 Then


                            '*** tjekker om extsysid findes i forvejen ***'

                            extsysid = 0
                            Dim strSQLext As String = "SELECT extsysid FROM Timer WHERE extsysid = '" & intTempImpId & "' AND origin = " & importFrom
                            'Dim strSQLext As String = "SELECT extsysid FROM Timer WHERE extsysid = " & intTempImpId & " AND origin = " & importFrom
                            objCmd = New OdbcCommand(strSQLext, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                            If objDR.Read() = True Then

                                extsysid = objDR("extsysid")

                            End If

                            objDR.Close()

                            If extsysid <> 0 Then
                                errThisTOno = 6
                            End If

                            '***'

                        End If
                    Catch ex As Exception
                        Throw New Exception("If errThisTOno=0 SELECT extsysid (ozimphours_timetag value: " & intTempImpId & ") FROM Timer error: " + ex.Message)
                    End Try

                End If


                'Return "<br>errThisTOno = " + errThisTOno.ToString + " <br> xdbnavn: " + dbnavn + "<br>cdDato:" + cdDato + "<br>intTempImpId: " + intTempImpId + "<br>dlbTimer: " + dlbTimer + "<br>intJobNr:" + intJobNr + "<br>intMedarbId: " + intMedarbId + "<br>timerkom: " + timerkom + "<br><br>ok"
                'Return "<br>errThisTOno = " + errThisTOno.ToString + " isNull:" + Len(errThisTOno.ToString)



                'If String.IsNullOrEmpty(errThisTOno.ToString()) = False Then
                'errThisTOno = errThisTOno 'Check for null value 
                'Else
                'errThisTOno = 100
                'End If

                'errThisTOno = 0

                If errThisTOno <> 6 Then '<> 6 Then
                    'Return "hej" +errThisTOno
                    'Else
                    '   Return "better"
                    'End if


                    Try
                        '*** Er timer angive korrekt ***'
                        If Len(Trim(dlbTimer)) = 0 Or InStr(dlbTimer, "-") <> 0 Or dlbTimer = 0 Then

                            dlbTimer = 0
                            errThisTOno = 7

                        End If
                    Catch ex As Exception
                        Throw New Exception("Result errThisTOno=7 error: " + ex.Message)
                    End Try


                    Try
                        '*** Finder medarbejder oplysninger ***'
                        'If  String.IsNullOrEmpty(intMedarbId) <> true Then
                        If intMedarbId <> "" Then
                            intMedarbId = intMedarbId
                        Else
                            intMedarbId = ""
                            errThisTOno = 1
                        End If
                    Catch ex As Exception
                        Throw New Exception("Set intMedarbId error: " + ex.Message)
                    End Try


                    Try
                        If CInt(errThisTOno) <> 1 Then

                            'If importFrom = 10 Then 'Execelupload
                            'Dim strSQLme As String = "SELECT mid, mnavn, mnr FROM medarbejdere WHERE init = '" & intMedarbId & "' AND mansat <> 2" & mTypeSQL
                            'Else 'TimeTag
                            Dim strSQLme As String = "SELECT mid, mnavn, mnr FROM medarbejdere WHERE mid = " & intMedarbId & " AND mansat <> 2"
                            'End If

                            objCmd = New OdbcCommand(strSQLme, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                            If objDR.Read() = True Then

                                meNavn = objDR("mnavn")
                                If Len(Trim(meNavn)) <> 0 Then
                                    meNavn = DecodeUTF8(meNavn)
                                End If

                                meNr = objDR("mnr")
                                meID = objDR("mid")

                            End If

                            objDR.Close()

                        End If
                        '***'
                    Catch ex As Exception
                        Throw New Exception("If errThisTOno <> 0 SELECT mid, mnavn, mnr FROM medarbejder error: " + ex.Message)
                    End Try


                    Try
                        If meID <> 0 Then   'meID=1
                            meID = meID
                        Else
                            meID = 0
                            errThisTOno = 11
                        End If
                    Catch ex As Exception
                        Throw New Exception("Set meID error: " + ex.Message)
                    End Try

                    Try
                        '*** Finder job oplysninger ****'
                        If Len(Trim(intJobNr)) <> 0 Then
                            '** Tjekker at jobnr er nummer (Ikke VIetnam)
                            If importFrom <> 1 Then


                                If IsNumeric(intJobNr) Then
                                    intJobNr = intJobNr
                                Else
                                    intJobNr = 0 'intJobNr
                                    errThisTOno = 2
                                End If


                            Else

                                intJobNr = Replace(intJobNr, "_", "")

                            End If


                        Else
                            intJobNr = 0
                            errThisTOno = 2
                        End If
                    Catch ex As Exception
                        Throw New Exception("Set intJobNr error: " + ex.Message)
                    End Try










                    If CInt(errThisTOno) = 0 Then

                        Try
                            jobId = intJobNr '** Skifter, jobnr bliver sat nedenfor 
                            Dim strSQLjob As String = "SELECT id, jobnavn, fastpris, serviceaft, jobknr, valuta, jobnr FROM job WHERE id = " & jobId ' jobnr = '" & intJobNr & "'"
                            objCmd = New OdbcCommand(strSQLjob, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                            If objDR.Read() = True Then

                                'jobNavn = Replace(objDR("jobnavn"), "'", "''")
                                'jobNavn = EncodeUTF8(jobNavn)
                                'jobNavn = DecodeUTF8(jobNavn)
                                jobId = objDR("id")
                                jobFastPris = objDR("fastpris")
                                jobSeraft = objDR("serviceaft")
                                jobKid = objDR("jobknr")
                                intJobNr = objDR("jobnr")
                                'jobValuta = objDR("valuta")                                          

                            End If

                            objDR.Close()
                            '***'
                        Catch ex As Exception
                            Throw New Exception("If errThisTOno=0 SELECT id, jobnavn, fastpris, serviceaft, jobknr, valuta FROM job error: " + ex.Message)
                        End Try


                        Try
                            '** Job ikke er udfyldt korrekt ***'
                            If CInt(jobId) = 0 Then

                                'jobNavn = ""
                                jobId = 0
                                jobFastPris = 0
                                jobSeraft = 0
                                jobKid = 0
                                errThisTOno = 21

                            End If
                        Catch ex As Exception
                            Throw New Exception("If jobId=0 initialize error: " + ex.Message)
                        End Try

                    End If


                    aktId = 0
                    If CInt(errThisTOno) = 0 Then

                        Try

                            '*** Finder akt. oplysninger ****'
                            '*** aktivitet skal være aktiv ****'
                            '*** Fase ??? ***'
                            Dim strSQLa As String
                            strSQLa = "SELECT id, fakturerbar, navn FROM aktiviteter WHERE navn = '" & aktNavn & "' AND job = " & jobId & " AND aktstatus = 1" 'AND navn = '" & aktnavnUse & "'  AND fakturerbar = " & akttypeUse & "" 'Interviewer '


                            objCmd = New OdbcCommand(strSQLa, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                            If objDR.Read() = True Then

                                'aktNavn = objDR("navn")
                                'aktNavn = EncodeUTF8(aktNavn)
                                'aktNavn = DecodeUTF8(aktNavn)
                                aktId = objDR("id")
                                aktFakturerbar = objDR("fakturerbar")

                            End If

                            objDR.Close()

                            '** Blev der fundet en aktivitet
                            If aktId <> 0 Then
                                aktId = aktId
                                errThisTOno = 0
                            Else
                                aktId = 0
                                errThisTOno = 3
                            End If

                            '***'
                        Catch ex As Exception

                            Throw New Exception("If errThisTOno=0 SELECT id, fakturerbar, navn FROM aktiviteter error: " + ex.Message)
                        End Try

                    End If


                    If CInt(errThisTOno) = 0 Then

                        '*** Finder medarb timepris og kostpris ***'
                        '*** Først prøves aktivitet derefter job ***'


                        For mtp = 0 To 1

                            If foundone = "n" Then



                                If mtp = 0 Then
                                    aktIdUse = aktId
                                Else
                                    aktIdUse = 0
                                End If

                                'Dim strSQLer As String = "INSERT INTO timer_imp_err (dato, extsysid, errid, jobid, jobnr, med_init, timeregdato, timer, origin) " _
                                '          & " VALUES " _
                                '           & " ('" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "', '" & intTempImpId & "', "& errThisTOno &", 0,'" & intJobNrTxt & "_" & intJobNr & "', '" & intMedarbId & "_" & tempM & "_" & meID & "', '" & Year(cdDato) & "/" & Month(cdDato) & "/" & Day(cdDato) & "', 5, " & importFrom & ")"


                                'objCmd = New OdbcCommand(strSQLer, objConn)
                                'objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                'objDR.Close()

                                Try
                                    Dim strSQLmtp As String = "SELECT id AS tpid, jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta FROM timepriser WHERE jobid = " & jobId & " AND aktid = " & aktIdUse & " AND medarbid =  " & meID

                                    objCmd = New OdbcCommand(strSQLmtp, objConn)
                                    objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                                    If objDR.Read() = True Then

                                        foundone = "y"
                                        intTimepris = objDR("6timepris")
                                        intValuta = objDR("6valuta")
                                        tpid = objDR("tpid")
                                        timeprisAlt = objDR("timeprisalt")


                                        If CDbl(intTimepris) = 0 And CInt(timeprisAlt) <> 6 Then


                                            Select Case timeprisAlt
                                                Case 1
                                                    timeprisalernativ = "timepris_a1"
                                                    valutaAlt = "tp1_valuta"
                                                Case 2
                                                    timeprisalernativ = "timepris_a2"
                                                    valutaAlt = "tp2_valuta"
                                                Case 3
                                                    timeprisalernativ = "timepris_a3"
                                                    valutaAlt = "tp3_valuta"
                                                Case 4
                                                    timeprisalernativ = "timepris_a4"
                                                    valutaAlt = "tp4_valuta"
                                                Case 5
                                                    timeprisalernativ = "timepris_a5"
                                                    valutaAlt = "tp5_valuta"

                                                Case Else
                                                    timeprisalernativ = "timepris"
                                                    valutaAlt = "tp0_valuta"
                                            End Select



                                            If Len(Trim(timeprisalernativ)) <> 0 And Len(Trim(valutaAlt)) <> 0 Then

                                                Dim strSQL3 As String = "SELECT mid, " & timeprisalernativ & " AS useTimepris, " & valutaAlt & " AS useValuta, " _
                                                & " medarbejdertype FROM medarbejdere, medarbejdertyper WHERE mid =" & meID & " AND medarbejdertyper.id = medarbejdertype"

                                                objCmd = New OdbcCommand(strSQL3, objConn)
                                                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                                                If objDR2.Read() = True Then
                                                    intTimepris = objDR2("useTimepris")
                                                    intValuta = objDR2("useValuta")
                                                End If
                                                objDR2.Close()


                                                If Len(Trim(intTimepris)) <> 0 Then
                                                    intTimepris = Replace(intTimepris, ",", ".")

                                                End If



                                            End If


                                        End If




                                    End If

                                    objDR.Close()
                                    '***'

                                Catch ex As Exception
                                    Throw New Exception("SELECT id AS tpid, jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta FROM timeprise OR SELECT FROM FROM medarbejdere, medarbejdertyper error: " + ex.Message)
                                End Try


                            End If 'foundone

                        Next


                        Try

                            '**************************************************************'
                            '*** Hvis timepris ikke findes på job bruges Gen. timepris fra '
                            '*** Fra medarbejdertype, og den oprettes på job **************'
                            '**************************************************************'
                            tprisGen = 0
                            valutaGen = 1

                            Dim SQLmedtpris As String = "SELECT medarbejdertype, timepris, tp0_valuta, kostpris, mnavn FROM medarbejdere, medarbejdertyper " _
                            & " WHERE Mid = " & meID & " AND medarbejdertyper.id = medarbejdertype"

                            objCmd = New OdbcCommand(SQLmedtpris, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                            If objDR.Read() = True Then

                                If objDR("kostpris") <> 0 Then
                                    kostpris = objDR("kostpris")
                                Else
                                    kostpris = 0
                                End If

                                tprisGen = objDR("timepris")
                                valutaGen = objDR("tp0_valuta")

                            End If

                            objDR.Close()
                            '** Slut timepris **
                        Catch ex As Exception
                            Throw New Exception(" Hvis timepris ikke findes på job bruges Gen, SELECT FROM medarbejdere, medarbejdertyper error: " + ex.Message)
                        End Try

                        Try
                            '**** Opdaterer timepris på job ***'
                            If foundone = "n" Then

                                intTimepris = Replace(tprisGen, ",", ".")
                                intValuta = valutaGen

                                Dim strSQLtpris As String = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) " _
                                & " VALUES (" & jobId & ", 0, " & meID & ", 0, " & intTimepris & ", " & intValuta & ")"

                                objCmd = New OdbcCommand(strSQLtpris, objConn)
                                objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                objDR.Close()

                            End If
                        Catch ex As Exception
                            Throw New Exception(" Opdaterer timepris på job, INSERT INTO timepriser error: " + ex.Message)
                        End Try


                        Try
                            '**** Tjekker om akt. er sat til fast timepris for alle medarbjedere ***'
                            '**** overruler medarbejer timepris på akt. og job *********************'
                            Dim brug_fasttp As Integer = 0
                            Dim brug_fastkp As Integer = 0
                            Dim fasttp As Double = 0
                            Dim fastkp As Double = 0
                            Dim fasttp_val As Integer = 0
                            '*** Tjekker om aktiviteten er sat til ens timpris for alle medarbejdere (overskriver medarbejderens egen timepris)
                            Dim strSQLtjkAktTp As String = "SELECT brug_fasttp, brug_fastkp, fasttp, fasttp_val, fastkp, fastkp_val FROM aktiviteter WHERE id = " & aktIdUse

                            objCmd = New OdbcCommand(strSQLtjkAktTp, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                            If objDR.Read() = True Then

                                If CInt(objDR("brug_fasttp")) = 1 Then
                                    brug_fasttp = 1
                                    fasttp = objDR("fasttp")
                                    fasttp = Replace(fasttp, ".", "")
                                    fasttp = Replace(fasttp, ",", ".")

                                    fasttp_val = objDR("fasttp_val")
                                End If


                                If CInt(objDR("brug_fastkp")) = 1 Then
                                    brug_fastkp = 1
                                    fastkp = objDR("fastkp")
                                    fastkp = Replace(fastkp, ".", "")
                                    fastkp = Replace(fastkp, ",", ".")
                                End If

                            End If
                            objDR.Close()



                            If CInt(brug_fasttp) = 1 Then
                                intTimepris = fasttp
                                intValuta = fasttp_val
                            End If

                            If CInt(brug_fastkp) = 1 Then
                                kostpris = fastkp
                            End If

                        Catch ex As Exception
                            Throw New Exception("Tjekker om akt. er sat til fast timepris for alle medarbjedere, SELECT brug_fasttp, brug_fastkp, fasttp, fasttp_val, fastkp, fastkp_val FROM aktiviteter error: " + ex.Message)
                        End Try

                    End If

                    Try

                        If CInt(errThisTOno) = 0 Then

                            '** Finder valuta og kurs **'
                            If Len(Trim(intValuta)) <> 0 And intValuta <> 0 Then
                                intValuta = intValuta
                            Else
                                intValuta = 0
                                errThisTOno = 4
                            End If

                            Dim strSQLv As String = "SELECT kurs FROM valutaer WHERE id = " & intValuta
                            objCmd = New OdbcCommand(strSQLv, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                            If objDR.Read() = True Then

                                kurs = objDR("kurs")

                            End If

                            objDR.Close()

                        End If
                        '***'		       
                    Catch ex As Exception
                        Throw New Exception("If errThisTOno = 0 SELECT kurs FROM valutaer error: " + ex.Message)
                    End Try

                    Try
                        If CInt(errThisTOno) = 0 Then

                            '*** Finder kunde oplysninger ****'
                            If Len(Trim(jobKid)) <> 0 And jobKid <> 0 Then
                                jobKid = jobKid
                            Else
                                jobKid = 0
                                errThisTOno = 5
                            End If

                            Dim strSQLk As String = "SELECT kkundenavn, kkundenr FROM kunder WHERE kid = " & jobKid
                            objCmd = New OdbcCommand(strSQLk, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                            If objDR.Read() = True Then

                                kNavn = objDR("kkundenavn")
                                kNavn = EncodeUTF8(kNavn)
                                kNavn = DecodeUTF8(kNavn)
                                kNr = jobKid 'objDR("kkundenr")


                            End If

                            objDR.Close()
                            '***'

                            If Len(Trim(kNavn)) = 0 Then
                                errThisTOno = 51
                            End If

                        End If

                    Catch ex As Exception
                        Throw New Exception("If errThisTOno = 0, SELECT kkundenavn, kkundenr FROM kunder error: " + ex.Message)
                    End Try

                    Try
                        '***** er uge lukket '***
                        If CInt(errThisTOno) = 0 Then

                            '** Manger at tilføje månedafslutning ***'
                            ugeErLukket = 0
                            Dim strSQLk As String = "SELECT mid, uge FROM ugestatus WHERE (WEEK(uge, 1) = WEEK('" & cdDatoSQL & "', 1) AND YEAR(uge) = YEAR('" & cdDatoSQL & "')) AND mid = " & meID
                            objCmd = New OdbcCommand(strSQLk, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                            If objDR.Read() = True Then

                                ugeErLukket = 1


                            End If

                            objDR.Close()
                            '***'


                            '**** IGNORER UGE lukket midlertidigt *****'
                            If ugeErLukket = 1 Then
                                errThisTOno = 9
                            End If

                        End If
                    Catch ex As Exception
                        Throw New Exception("er uge lukket, SELECT mid, uge FROM usestatus error: " + ex.Message)
                    End Try

                    'errThisTOno = 0


                    dlbTimer = Replace(dlbTimer, ",", ".")


                    Call jq_format(timerkom)
                    timerkom = jq_formatTxt

















                    '*********************************************************************************************************************'
                    '*********************************************************************************************************************'
                    '***********    Indlæser timer
                    '*********************************************************************************************************************'
                    '*********************************************************************************************************************'

                    If CInt(errThisTOno) = 0 Then
                        Try




                        Catch ex As Exception
                            Throw New Exception("jq_format error: " + ex.Message)
                        End Try

                        Try


                            Select dbnavn
                                Case "wilke", "intranet" 'SKLA rettes
                                    'intTempImpId = 0
                                    intTempImpId = 99
                                Case Else
                                    intTempImpId = extsysid
                            End Select


                            '*** Indlæser Timer ***'
                            Dim strSQL As String = "INSERT INTO timer " _
                            & "(" _
                            & " timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, timerkom, " _
                            & " TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, " _
                            & " editor, kostpris, seraft, " _
                            & " valuta, kurs, extSysId, sttid, sltid, origin" _
                            & ") " _
                            & " VALUES " _
                            & " (" _
                            & dlbTimer & ", " & aktFakturerbar & ", " _
                            & "'" & cdDatoSQL & "', " _
                            & "'" & meNavn & "', " _
                            & meID & ", " _
                            & "'" & jobNavn & "', " _
                            & "'" & intJobNr & "', " _
                            & "'" & kNavn & "'," _
                            & kNr & ", " _
                            & "'" & timerkom & "', " _
                            & aktId & ", " _
                            & "'" & aktNavn & "', " _
                            & Year(Now) & ", " _
                            & intTimepris & ", " _
                            & "'" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "', " _
                            & jobFastPris & ", " _
                            & "'00:00:01', " _
                            & "'TimeTag', " _
                            & Replace(kostpris, ",", ".") & ", " _
                            & jobSeraft & ", " _
                            & intValuta & ", " _
                            & Replace(kurs, ",", ".") & ", '" & intTempImpId & "', '" & stTid & "', '" & slTid & "', " & importFrom & ")"


                            '** Mangler salgs og kost priser ***'
                            'Return strSQL & "<br><br>"

                            'Response.write(strSQL)

                            'objCmd = New OdbcCommand(strSQL, objConn)
                            'intCountInserted += objCmd.ExecuteNonQuery() '(CommandBehavior.closeConnection)

                            'Dim strSQL2 As String = "DELETE FROM timer_imp_err WHERE extsysid = '" & intTempImpId & "' AND origin = " & importFrom

                            objCmd = New OdbcCommand(strSQL, objConn)
                            objCmd.ExecuteNonQuery() '(CommandBehavior.closeConnection)


                            'errThisTOnoAll = errThisTOnoAll

                        Catch ex As Exception
                            Throw New Exception("If errThisTOno = 0, Indlæser Timer (fejl i insert): " + ex.Message)
                        End Try


                    Else '** Indlæser ind til ErrLog // Timer der skal efterbehandles på ugesedlen 

                        Try





                            Select Case errThisTOno
                                Case 1, 11
                                    errTxtThis = "Medarbejder ikke fundet"
                                Case 2, 21
                                    errTxtThis = "Job ikke fundet"
                                Case 3, 31
                                    errTxtThis = "Aktivitet ikke fundet"
                                Case 4
                                    errTxtThis = "Valuta ikke fundet på job"
                                Case 5, 51
                                    errTxtThis = "Kunde ikke fundet på job"
                                Case 6
                                    errTxtThis = "Record er allerede indlæst en gang"
                                Case 7
                                    errTxtThis = "Timer ikke angivet korrekt"
                                Case 8
                                    errTxtThis = "Der er fejl i dato formatet"
                                Case 9
                                    errTxtThis = "Periode er lukket"
                            End Select


                            If importFrom = 10 Then
                                errThisTOnoStr = errThisTOnoStr & "Fejl: " & errThisTOno.ToString() & " - " & errTxtThis & " (jobnr: " & intJobNrTxt & ", medarb.: " & intMedarbId & ")<br>"
                            End If




                            If importFrom = 10 Or importFrom = 11 Then 'TimeTag, indlæser til timer_import_temp, hvis der ikke var et direkte match på job og aktivitet






                                '****** TESTER ****'
                                'Try


                                '*** Test indlæser ellae records i err db ***'

                                'Dim strSQLerrTest As String = "INSERT INTO timer_imp_err (dato, extsysid, errid, jobid, jobnr, med_init, timeregdato, timer, origin) " _
                                '& " VALUES " _
                                '& " ('" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "', 0, 99, 0,'" & intJobNrTxt & " t:" & dlbTimer & " dt: " & Year(cdDato) & "/" & Month(cdDato) & "/" & Day(cdDato) & " sysID: " & intTempImpId & " medid: " & intMedarbId & "', 'SK', '2014-03-06', 2, 11)"

                                'Dim strSQLerrTest As String = "INSERT INTO timer_imp_err (errid, dato, jobnr, timer, med_init, origin) VALUES (" & errThisTOno & ", '" & cdDatoSQL & "', '" & ds.Tables(0).Rows(t).Item(8) & "', '" & dlbTimer & "', '" & intMedarbId & "', "& importFrom &")"



                                'objCmd = New OdbcCommand(strSQLerrTest, objConn)
                                'objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                'objDR.Close()

                                'alternativ
                                'objCmd = New OdbcCommand(strSQLerrTest, objConn)
                                'objCmd.ExecuteNonQuery() '(CommandBehavior.closeConnection)



                                'Return "Er i loopet:  linjer: " + ds.Tables(0).Rows.Count.ToString() + " importFrom: " + importFrom + " strSQLerrTest: " + strSQLerrTest


                                'Catch ex As Exception

                                'Return "Erro of hi Therwe " + ex.Message + ": " + strSQLerrTest
                                'End Try







                                Dim strSQLer As String = "INSERT INTO timer_import_temp (dato, origin, medarbejderid, jobid, aktnavn, timer, tdato, timerkom, lto, editor, overfort, jobnavn, aktid, errid, errmsg) " _
                                & " VALUES('" & cdDatoSQL & "'," & importFrom & ",'" & meID & "'," & jobId & ",'" & aktNavn & "'," & dlbTimer & "," _
                                & "'" & cdDatoSQL & "','" & timerkom & "','" & lto & "','" & editorIn & "',0,'" & jobNavn & "', " & aktId & "," & errThisTOno & ",'" & errTxtThis & "')"

                                objCmd = New OdbcCommand(strSQLer, objConn)
                                objCmd.ExecuteNonQuery() '(CommandBehavior.closeConnection)





                            End If


                        Catch ex As Exception
                            Throw New Exception("If errThisTOno <> 0, INSERT INTO timer_imp_err error: " + ex.Message)
                        End Try


                    End If


                End If ' 6 findes i forvejen

                t = t + 1

            Next

        Next

        'Catch

        'Return ex.Message + "<br>ErrNO: " & Err.Number & " ErrDesc: " & Err.Description & " # " & errThisTOno & " # CATI recordID: (" & intTempImpId & ")"
        'Exit Try

        ' '''Finally            
        'End Try


        'rtErrTxt(errThisTOnoAll)
        'errThisTOnoAll = ""
        'Return errThisTOnoAll

        '*** Mail ***'

        If err_jobnr <> String.Empty Then
            err_jobnr = err_jobnr.Substring(0, err_jobnr.Length - 2)
        End If



        'Dim errThisTOnoStr As String = errThisTOno.ToString()
        Return "succes " + intCountInserted.ToString() + " linje(r) indlæst. jobnr: " + err_jobnr + " errid:" + errThisTOnoStr

    End Function



    'Public Function rtErrTxt(ByVal erTxt As String) As String
    ' Dim eTxt As String
    '    eTxt = erTxt

    '   Return eTxt
    'End Function


End Class






