<%@ WebService language="VB" class="to_import_hours" %>
Imports System
Imports System.Web.Services
Imports System.Data
Imports System.Data.Odbc


Imports System.Globalization
Imports System.Threading

Public Class to_import_hours

    Public opdtp As Integer = 0
    Public lto As String = ""
    Public testValue As String
    Public testmode As Integer = 0

    Public timerkom As String
    Public intMedarbId As String
    Public intTempImpId, intJobNr As String
    Public dlbTimer As String
    Public cdDato As Date
    Public importFrom As String

    Public errThisTOno As Integer = 0

    Public meNavn As String
    Public meNr As String
    Public meID As Integer
    Public medarbejdertypeThis AS Integer = 0
    Public mTypeSQL As String

    Public jobNavn As String
    Public jobId As Double
    Public jobFastPris As Integer
    Public jobSeraft As Double
    Public jobKid As Double
    Public jobValuta As Integer

    Public aktNavn, aktnavnUse As String
    Public aktId As Double = 0
    Public aktFakturerbar As Integer

    Public kurs As String

    Public kNavn As String
    Public kNr As Double

    Public extsysid As Double

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
    Public kp1_valuta As Integer = 1
    Public kp1_valuta_kurs As Double = 100

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

    Public Ite As Integer = 0

    'Public errThisTOnoAll As Integer = 0

    Public errThisTOnoStr As String = "<br>Fejlkoder for linjer der IKKE blev indlæst:<br>"
    Public errTxtThis As String


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


    <WebMethod()> Public Function timeout_importTimer_rack(ByVal ds As DataSet) As String

        'On Error Resume Next




        'Dim con As New SqlConnection(Application("MyDatabaseConnectionString"))
        'Dim daCust As New SqlDataAdapter("Select * From tblCustomer", con)
        'Dim cbCust As New SqlCommandBuilder(daCust)
        'daCust.Update(ds, "Cust")
        'Return ds        

        Dim objConn As OdbcConnection
        Dim objConn2 As OdbcConnection
        Dim objCmd As OdbcCommand
        'Dim objDataSet As New DataSet
        Dim objDR As OdbcDataReader
        Dim objDR2 As OdbcDataReader
        Dim dt As DataTable
        Dim dr As DataRow

        Dim strConn As String
        'strConn2 = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_epi;User=outzource;Password=SKba200473;"
        'strConn2 = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_oko;User=to_outzource2;Password=SKba200473;"
        'strConn2 = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=root;pwd=;database=timeout_epi; OPTION=32"
        'strConn = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_epi;User=to_outzource2;Password=SKba200473;"
        'strConn = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=outzource; pwd=SKba200473; database=timeout_epi;"
        'strConn2 = "timeout_epi64"

        'Dim objTable As DataTable

        Dim t As Double = 0

        'Try


        For Each dt In ds.Tables
            For Each dr In dt.Rows






                ' Dim strSQLerrTest2 As String = "INSERT INTO timer_imp_err (dato, extsysid, errid) VALUES ('2016/06/10', '1111', " + t + ")"

                '         objCmd = New OdbcCommand(strSQLerrTest2, objConn2)
                '        objCmd.ExecuteReader '(CommandBehavior.closeConnection)



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
                kp1_valuta = 1
                kp1_valuta_kurs = 100

                'If ds.Tables("timer_import_temp").Columns('0') Then '.Contains("origin")
                'importFrom = ds.Tables("timer_import_temp").Rows(t).Item("origin")
                'importFrom = ds.Tables(0).Rows(t).Item(3)
                'Else
                importFrom = 10 '3
                'End If

                dbnavn = ""

                'Return "OK HEJ DER"  'ex.Message + "<br>ErrNO: " & Err.Number & " ErrDesc: " & Err.Description & " # " & errThisTOno & " # CATI recordID: (" & intTempImpId & ")"

                '*** DB FELTER RÆKKEFØLGE
                'intTempImpId    = 0
                'tempM          = 4
                'intJobNr       = 5
                'aktnavnUse      = 6
                'dldTimer       = 7
                'Tdato           = 8
                'LTO             = 9
                'timerkom        = 10



                Try
                    'If ds.Tables("timer_import_temp").Columns.Contains("lto") Then
                    If String.IsNullOrEmpty(ds.Tables(0).Rows(t).Item(9)) = False Then


                        dbnavn = ds.Tables(0).Rows(t).Item(9)
                    Else
                        dbnavn = "DBIKKEFUNDET"
                    End If

                    'dbnavn = "epi"

                    If dbnavn = "outz" Or dbnavn = "intranet - local" Then
                        dbnavn = "intranet"
                    Else
                        dbnavn = dbnavn
                    End If

                    'end if


                    lto = dbnavn

                Catch ex As Exception
                    Throw New Exception("Get dbnavn error: " + ex.Message)
                End Try


                '**** Conn string ****'

                If t = 0 Then

                    Try

                        'importFrom = "3" Then 'MMMI / ekstern
                        'dbnavn = "intranet"
                        '81.19.249.35
                        'strConn = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_"+ dbnavn +";User=to_outzource2;Password=SKba200473;"
                        'strConn = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=outzource; pwd=SKba200473; database=timeout_epi;"
                        'strConn = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_epi;User=to_outzource2;Password=SKba200473;"
                        'strConn = "timeout_" & dbnavn & "64"
                        strConn = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" + dbnavn + ";Uid=to_outzource2;Password=SKba200473;"


                        '** Åbner Connection ***'
                        objConn = New OdbcConnection(strConn)
                        objConn.Open()

                        '** Åbner Connection ***'
                        objConn2 = New OdbcConnection(strConn)
                        objConn2.Open()

                    Catch ex As Exception
                        Throw New Exception("If t=0 database connect and open error: " + ex.Message)
                    End Try


                End If






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

                    'Dim dt_len As Integer = Len(cdDato)
                    'Dim dt_left As String = Left(cdDato, dt_len - 4)
                    'Dim dt_right As String = Right(cdDato, 4)
                    'Dim newdtUS As String =dt_right+"/"+ dt_left                        

                    'cdDatoSQL = newdtUS
                    cdDatoSQL = Year(cdDato) & "/" & Month(cdDato) & "/" & Day(cdDato)

                Catch ex As Exception
                    Throw New Exception("Get cdDato error: " + ex.Message)
                End Try



                '*** Test indlæser ellae records i err db ***'

                'Dim strSQLer As String = "INSERT INTO timer_imp_err (dato, extsysid, errid, jobid, jobnr, med_init, timeregdato, timer, origin) " _
                '& " VALUES " _
                '& " ('" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "', 0, 99, 0,'" & intJobNrTxt & " t:" & dlbTimer & " dt: " & Year(cdDato) & "/" & Month(cdDato) & "/" & Day(cdDato) & " sysID: " & intTempImpId & " medid: " & intMedarbId & "', '0', '2000-01-01', 0, 1)"

                'Dim strSQLerrTest As String = "INSERT INTO timer_imp_err (dato, jobnr, timer, med_init) VALUES ('" & cdDato & "', '" & intJobNr & "', '" & dlbTimer & "', '" & intMedarbId & "')"

                'Dim strSQLerrTest As String = "INSERT INTO timer_imp_err (dato, extsysid, errid) VALUES ('" & cdDatoSQL & "', '" & cdDato & "', '1')"

                'objCmd = New OdbcCommand(strSQLerrTest, objConn)
                'objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                'objDR.Close()

                'errThisTOno = 6
                '*** end **'


                Try
                    '*** if record findes i timereg i forvejen ****'
                    intTempImpId = ds.Tables(0).Rows(t).Item(0) '0 
                Catch ex As Exception
                    Throw New Exception("Get intTempImpId error:" + ex.Message)
                End Try



                Try
                    '**** Aktivitets navn ****
                    aktnavnUse = ""
                    'If ds.Tables("timer_import_temp").Columns.Contains("aktnavn") Then
                    If String.IsNullOrEmpty(ds.Tables(0).Rows(t).Item(6)) = False Then
                        aktnavnUse = ds.Tables(0).Rows(t).Item(6)
                        aktnavnUse = aktnavnUse.ToString()
                    End If
                Catch ex As Exception
                    Throw New Exception("Get aktnavnUse error:" + ex.Message)
                End Try


                Try
                    '** Kommentar **' 11??
                    'timerkom = ""
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
                    '** Jobnr **'
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
                    '** Timer ***'
                    If String.IsNullOrEmpty(ds.Tables(0).Rows(t).Item(7)) = False Then
                        dlbTimer = ds.Tables(0).Rows(t).Item(7)
                    Else
                        dlbTimer = 0
                    End If
                Catch ex As Exception
                    Throw New Exception("Get dlbTimer error:" + ex.Message)
                End Try



                'Return "<br>HEJ DER 9: dbnavn: " + dbnavn + " cdDato: " + cdDato

                '*************************************************************************************************************************
                '****** DATA MODATGET FRA RECORDSET  --> Begynder indlæsning og datamodel ************************************************



                Try
                    If CInt(errThisTOno) = 0 Then


                        '*** tjekker om extsysid findes i forvejen ***'

                        extsysid = 0
                        Dim strSQLext As String = "SELECT extsysid FROM Timer WHERE extsysid = '" & intTempImpId & "' AND origin = " & importFrom
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
                    Throw New Exception("If errThisTOno=0 ozimphours SELECT extsysid FROM Timer error: " + ex.Message)
                End Try



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
                        If Len(Trim(dlbTimer)) = 0 Or dlbTimer = 0 Then '** Or InStr(dlbTimer, "-") <> 0

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

                        medarbejdertypeThis = 0
                        If CInt(errThisTOno) <> 1 Then

                            Dim strSQLme As String = "SELECT mid, mnavn, mnr, medarbejdertype FROM medarbejdere WHERE init = '" & intMedarbId & "' AND mansat <> 2 "

                            objCmd = New OdbcCommand(strSQLme, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                            If objDR.Read() = True Then

                                meNavn = objDR("mnavn")
                                If Len(Trim(meNavn)) <> 0 Then
                                    meNavn = DecodeUTF8(meNavn)
                                End If

                                meNr = objDR("mnr")
                                meID = objDR("mid")
                                medarbejdertypeThis = objDR("medarbejdertype")


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
                            If importFrom <> "1" Then


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
                            Dim strSQLjob As String = "SELECT id, jobnavn, fastpris, serviceaft, jobknr, valuta FROM job WHERE jobnr = '" & intJobNr & "'"
                            objCmd = New OdbcCommand(strSQLjob, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                            If objDR.Read() = True Then

                                jobNavn = Replace(objDR("jobnavn"), "'", "''")
                                jobNavn = EncodeUTF8(jobNavn)
                                jobNavn = DecodeUTF8(jobNavn)
                                jobId = objDR("id")
                                jobFastPris = objDR("fastpris")
                                jobSeraft = objDR("serviceaft")
                                jobKid = objDR("jobknr")
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

                                jobNavn = ""
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
                            strSQLa = "SELECT id, fakturerbar, navn FROM aktiviteter WHERE job = " & jobId & " AND navn = '" & aktnavnUse & "' AND aktstatus = 1" ' AND fakturerbar = " & akttypeUse & "" 'Interviewer '


                            objCmd = New OdbcCommand(strSQLa, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                            If objDR.Read() = True Then

                                aktNavn = objDR("navn")
                                aktNavn = EncodeUTF8(aktNavn)
                                aktNavn = DecodeUTF8(aktNavn)
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

                    End If 'CInt(errThisTOno) = 0 


                    If CInt(errThisTOno) = 0 Then
                        '***************************************************************************************************
                        '*** Finder medarb timepris og kostpris ***'
                        '*** Først prøves aktivitet derefter job ***'
                        '***************************************************************************************************
                        tprisGen = 0
                        'Ite = 1
                        'If Ite = 1 Then
                        If InStr(lto, "epi") = 1 And intMedarbId = "Asia" Then



                            If lto = "epi" Or lto = "epi2017" Then
                                'EPI hardcoded Vietnametimer **'
                                valutaGen = 9 '1
                                intTimepris = 11
                                kostpris = 11

                            End If

                            If lto = "epi_no" Then
                                'EPI NO hardcoded Vietnametimer **'
                                valutaGen = 1
                                intTimepris = 180
                                kostpris = 60
                            End If

                            tprisGen = intTimepris
                            intValuta = valutaGen
                            foundone = "y"




                        Else

                            For mtp = 0 To 1

                                If foundone = "n" Then 'Hvis fundet på aktivitet, søges der ikke på job



                                    If mtp = 0 Then
                                        aktIdUse = aktId
                                    Else
                                        aktIdUse = 0
                                    End If



                                    Try





                                        Dim strSQLmtp As String = "SELECT id AS tpid, jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta FROM timepriser WHERE jobid = " & jobId & " AND aktid = " & aktIdUse & " AND medarbid =  " & meID

                                        objCmd = New OdbcCommand(strSQLmtp, objConn2)
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

                                                    objCmd = New OdbcCommand(strSQL3, objConn2)
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
                                        Throw New Exception("SELECT id AS tpid, jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta FROM timeprise OR SELECT FROM medarbejdere, medarbejdertyper error: " + ex.Message)
                                    End Try


                                End If 'foundone

                            Next



                            Try




                                '**************************************************************'
                                '*** Hvis timepris ikke findes på job bruges Gen. timepris fra '
                                '*** Fra medarbejdertype, og den oprettes på job **************'
                                '**************************************************************'

                                'foundone = foundone
                                'intValuta = valutaGen

                                '*** IKKE EPINION tp for stor

                                If foundone = "n" Then
                                    Dim SQLmedtpris As String = "SELECT timepris AS useTimepris, tp0_valuta, kostpris, kp1_valuta FROM medarbejdertyper WHERE id = " & medarbejdertypeThis

                                    objCmd = New OdbcCommand(SQLmedtpris, objConn2)
                                    objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                                    If objDR.Read() = True Then

                                        If objDR("kostpris") <> 0 Then
                                            kostpris = objDR("kostpris")
                                        Else
                                            kostpris = 0
                                        End If

                                        kp1_valuta = objDR("kp1_valuta")

                                        tprisGen = objDR("useTimepris")
                                        valutaGen = objDR("tp0_valuta")


                                        'Ite = 1
                                        'If Ite = 1 Then 'KUN EPI 'InStr(dbnavn, "epi") = 1 And
                                        intTimepris = objDR("useTimepris")
                                        intValuta = valutaGen 'objDR2("useValuta")

                                        If Len(Trim(intTimepris)) <> 0 Then
                                            intTimepris = Replace(intTimepris, ",", ".")

                                        End If

                                        'End If


                                    End If

                                    objDR.Close()
                                    '** Slut timepris **

                                End If 'foundone

                            Catch ex As Exception
                                Throw New Exception("medarbejdertypeThis: " + medarbejdertypeThis.ToString + " Hvis timepris ikke findes på job bruges Gen, SELECT FROM medarbejdere, medarbejdertyper error: " + ex.Message)
                            End Try





                            Try



                                opdtp = 0
                                If opdtp = 1 Then
                                    '**** Opdaterer timepris på job ***'
                                    '**** ALDRIG timepriser vokser eksponentielt ***'

                                    intTimepris = Replace(tprisGen, ",", ".")
                                    intValuta = valutaGen

                                    Dim strSQLtpris As String = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) " _
                                    & " VALUES (" & jobId & ", 0, " & meID & ", 0, " & intTimepris & ", " & intValuta & ")"

                                    objCmd = New OdbcCommand(strSQLtpris, objConn2)
                                    objCmd.ExecuteReader() '(CommandBehavior.closeConnection)

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
                                Dim strSQLtjkAktTp As String = "SELECT brug_fasttp, brug_fastkp, fasttp, fasttp_val, fastkp, fastkp_val FROM aktiviteter WHERE id = " & aktId

                                objCmd = New OdbcCommand(strSQLtjkAktTp, objConn2)
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


                        End If 'EPi + ASIA

                    End If 'CInt(errThisTOno) = 0 




                    Try

                        If CInt(errThisTOno) = 0 Then

                            '** Finder valuta og kurs **'
                            If Len(Trim(intValuta)) <> 0 And intValuta <> 0 And Len(Trim(kp1_valuta)) <> 0 And kp1_valuta <> 0 Then
                                intValuta = intValuta
                                kp1_valuta = kp1_valuta
                            Else
                                intValuta = 0
                                kp1_valuta = 0
                                errThisTOno = 4
                            End If

                            Dim strSQLv As String = "SELECT kurs FROM valutaer WHERE id = " & intValuta
                            objCmd = New OdbcCommand(strSQLv, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                            If objDR.Read() = True Then

                                kurs = objDR("kurs")

                            End If

                            objDR.Close()

                            '** Kostprisvaluta **'
                            Dim strSQLvk As String = "SELECT kurs FROM valutaer WHERE id = " & kp1_valuta
                            objCmd = New OdbcCommand(strSQLvk, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                            If objDR.Read() = True Then

                                kp1_valuta_kurs = objDR("kurs")

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


                            ugeErLukket = 0
                            Dim strSQLk As String = "SELECT mid, uge FROM ugestatus WHERE (WEEK(uge) = WEEK('" & cdDatoSQL & "') AND YEAR(uge) = YEAR('" & cdDatoSQL & "')) AND mid = " & meID
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



                    'End If 'errThisTOno = 0



                    '**** INDLÆSER TIMER   

                    If CInt(errThisTOno) = 0 Then
                        Try

                            dlbTimer = Replace(dlbTimer, ",", ".")


                            Call jq_format(timerkom)
                            timerkom = jq_formatTxt

                            'Call jq_format(kNavn)
                            'kNavn = jq_formatTxt

                            'Call jq_format(jobNavn)
                            'jobNavn = jq_formatTxt

                            'Call DecodeUTF8(jobNavn)

                        Catch ex As Exception
                            Throw New Exception("jq_format error: " + ex.Message)
                        End Try

                        Try

                            '*** Indlæser Timer ***'
                            Dim strSQL As String = "INSERT INTO timer " _
                            & "(" _
                            & " timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, timerkom, " _
                            & " TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, " _
                            & " editor, kostpris, seraft, " _
                            & " valuta, kurs, extSysId, sttid, sltid, origin, kpvaluta, kpvaluta_kurs" _
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
                            & "'Excel Import', " _
                            & Replace(kostpris, ",", ".") & ", " _
                            & jobSeraft & ", " _
                            & intValuta & ", " _
                            & Replace(kurs, ",", ".") & ", '" & intTempImpId & "', '00:00:00', '00:00:00', " & importFrom & ", " & kp1_valuta & ", " & Replace(kp1_valuta_kurs, ",", ".") & ")"


                            '** Manger salgs og kost priser ***'
                            'Return strSQL & "<br><br>"

                            'Response.write(strSQL)

                            objCmd = New OdbcCommand(strSQL, objConn)
                            'objCmd.ExecuteNonQuery()

                            intCountInserted += objCmd.ExecuteNonQuery() '(CommandBehavior.closeConnection)
                            'objDR.Close()

                            Dim strSQL2 As String = "DELETE FROM timer_imp_err WHERE extsysid = '" & intTempImpId & "' AND origin = " & importFrom

                            objCmd = New OdbcCommand(strSQL2, objConn)
                            objCmd.ExecuteNonQuery()
                            'CommandBehavior.CloseConnection()
                            'objCmd.closeConnection()



                            'errThisTOnoAll = errThisTOnoAll

                        Catch ex As Exception
                            Throw New Exception("If errThisTOno = 0, Indlæser Timer, INSERT INTO timer OR DELETE FROM timer_imp_err error: " + ex.Message)
                        End Try





                    Else '** Indlæser ind til ErrLog 

                        Try

                            err_jobnr += intJobNr.ToString() + ", "

                            Dim strSQLer As String = "INSERT INTO timer_imp_err (dato, extsysid, errid, jobid, jobnr, med_init, timeregdato, timer, origin) " _
                                & " VALUES " _
                                & " ('" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "', '" & intTempImpId & "', " + errThisTOno.ToString + ", " & jobId & ",'" & intJobNrTxt & "', '" & intMedarbId & "', '" & cdDatoSQL & "', " & dlbTimer & ", " & importFrom & ")"

                            objCmd = New OdbcCommand(strSQLer, objConn)
                            objCmd.ExecuteNonQuery() '(CommandBehavior.closeConnection)

                            'objDR.Close()

                            'Response.write(strSQLer)

                            'errThisTOnoAll = errThisTOnoAll & ";" & errThisTOno & "(" & intTempImpId & ")"

                            'errThisTOnoAll = errThisTOnoAll + 1    

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

                            errThisTOnoStr = errThisTOnoStr & "Fejl: " & errThisTOno.ToString() & " - " & errTxtThis & " (jobnr: " & intJobNrTxt & ", medarb.: " & intMedarbId & ")<br>"

                        Catch ex As Exception
                            Throw New Exception("If errThisTOno <> 0, INSERT INTO timer_imp_err error: " + ex.Message)
                        End Try


                    End If


                End If ' 6 findes i forvejen

                t = t + 1

                'Return "<br>Antal records: : " & t

            Next

        Next





        Dim strSQLupdtemp As String = "UPDATE timer_import_temp SET overfort = 1 WHERE overfort = 0"

        objCmd = New OdbcCommand(strSQLupdtemp, objConn)
        objCmd.ExecuteNonQuery()
        'objCmd.closeConnection()

        objConn.Close()
        objConn2.Close()

        'Catch

        'Return "OK HEJ DER"  'ex.Message + "<br>ErrNO: " & Err.Number & " ErrDesc: " & Err.Description & " # " & errThisTOno & " # CATI recordID: (" & intTempImpId & ")"

        'Exit Try

        '*** Finally            
        'End Try

        'rtErrTxt(errThisTOnoAll)
        'errThisTOnoAll = ""
        'Return errThisTOnoAll

        '*** Mail ***'

        If err_jobnr <> String.Empty Then
            err_jobnr = err_jobnr.Substring(0, err_jobnr.Length - 2)
        End If



        'Dim errThisTOnoStr As String = errThisTOno.ToString()
        Return "Succes " + intCountInserted.ToString() + " linje(r) indlæst.<br><br>Du kan lukke denne side ned nu. [<a href=""Javascript:window.close();"">X</a>]"  'jobnr: " + err_jobnr + " errid:" + errThisTOnoStr



    End Function



    'Public Function rtErrTxt(ByVal erTxt As String) As String
    ' Dim eTxt As String
    '    eTxt = erTxt

    '   Return eTxt
    'End Function





End Class






