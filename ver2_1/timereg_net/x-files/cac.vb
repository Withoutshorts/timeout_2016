Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.Odbc
Imports System.Web.UI
Imports System.Web.UI.WebControls

Public Class CAC : Inherits Page





    Public errThisAll As String = ""

    Public intMedarbId As String
    Public intCatiId, intJobNr As String
    Public dlbTimer As String
    Public cdDato As String

    Public errThis As Integer = 0

    Public meId As Integer
    Public meNavn As String
    Public meNr As Integer

    Public jobNavn As String
    'Public jobId As Integer
    Public jobFastPris As Integer
    Public jobSeraft As Integer
    Public jobKid As Integer
    Public jobValuta As Integer

    Public intjobId As Integer

    Public aktNavn As String
    Public aktId As Integer = 0
    Public aktFakturerbar As Integer

    Public kurs As String

    Public kNavn As String
    Public kNr As String

    Public extsysid As Integer

    Public intTimepris As String = 0
    Public timeprisAlt As Integer = 0
    Public intValuta As Integer
    Public tpid As Integer
    Public foundone As String = "n"
    Public timeprisalernativ, valutaAlt As String
    Public mtp As Integer
    Public aktIdUse As Integer

    Public kostpris As String
    Public tprisGen As String
    Public valutaGen As Integer = 1

    Public komm As String = ""
    Public strSQLaTXT As String = ""



    Public Function upDateCatiTimer2(ByVal ds As DataSet) As String



        'Dim con As New SqlConnection(Application("MyDatabaseConnectionString"))
        'Dim daCust As New SqlDataAdapter("Select * From tblCustomer", con)
        'Dim cbCust As New SqlCommandBuilder(daCust)
        'daCust.Update(ds, "Cust")
        'Return ds

        Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=81.19.249.35;Database=timeout_epi;User=outzource;Password=SKba200473;"
        'Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=81.19.249.35;Database=timeout_epi_catitest;User=outzource;Password=SKba200473;"
        'Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_epi_catitest;User=outzource;Password=SKba200473;"
        'Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_intranet;User=outzource;Password=SKba200473;"

        Dim objConn As OdbcConnection
        Dim objCmd As OdbcCommand
        'Dim objDataSet As New DataSet
        Dim objDR As OdbcDataReader
        Dim objDR2 As OdbcDataReader
        Dim dt As DataTable
        Dim dr As DataRow

        Dim objDR4 As OdbcDataReader



        '** Åbner Connection ***'
        objConn = New OdbcConnection(strConn)
        objConn.Open()

        'Dim objTable As DataTable

        Dim t As Integer = 0

        For Each dt In ds.Tables
            For Each dr In dt.Rows


                errThis = 0
                meId = 0
                extsysid = 0
                aktId = 0
                intJobNr = 0
                intjobId = 0
                intValuta = 0




                intCatiId = ds.Tables("job_imp_timer").Rows(t).Item("id")
                intMedarbId = ds.Tables("job_imp_timer").Rows(t).Item("medarbejderid")

                intJobNr = ds.Tables("job_imp_timer").Rows(t).Item("jobid")
                intJobNr = Trim(intJobNr)


                dlbTimer = ds.Tables("job_imp_timer").Rows(t).Item("timer")
                cdDato = ds.Tables("job_imp_timer").Rows(t).Item("dato")


                komm = ds.Tables("job_imp_timer").Rows(t).Item("komm")

                'Dim indtastningfindes AS integer = 0 

                '*** Finder medarbejder oplysninger ***'
                If intMedarbId <> "" Then
                    intMedarbId = intMedarbId
                Else
                    intMedarbId = 0
                    errThis = 1
                End If



                If errThis <> 1 Then


                    Dim strSQLme As String = "SELECT mid, mnavn, mnr FROM medarbejdere WHERE init = '" & intMedarbId & "'"
                    'mid = " & intMedarbId
                    objCmd = New OdbcCommand(strSQLme, objConn)
                    objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                    If objDR.Read() = True Then

                        meNavn = objDR("mnavn")
                        meNr = objDR("mnr")
                        meId = objDR("mid")

                    End If

                    objDR.Close()


                    '***'

                    'Return meId
                    'Return True

                    If CInt(meId) <> 0 Then
                        meId = meId
                    Else
                        '*** Tidligere ansatte konsulenter ***
                        meId = 356
                        meNavn = "Forhv. ans. konsulenter (CAC)"
                        meNr = "1000002"

                        errThis = 0
                    End If


                End If

                '*** Finder job oplysninger ****'
                If intJobNr <> "" Then
                    intJobNr = intJobNr
                Else
                    intJobNr = 0
                    errThis = 2
                End If




                Dim strSQLjob As String = "SELECT id, jobnavn, fastpris, serviceaft, jobknr, valuta FROM job WHERE jobnr = " & intJobNr
                objCmd = New OdbcCommand(strSQLjob, objConn)
                objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                If objDR.Read() = True Then

                    jobNavn = Replace(objDR("jobnavn"), "'", "''")
                    intjobId = objDR("id")
                    jobFastPris = objDR("fastpris")
                    jobSeraft = objDR("serviceaft")
                    jobKid = objDR("jobknr")
                    'jobValuta = objDR("valuta")



                End If

                objDR.Close()
                '***'



                If CInt(intjobId) = 0 Then
                    errThis = 21
                End If

                'jobNavn = ""
                ' ' intjobId = 0
                ' jobFastPris = 0
                'jobSeraft = 0
                'jobKid = 0
                'intJobNr = 0
                'errThis = 2
                'Else

                'jobNavn = jobNavn 'objDR("jobnavn")
                'intjobId = intjobId 'objDR("id")
                'jobFastPris = jobFastPris 'objDR("fastpris")
                'jobSeraft = jobSeraft ' objDR("serviceaft")
                'jobKid = jobKid 'objDR("jobknr")

                'End If





                If errThis = 0 Then

                    '*** Finder akt. oplysninger ****'
                    '*** Fase ??? ***'

                    
                    Dim strSQLa As String = "SELECT id, fakturerbar, navn, job FROM aktiviteter WHERE job = " & intjobId & " AND aktstatus = 1 AND navn LIKE 'CAC%'"
                    'AND navn LIKE 'CAC%'"
                    'C indlæste timer
                    '200 " '" & intjobId & "" ' AND navn LIKE 'C%' AND aktstatus = 1 navn = 'Interviewer'
                    objCmd = New OdbcCommand(strSQLa, objConn)
                    objDR4 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                    If objDR4.Read() = True Then

                        aktNavn = objDR4("navn")
                        aktId = objDR4("id")
                        aktFakturerbar = objDR4("fakturerbar")


                    End If

                    objDR4.Close()

                    '** Blev der fundet en aktivitet
                    If aktId <> 0 Then
                        aktId = aktId
                    Else
                        aktId = 0
                        errThis = 3
                    End If

                    'strSQLaTXT = strSQLaTXT & "<br>" & "(" & intjobId & ") - atk: " & aktNavn & " ;; " & strSQLa

                    '***'

                End If




                'strSQLaTXT = strSQLaTXT & "<br>" & "(" & intjobId & ") - atk: " & aktNavn & " ;; " & strSQLmtp

                'Return True



                If errThis = 0 Then

                    '*** Finder medarb timepris og kostpris ***'
                    '*** Først prøves aktivitet derefter job ***'


                    For mtp = 0 To 1

                        If foundone = "n" Then



                            If mtp = 0 Then
                                aktIdUse = aktId
                            Else
                                aktIdUse = 0
                            End If


                            Dim strSQLmtp As String = "SELECT id AS tpid, jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta FROM timepriser WHERE jobid = " & intjobId & " AND aktid = " & aktId & " AND medarbid =  " & meId

                            'strSQLaTXT = strSQLmtp
                            'Return strSQLaTXT
                            'Return True


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





                                    Dim strSQL3 As String = "SELECT mid, " & timeprisalernativ & " AS useTimepris, " & valutaAlt & " AS useValuta, " _
                                    & " medarbejdertype FROM medarbejdere, medarbejdertyper WHERE mid =" & meId & " AND medarbejdertyper.id = medarbejdertype"

                                    objCmd = New OdbcCommand(strSQL3, objConn)
                                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                                    If objDR2.Read() = True Then
                                        intTimepris = objDR2("useTimepris")
                                        intValuta = objDR2("useValuta")
                                    End If
                                    objDR2.Close()


                                    intTimepris = Replace(intTimepris, ",", ".")





                                End If




                            End If

                            objDR.Close()
                            '***'

                        End If 'foundone

                    Next





                    '**************************************************************'
                    '*** Hvis timepris ikke findes på job bruges Gen. timepris fra '
                    '*** Fra medarbejdertype, og den oprettes på job **************'
                    '**************************************************************'
                    Dim SQLmedtpris As String = "SELECT medarbejdertype, timepris, tp0_valuta, kostpris, mnavn FROM medarbejdere, medarbejdertyper " _
                    & " WHERE Mid = " & meId & " AND medarbejdertyper.id = medarbejdertype"

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





                    '**** Opdaterer timepris på job ***'
                    If foundone = "n" Then
                        intTimepris = Replace(tprisGen, ",", ".")
                        intValuta = valutaGen

                        Dim strSQLtpris As String = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) " _
                        & " VALUES (" & intjobId & ", 0, " & meId & ", 0, " & intTimepris & ", " & intValuta & ")"



                        objCmd = New OdbcCommand(strSQLtpris, objConn)
                        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)



                    End If





                End If






                If errThis = 0 Then


                    '** Finder valuta og kurs **'
                    If intValuta <> 0 Then
                        intValuta = intValuta
                    Else
                        intValuta = 0
                        errThis = 4
                    End If


                    Dim strSQLv As String = "SELECT kurs FROM valutaer WHERE id = " & intValuta
                    objCmd = New OdbcCommand(strSQLv, objConn)
                    objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                    If objDR.Read() = True Then

                        kurs = objDR("kurs")

                    End If

                    objDR.Close()
                    '***'		            


                End If



                If errThis = 0 Then

                    '*** Finder kunde oplysninger ****'
                    If jobKid <> 0 Then
                        jobKid = jobKid
                    Else
                        jobKid = 0
                        errThis = 5
                    End If


                    Dim strSQLk As String = "SELECT kkundenavn, kkundenr FROM kunder WHERE kid = " & jobKid
                    objCmd = New OdbcCommand(strSQLk, objConn)
                    objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                    If objDR.Read() = True Then

                        kNavn = objDR("kkundenavn")
                        kNr = objDR("kkundenr")


                    End If

                    objDR.Close()
                    '***'

                End If


                If errThis = 0 Then

                    '*** tjekker om extsysid findes i forvejen ***'
                    Dim strSQLext As String = "SELECT extsysid FROM Timer WHERE extsysid = " & intCatiId
                    objCmd = New OdbcCommand(strSQLext, objConn)
                    objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                    If objDR.Read() = True Then

                        extsysid = objDR("extsysid")

                    End If

                    objDR.Close()

                    If extsysid <> 0 Then
                        errThis = 6
                    End If

                    '***'

                End If



                If CInt(errThis) = 0 Then

                    '*** Indlæser Timer ***'
                    Dim strSQL As String = "INSERT INTO timer " _
                    & "(" _
                    & " timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, " _
                    & " TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, " _
                    & " editor, kostpris, seraft, " _
                    & " valuta, kurs, extSysId, timerkom, sttid, sltid " _
                    & ") " _
                    & " VALUES " _
                    & " (" _
                    & Replace(dlbTimer, ",", ".") & ", " & aktFakturerbar & ", " _
                    & "'" & Year(cdDato) & "/" & Month(cdDato) & "/" & Day(cdDato) & "', " _
                    & "'" & meNavn & "', " _
                    & meId & ", " _
                    & "'" & jobNavn & "', " _
                    & intJobNr & ", " _
                    & "'" & kNavn & "', " _
                    & kNr & ", " _
                    & aktId & ", " _
                    & "'" & aktNavn & "', " _
                    & Year(Now) & ", " _
                    & intTimepris & ", " _
                    & "'" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "', " _
                    & jobFastPris & ", " _
                    & "'00:00:01', " _
                    & "'CaTi Import', " _
                    & Replace(kostpris, ",", ".") & ", " _
                    & jobSeraft & ", " _
                    & intValuta & ", " _
                    & Replace(kurs, ",", ".") & ", " & intCatiId & ", '" & Replace(komm, "'", "''") & "', '00:00:00', '00:00:00')"


                    '** Manger salgs og kost priser ***'

                    'Response.write(strSQL)

                    objCmd = New OdbcCommand(strSQL, objConn)
                    objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    objDR.Close()

                    errThisAll = errThisAll

                Else '** Indlæser ind til ErrLog


                    Dim strSQLer As String = "INSERT INTO timer_imp_err (dato, extsysid, errid) " _
                    & " VALUES " _
                    & " ('" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "', " & intCatiId & ", " & errThis & ")"


                    objCmd = New OdbcCommand(strSQLer, objConn)
                    objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    objDR.Close()

                    'Response.write(strSQLer)

                    errThisAll = errThisAll & ";" & errThis & "(" & intCatiId & ")"

                End If




                t = t + 1

            Next

        Next



        Return errThisAll
        Return strSQLaTxt


    End Function



End Class







