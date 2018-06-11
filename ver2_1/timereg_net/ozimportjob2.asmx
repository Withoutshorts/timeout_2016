



<%@ WebService language="VB" class="oz_importjob2" %>
Imports System
Imports System.Web.Services
Imports System.Data
Imports System.Data.Odbc

'***************************************************************************************
'*** DENNE SERVICE BENYTTES BL.A AF 
'*** OKO NAV
'*** Wilke (bruger ozimportjob2_ds)
'*** DENCKER Monitor
'*** TIA
'***************************************************************************************



Public Class oz_importjob2


    'errThis = 11 jobnr findes i forvejen







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
    Function jq_format(ByVal jq_str)

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



    Public editor As String
    Public id As Integer = 0
    Public dato As Date
    Public jobstartdato As Date
    Public jobslutdato As Date
    Public origin As String
    Public jobnavn As String
    Public jobnr As String = "0"
    Public jobans As Integer = 0
    Public lto As String
    Public importtype As String
    Public beskrivelse As String
    Public errThis As Integer = 0
    Public kid As Integer

    Public jobStatus As Integer = 1

    Public kunderef As Integer = 0

    Public rekvisitionsnr As String
    Public bruttooms As Double
    Public internnote As String = ""
    Public jobansInit As String = ""
    Public strjobnr As String = "0"

    Public strantalArbPakker As String = "0"

    Public tprisGen As String = "0"
    Public valutaGen As Integer = 1
    Public kostpris As Double
    Public intTimepris As String = "0"
    Public intValuta As Integer = valutaGen

    Public lastAktID As Integer
    Public lastID As Integer

    Public timeprisalt As Integer
    Public kundenavnTxt As String = ""
    Public opretJobOk As Integer = 0
    Public jobnrTjk As String = "0"
    Public jobID As Integer = 0
    Public timereguseJobFindes As Integer = 0
    Public projektKategori As Integer = 0
    Public tpFundet As Integer = 0
    Public aktStatus As Integer = 1

    Public aktFields As String = ""

    Public aktnavn As String = ""
    Public aktstdato As Date
    Public aktsldato As Date
    Public fomr As Integer = 0
    Public sort As Integer
    Public aktvarenr As String = ""
    Public antalstk As Double = 0
    Public lastJobnr As String = "0"
    Public kundenrTxt As String = "0"

    Public intCountInserted As Integer = 0
    Public strAktFase As String = ""

    Public kpers, rekvnr As String
    Public progrpId As Integer = 0
    Public progrp As String = "0"

    <WebMethod()> Public Function createjob2(ByVal ds As DataSet) As String



        'On Error Resume Next


        'Dim con As New SqlConnection(Application("MyDatabaseConnectionString"))
        'Dim daCust As New SqlDataAdapter("Select * From tblCustomer", con)
        'Dim cbCust As New SqlCommandBuilder(daCust)
        'daCust.Update(ds, "Cust")
        'Return ds


        'Fra EXCEL / timer_import_temp
        'lto = ds.Tables("tb_to_var").Rows(0).Item("lto")

        'Return "HEJ: "& lto

        lto = "demo"

        Dim dt As DataTable
        Dim dr As DataRow

        For Each dt In ds.Tables
            For Each dr In dt.Rows


                lto = ds.Tables("tb_to_var").Rows(0).Item("lto")
                importtype = ds.Tables("tb_to_var").Rows(0).Item("importtype")

            Next

        Next




        Dim objConn As OdbcConnection
        Dim objConn2 As OdbcConnection
        Dim objConn3 As OdbcConnection

        Dim objCmd As OdbcCommand
        Dim objCmd2 As OdbcCommand
        'Dim objDataSet As New DataSet
        Dim objDR As OdbcDataReader
        Dim objDR2 As OdbcDataReader
        Dim objDR3 As OdbcDataReader
        Dim objDR6 As OdbcDataReader
        Dim objDR4 As OdbcDataReader
        'Dim dt As DataTable
        'Dim dr As DataRow

        Dim orderBySQL As String = ""
        Dim strConn As String
        Dim t As Double = 0

        'Try


        'If lto = "dencker" Then
        ' lto = "dencker_test"
        'End If



        strConn = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" & lto & ";User=to_outzource2;Password=SKba200473;"
        'strConn = "Driver={MySQL ODBC 3.51 Driver};Server=195.189.130.210;Database=timeout_oko;User=to_outzource2;Password=SKba200473;"
        'strConn = "Provider=MSDASQL;driver={MySQL ODBC 3.51 Driver}; Server=195.189.130.210; Port=3306; User=outzource; Password=SKba200473; Database=timeout_oko; Option=3;"

        'strConn = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=root;pwd=;database=timeout_oko; OPTION=32"
        'strConn = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=root;pwd=;database=timeout_wilke; OPTION=32"





        '** Åbner Connection ***'
        objConn = New OdbcConnection(strConn)
        objConn.Open()

        objConn2 = New OdbcConnection(strConn)
        objConn2.Open()

        'Try

        'Dim strSQLjnjUpd As String = "UPDATE job_import_temp SET aktnavn = '" + importtype + "' WHERE id > 0 AND overfort = 0 AND errid = 0 "
        'objCmd = New OdbcCommand(strSQLjnjUpd, objConn)
        'objCmd.ExecuteReader() '(CommandBehavior.closeConnection)


        'Return "Succes + imptype: " + importtype + "#"
        'Catch ex As Exception
        'Throw New Exception("Get dlbTimer error:" + ex.Message)
        'End Try

        'Exit Function


        If (importtype = "d1" Or importtype = "t1") Then

            If (importtype = "d1") Then
                '*** KUN FØRSTE linje 4 for hver job. Linje 4 = jobnavn
                Dim strSQLjnj As String = "SELECT id, dato, editor, origin, jobnavn, jobnr, jobstartdato, jobslutdato, jobans, lto, "
                strSQLjnj += "beskrivelse, kundenavn, kundenr, kpers, rekvnr FROM job_import_temp "
                strSQLjnj += "WHERE id > 0 And overfort = 0 AND jobnavn <> '' And errid = 0 And beskrivelse = 4 GROUP BY jobnr ORDER BY jobnr"
                'Dim strSQLjnj As String = "Select id FROM job_import_temp WHERE id > 0 And overfort = 10 And errid = 0 " & orderBySQL
                objCmd = New OdbcCommand(strSQLjnj, objConn)
                objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)


            Else 'TIA t1

                '*** KUN FØRSTE linje 4 for hver job. Linje 4 = jobnavn
                Dim strSQLjnj As String = "Select id, dato, editor, origin, jobnavn, jobnr, jobstartdato, jobslutdato, jobans, lto, beskrivelse, kundenavn, kundenr, projgrp FROM job_import_temp WHERE id > 0 And overfort = 0 AND jobnavn <> '' And errid = 0  GROUP BY jobnr ORDER BY jobnr"
                'Dim strSQLjnj As String = "Select id FROM job_import_temp WHERE id > 0 And overfort = 10 And errid = 0 " & orderBySQL
                objCmd = New OdbcCommand(strSQLjnj, objConn)
                objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

            End If

        Else

            Select Case lto
                Case "dencker", "dencker_test"
                    orderBySQL = " ORDER BY jobnr, sort"
                    aktFields = ", aktnavn, aktstdato, aktsldato, fomr, sort, aktvarenr, stkantal, kundenavn"
                Case Else
                    orderBySQL = " ORDER BY id"
                    aktFields = ""
            End Select


            '*** HENTER JOB FRA JOB_IMPORT_TEMP '****
            Dim strSQLjnj As String = "Select id, dato, editor, origin, jobnavn, jobnr, jobstartdato, jobslutdato, jobans, lto, beskrivelse " & aktFields & " FROM job_import_temp WHERE id > 0 And overfort = 0 AND jobnavn <> '' And errid = 0 " & orderBySQL
            'Dim strSQLjnj As String = "SELECT id FROM job_import_temp WHERE id > 0 And overfort = 10 And errid = 0 " & orderBySQL
            objCmd = New OdbcCommand(strSQLjnj, objConn)
            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

        End If





        While objDR.Read() = True


            If (importtype = "d1") Then

                aktnavn = ""
                aktstdato = "2002-01-01"
                aktsldato = "2002-01-01"
                fomr = 0
                sort = 0
                aktvarenr = ""
                antalstk = 0


                kunderef = 0
                kpers = objDR("kpers")
                rekvisitionsnr = objDR("rekvnr")

            Else


                Select Case lto
                    Case "dencker", "dencker_test"
                        aktnavn = objDR("aktnavn")
                        aktstdato = objDR("aktstdato")
                        aktsldato = objDR("aktsldato")
                        fomr = 0 'objDR("fomr")
                        sort = objDR("sort")
                        aktvarenr = objDR("aktvarenr")
                        antalstk = objDR("stkantal")


                    Case Else
                        aktnavn = ""
                        aktstdato = "2002-01-01"
                        aktsldato = "2002-01-01"
                        fomr = 0
                        sort = 0
                        aktvarenr = ""
                        antalstk = 0
                End Select

                kpers = ""
                kunderef = 0
                rekvisitionsnr = 0

            End If


            errThis = 0

            id = objDR("id")
            dato = objDR("dato")

            origin = objDR("origin")
            editor = objDR("editor")

            'jobnavn = EncodeUTF8(objDR("jobnavn"))
            'jobnavn = DecodeUTF8(jobnavn)
            'jobnavn = jq_format(objDR("jobnavn"))
            'jobnavn = jq_formatTxt
            jobnavn = Replace(objDR("jobnavn"), "'", "")
            jobnavn = Replace(jobnavn, "''", "")
            jobnavn = Replace(jobnavn, ";", "")

            jobnr = objDR("jobnr")
            jobnrTjk = objDR("jobnr")

            jobstartdato = objDR("jobstartdato")
            jobslutdato = objDR("jobslutdato")

            jobansInit = objDR("jobans")


            lto = objDR("lto")



            beskrivelse = IsDBNull(objDR("beskrivelse")) '** VARCHAR FOR IKKE AT FEJLE (TEXT fejler)
            If beskrivelse <> True Then
                beskrivelse = objDR("beskrivelse")
                beskrivelse = beskrivelse.ToString()
            Else
                beskrivelse = ""
            End If



            If (importtype = "d1" Or importtype = "t1") Then 'BLA TIA t1

                kundenavnTxt = objDR("kundenavn")
                kundenrTxt = objDR("kundenr")

                If (importtype = "t1") Then
                    progrp = objDR("projgrp")


                    Dim strSQLprogrp As String = "Select id FROM projektgrupper WHERE navn LIKE '" & progrp & "%'"
                    objCmd = New OdbcCommand(strSQLprogrp, objConn)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    progrpId = 1
                    If objDR2.Read() = True Then

                        progrpId = objDR2("id")

                    End If


                Else
                    progrp = 0
                    progrpId = 1
                End If


            Else

                Select Case lto
                    Case "dencker", "dencker_test"
                        kundenavnTxt = objDR("kundenavn")
                        kundenrTxt = "0"
                    Case Else
                        kundenavnTxt = beskrivelse
                        kundenrTxt = jobnr
                End Select


                progrp = 0
                progrpId = 1

            End If



            '**** Henter Jobans '****
            Dim strSQLminit As String
            If (importtype = "d1") Then
                strSQLminit = " mnavn = '" + jobansInit + "'"
            Else
                strSQLminit = " init = '" + jobansInit + "'"
            End If

            Dim strSQLjobans As String = "Select mid FROM medarbejdere WHERE " + strSQLminit
            objCmd = New OdbcCommand(strSQLjobans, objConn)
            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            jobans = 0
            If objDR2.Read() = True Then

                jobans = objDR2("mid")

            End If
            objDR2.Close()

            '**** Jobnr *****
            If (importtype = "d1") Then
                strjobnr = jobnr.ToString
            Else

                Select Case lto
                    Case "dencker", "dencker_test"
                        strjobnr = jobnr.ToString

                        Dim intPos_ As Integer = InStr(strjobnr, "-")
                        Dim intLngt As Integer = strjobnr.Length

                        If (intPos_ > 1) Then

                            Dim strjobnr_opr As String = strjobnr
                            strjobnr = Left(strjobnr, intPos_ - 1)

                            If (intLngt - intPos_) > 1 Then 'Hvis det er pos 10-99
                                strAktFase = "Pos-" + Right(strjobnr_opr, 2) + "-" + jobnavn 'strjobnr
                            Else
                                strAktFase = "Pos-0" + Right(strjobnr_opr, 1) + "-" + jobnavn 'strjobnr
                            End If



                        Else



                            strAktFase = "Pos-01-" + jobnavn 'strjobnr

                        End If

                    Case Else
                        strjobnr = jobnr.ToString
                End Select


            End If


            '**** FOMR ****'
            If (importtype = "d1") Or lto = "dencker" Then 'Dencker 'Både job og aktviviteter

                Select Case Left(jobnavn, 1)

                    Case "P"
                        fomr = 23 'Plastic
                    Case "T"

                        If Left(jobnavn, 2) = "TV" Then
                            fomr = 1 'værktøj
                        Else
                            fomr = 3 'Tool
                        End If

                    Case "S"

                        fomr = 2 'Service
                    Case "3"

                        fomr = 27' 3D
                    Case "O"
                        fomr = 25 'orange

                    Case "I"
                        fomr = 31
                    Case Else

                        fomr = fomr

                End Select

            Else

                fomr = fomr

            End If

            '*** Tjekker om jobnr findes ***
            Dim opdaterJob As Integer = 0
            Dim strSQLjobnrfindes As String = "Select jobnr FROM job WHERE jobnr = '" + strjobnr + "'"
            objCmd = New OdbcCommand(strSQLjobnrfindes, objConn)
            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

            If objDR2.Read() = True Then

                errThis = 0
                opdaterJob = 1

            End If
            objDR2.Close()


            '********************************************************************************
            '*** Opretter Kunde ***'
            '********************************************************************************
            Select Case lto
                Case "oko"
                    kid = 1 ' ALTID ØKOLOGISK
                Case Else

                    '*** Tjekker om kundenavn findes ***
                    kid = 0

                    Dim strSQKfindkunde As String = "SELECT kid FROM kunder WHERE kkundenavn = '" + kundenavnTxt + "'"
                    objCmd = New OdbcCommand(strSQKfindkunde, objConn)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                    If objDR2.Read() = True Then

                        errThis = 0
                        kid = objDR2("kid")

                    End If
                    objDR2.Close()

                    If kid = 0 Then

                        '*** OPRETTER KUNDE **
                        Dim strSQKkundeins As String = "INSERT INTO kunder SET kkundenavn = '" + kundenavnTxt + "', kkundenr = '" + kundenrTxt + "', ketype = 'ke', ktype = 0"
                        objCmd = New OdbcCommand(strSQKkundeins, objConn)
                        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                        objDR2.Close()


                        '** HENTER ID PÅ NETOP oprettede ***'
                        Dim strSQKfindkundeNy As String = "SELECT kid FROM kunder WHERE kid <> 0 ORDER BY kid DESC LIMIT 1"
                        objCmd = New OdbcCommand(strSQKfindkundeNy, objConn)
                        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                        If objDR2.Read() = True Then


                            kid = objDR2("kid")

                        End If
                        objDR2.Close()

                    Else '** Opdaterer kundenavn

                        If (importtype = "d1") And (lto = "dencker" Or lto = "dencker_test") Then


                            Dim strSQKkundeupd As String = "UPDATE kunder SET kkundenavn = '" + kundenavnTxt + "' WHERE kkundenr = '" + kundenrTxt + "'"
                            objCmd = New OdbcCommand(strSQKkundeupd, objConn)
                            objCmd.ExecuteReader() '(CommandBehavior.closeConnection)

                        End If


                    End If


                    '*** Kontaktperson ***
                    If (importtype = "d1") And (lto = "dencker" Or lto = "dencker_test") Then

                        '** Tjekker om Kpers findes ***'
                        kunderef = 0
                        Dim strSQKfindKpers As String = "SELECT id FROM kontaktpers WHERE navn = '" & kpers & "' AND kundeid = " & kid & " ORDER BY id DESC LIMIT 1"
                        objCmd = New OdbcCommand(strSQKfindKpers, objConn)
                        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                        If objDR2.Read() = True Then


                            kunderef = objDR2("id")

                        End If
                        objDR2.Close()


                        If CInt(kunderef) = 0 Then

                            Dim strSQKkperIns As String = "INSERT INTO kontaktpers (kundeid, navn) VALUES (" & kid & ", '" & kpers & "')"
                            objCmd = New OdbcCommand(strSQKkperIns, objConn)
                            objCmd.ExecuteReader() '(CommandBehavior.closeConnection)


                            Dim strSQKfindKpersNy As String = "SELECT id FROM kontaktpers WHERE navn = '" & kpers & "' AND kundeid = " & kid & " ORDER BY id DESC LIMIT 1"
                            objCmd = New OdbcCommand(strSQKfindKpersNy, objConn)
                            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                            If objDR2.Read() = True Then


                                kunderef = objDR2("id")

                            End If
                            objDR2.Close()

                        End If

                    End If



            End Select



            bruttooms = 0
            internnote = ""



            'Return "Webservice Msg dt: " & jobstartdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture)

            'Call DecodeUTF8(jobnavn)
            'jobnavn = 

            '********************************************************************************
            '*** Opretter Job ***'
            '********************************************************************************
            errThis = 0
            If CInt(errThis) = 0 Then





                'Return "Webservice Msg dt jobnrTjk: "+ jobnrTjk +" opretJobOk: "+ opretJobOk 
                If (importtype = "d1" And beskrivelse = "4") Or (lto <> "dencker" And lto <> "dencker_test") Or importtype = "t1" Then
                    opretJobOk = 1
                Else
                    opretJobOk = 0
                End If


                If importtype = "t1" Or lto = "tia" Then

                    If LCase(beskrivelse) = "all" Then
                        jobStatus = 0
                    Else
                        jobStatus = 1
                    End If
                Else

                    If importtype = "d1" Or lto = "dencker" Then
                        jobStatus = 3
                    Else
                        jobStatus = 1
                    End If

                End If



                If CInt(opretJobOk) = 1 Then





                    If CInt(opdaterJob) = 1 Then 'opdater

                        '*** Opdater stadivæk OK? kommer an på hvilke linje SORT (Dencker) der blicver læst fra excel filen
                        Select Case lto
                            Case "dencker", "dencker_test"
                                If (lastJobnr <> strjobnr) Then 'CInt(sort) = 10
                                    opdaterJob = opdaterJob
                                Else
                                    opdaterJob = 0


                                End If
                            Case Else
                                opdaterJob = opdaterJob
                        End Select



                        If CInt(opdaterJob) = 1 Then 'opdater

                            Select Case lto
                                Case "dencker", "dencker_test"


                                    '** OPDATER job ed først 4 type
                                    '** NULSTILLER beskrivelse = '', bliver opdateret af de næste type 4 linjer

                                    '*** Opdaterer beskrivelse
                                    Dim jobBesk As String = ""
                                    Dim strSQKjobbesk As String = "SELECT jobnavn FROM job_import_temp WHERE jobnr = '" & jobnr & "' AND beskrivelse = 4 AND overfort = 0 ORDER BY id"
                                    objCmd = New OdbcCommand(strSQKjobbesk, objConn)
                                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                                    While objDR2.Read() = True


                                        jobBesk = jobBesk + "<br>" + objDR2("jobnavn")

                                    End While
                                    objDR2.Close()



                                    Dim strSQLjobUpd As String = ("Update job SET jobnavn = '" & jobnavn & "', jobnr = '" & jobnr & "', " _
                                    & " editor = '" & editor & "', " _
                                    & " dato = '" & dato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', beskrivelse = '" & jobBesk & "' WHERE jobnr = '" & jobnr & "'")

                                    objCmd = New OdbcCommand(strSQLjobUpd, objConn)
                                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                    objDR2.Close()

                                Case Else

                                    Dim strSQLjobUpd As String = "Update job SET jobnavn = '" & jobnavn & "', jobnr = '" & jobnr & "', jobstatus = " & jobStatus & ", " _
                                    & " jobstartdato = '" & jobstartdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "'," _
                                    & " jobslutdato = '" & jobslutdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', editor = '" & editor & "', " _
                                    & " dato = '" & dato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', beskrivelse = '" & beskrivelse & "',  jobans1 = " & jobans & ", " _
                                    & " kundekpers = " & kunderef & ""

                                    If importtype = "t1" Or lto = "tia" Then
                                        strSQLjobUpd += ", projektgruppe2 = " & progrpId
                                    End If

                                    strSQLjobUpd += " WHERE jobnr = '" & jobnr & "'"

                                    objCmd = New OdbcCommand(strSQLjobUpd, objConn)
                                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                    objDR2.Close()

                            End Select



                            '*** Tilføjer til timereg_usejob (for en sikkerhedsskyld)

                            '*** Finder jobid ***
                            jobID = 0



                            '**** Sikrer timereg usejob bliver udfyldt ved rediger job ****
                            Dim strSQLlastJobID As String = "SELECT id FROM job WHERE jobnr = '" & jobnr & "'"
                            objCmd = New OdbcCommand(strSQLlastJobID, objConn)
                            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                            If objDR2.Read() = True Then

                                jobID = objDR2("id")

                            End If
                            objDR2.Close()


                        End If


                        '**************************************************************'
                        '*** Hvis timepris ikke findes på job bruges Gen. timepris fra '
                        '*** Fra medarbejdertype, og den oprettes på job              *'
                        '*** BLIVER ALTID HENTET FRA Medarb.type for ØKO              *'
                        '**************************************************************'
                        Select Case lto
                            Case "oko"

                                objConn3 = New OdbcConnection(strConn)
                                objConn3.Open()

                                tprisGen = 0
                                valutaGen = 1
                                kostpris = 0
                                intTimepris = 0
                                intValuta = valutaGen

                                Dim SQLmedtpris As String = "SELECT medarbejdertype, timepris, timepris_a2, timepris_a3, tp0_valuta, kostpris, mnavn, mid FROM medarbejdere " _
                                & " LEFT JOIN medarbejdertyper ON (medarbejdertyper.id = medarbejdertype) " _
                                & " WHERE mid <> 0 AND mansat = 1 AND medarbejdertyper.id = medarbejdertype GROUP BY mid"

                                objCmd = New OdbcCommand(SQLmedtpris, objConn3)
                                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                                While objDR2.Read() = True

                                    If objDR2("kostpris") <> 0 Then
                                        kostpris = objDR2("kostpris")
                                    Else
                                        kostpris = 0
                                    End If


                                    Select Case lto
                                        Case "oko"


                                            Select Case Left(Trim(beskrivelse), 8)
                                                Case "NaturErh"

                                                    If objDR2("timepris_a2") <> 0 Then

                                                        tprisGen = objDR2("timepris_a2")
                                                    Else
                                                        tprisGen = 0
                                                    End If

                                                    timeprisalt = 2

                                                Case "Direktor"

                                                    If objDR2("timepris_a3") <> 0 Then

                                                        tprisGen = objDR2("timepris_a3")
                                                    Else
                                                        tprisGen = 0
                                                    End If

                                                    timeprisalt = 3

                                                Case Else

                                                    If objDR2("timepris") <> 0 Then

                                                        tprisGen = objDR2("timepris")
                                                    Else
                                                        tprisGen = 0
                                                    End If

                                                    timeprisalt = 0

                                            End Select


                                        Case Else

                                            If objDR2("timepris") <> 0 Then

                                                tprisGen = objDR2("timepris")
                                            Else
                                                tprisGen = 0
                                            End If

                                            timeprisalt = 0




                                    End Select 'lto




                                    valutaGen = objDR2("tp0_valuta")


                                    '**** Indlæser timepris på aktiviteter ***'
                                    intTimepris = tprisGen
                                    intTimepris = Replace(intTimepris, ",", ".")
                                    intValuta = valutaGen
                                    '** Slut timepris **


                                    '*** OPRETTER AKTIVITETER ***'
                                    Select Case lto
                                        Case "dencker", "dencker_test"
                                            '** DO NOTHING **
                                        Case Else




                                            '*** HENTER Aktiviteter **'
                                            Dim strSQLAkt As String = "SELECT id AS aktid FROM aktiviteter WHERE job =" & jobID & " AND fakturerbar = 1"
                                            objCmd2 = New OdbcCommand(strSQLAkt, objConn3)
                                            objDR3 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)
                                            While objDR3.Read() = True






                                                '*** OPDATERER / INS Timepris **'
                                                Dim strSQLtprisDEL As String = "SELECT id FROM timepriser WHERE jobid = " & jobID & " AND aktid = " & objDR3("aktid") & " AND medarbid = " & objDR2("mid")

                                                tpFundet = 0

                                                objCmd2 = New OdbcCommand(strSQLtprisDEL, objConn3)
                                                objDR6 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)

                                                If objDR6.Read() = True Then

                                                    tpFundet = 1

                                                End If
                                                objDR6.Close()

                                                If CInt(tpFundet) = 0 Then

                                                    Dim strSQLtprisIns As String = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) " _
                                                    & " VALUES (" & jobID & ", " & objDR3("aktid") & ", " & objDR2("mid") & ", " & timeprisalt & ", " & intTimepris & ", " & intValuta & ")"
                                                    objCmd2 = New OdbcCommand(strSQLtprisIns, objConn3)
                                                    objDR6 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)
                                                    objDR6.Close()

                                                Else

                                                    Dim strSQLtprisUpd As String = "UPDATE timepriser SET timeprisalt = " & timeprisalt & ", 6timepris = " & intTimepris & ", 6valuta = " & intValuta & "" _
                                                    & " WHERE jobid = " & jobID & " AND aktid = " & objDR3("aktid") & " AND medarbid = " & objDR2("mid")
                                                    objCmd2 = New OdbcCommand(strSQLtprisUpd, objConn3)
                                                    objDR6 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)
                                                    objDR6.Close()

                                                End If


                                            End While
                                            objDR3.Close()


                                    End Select 'LTO
                                    '** END aktiviteter



                                    '*** OPDATERER Timepris på jobbet ****'
                                    Dim strSQLtpris As String = "UPDATE timepriser SET timeprisalt = " & timeprisalt & ", 6timepris = " & intTimepris & ", 6valuta = " & intValuta & "" _
                                    & " WHERE jobid = " & jobID & " AND aktid = 0 AND medarbid = " & objDR2("mid")
                                    objCmd2 = New OdbcCommand(strSQLtpris, objConn3)
                                    objDR6 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)
                                    objDR6.Close()



                                End While
                                objDR2.Close()
                                objConn3.Close()
                                '** MEdarbejdertype

                        End Select 'LTO



                        '*** Tilføjer til JOBBANKEN på alle aktive medarbejdere

                        Select Case lto
                            Case "dencker", "dencker_test"

                            Case Else


                                Dim strSQLamed As String = "SELECT mid FROM medarbejdere WHERE mansat = 1"
                                objCmd = New OdbcCommand(strSQLamed, objConn2)
                                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                                While objDR2.Read() = True


                                    Dim strSQLtreguseFindes As String = "SELECT medarb FROM timereg_usejob WHERE medarb = " & objDR2("mid") & " AND jobid = " & jobID
                                    objCmd = New OdbcCommand(strSQLtreguseFindes, objConn)
                                    objDR3 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)



                                    timereguseJobFindes = 0
                                    If objDR3.Read() = True Then

                                        timereguseJobFindes = 1

                                    End If
                                    objDR3.Close()


                                    If CInt(timereguseJobFindes) = 0 Then
                                        Dim strSQL3 As String = "INSERT INTO timereg_usejob (medarb, jobid, forvalgt, forvalgt_sortorder, forvalgt_af, forvalgt_dt) VALUES " _
                                        & " (" & objDR2("mid") & ", " & jobID & ", 0, 0, 0, '" & dato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "')"

                                        objCmd = New OdbcCommand(strSQL3, objConn)
                                        objDR3 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                        objDR3.Close()


                                    End If



                                End While
                                objDR2.Close()


                        End Select 'LTO


                    Else 'opret

                        If beskrivelse = "4" And (lto = "dencker" Or lto = "dencker_test") Then
                            beskrivelse = ""


                            '*** Opdaterer beskrivelse
                            Dim jobBesk As String = ""
                            Dim strSQKjobbesk As String = "SELECT jobnavn FROM job_import_temp WHERE jobnr = '" & jobnr & "' AND beskrivelse = 4 AND overfort = 0 ORDER BY id"
                            objCmd = New OdbcCommand(strSQKjobbesk, objConn)
                            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                            While objDR2.Read() = True


                                beskrivelse = beskrivelse + "<br>" + objDR2("jobnavn")

                            End While
                            objDR2.Close()
                        End If


                        Dim strSQLjob As String = ("INSERT INTO job (jobnavn, jobnr, jobknr, jobTpris, jobstatus, jobstartdato," _
                    & " jobslutdato, editor, dato, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, " _
                    & " projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, " _
                    & " fakturerbart, budgettimer, fastpris, kundeok, beskrivelse, " _
                    & " ikkeBudgettimer, tilbudsnr, jobans1, " _
                    & " serviceaft, kundekpers, valuta, rekvnr, " _
                    & " risiko, job_internbesk, " _
                    & " jo_bruttooms, fomr_konto) VALUES " _
                    & "('" & jobnavn & "', " _
                    & "'" & strjobnr & "', " _
                    & "" & kid & ", " _
                    & "0, " _
                    & "" & jobStatus & ", " _
                    & "'" & jobstartdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', " _
                    & "'" & jobslutdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', " _
                    & "'" & editor & "', " _
                    & "'" & dato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', " _
                    & "10, " _
                    & "" & progrpId & ",1,1,1,1,1,1,1,1," _
                    & "1,0,0,0," _
                    & "'" & beskrivelse & "', " _
                    & "0,0, " _
                    & "" & jobans & "," _
                    & "0," & kunderef & ", " _
                    & "1, '" & rekvisitionsnr & "', " _
                    & "100,'" & internnote & "'," _
                    & "" & bruttooms & ", 0)")

                        'return "Webservice SQL: " & strSQLjob


                        objCmd = New OdbcCommand(strSQLjob, objConn)
                        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                        objDR2.Close()



                        '*** Finder jobid ***
                        Dim strSQLlastJobID As String = "SELECT id FROM job WHERE id <> 0 ORDER BY id DESC LIMIT 1"
                        objCmd = New OdbcCommand(strSQLlastJobID, objConn)
                        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                        If objDR2.Read() = True Then

                            lastID = objDR2("id")

                        End If
                        objDR2.Close()


                        '*********** timereg_usejob, så der kan søges fra jobbanken KUN VED OPRET JOB *********************
                        Select Case lto
                            Case "oko"
                            Case "wilke", "dencker", "dencker_test", "tia"

                                Dim strProjektgr1 As Integer = 10
                                Dim strProjektgr2 As Integer = 1
                                Dim strProjektgr3 As Integer = 1
                                Dim strProjektgr4 As Integer = 1
                                Dim strProjektgr5 As Integer = 1
                                Dim strProjektgr6 As Integer = 1
                                Dim strProjektgr7 As Integer = 1
                                Dim strProjektgr8 As Integer = 1
                                Dim strProjektgr9 As Integer = 1
                                Dim strProjektgr10 As Integer = 1



                                Dim strSQLpg As String = "SELECT MedarbejderId FROM progrupperelationer WHERE (" _
                                & " ProjektgruppeId = " & strProjektgr1 & "" _
                                & " OR ProjektgruppeId =" & strProjektgr2 & "" _
                                & " OR ProjektgruppeId =" & strProjektgr3 & "" _
                                & " OR ProjektgruppeId =" & strProjektgr4 & "" _
                                & " OR ProjektgruppeId =" & strProjektgr5 & "" _
                                & " OR ProjektgruppeId =" & strProjektgr6 & "" _
                                & " OR ProjektgruppeId =" & strProjektgr7 & "" _
                                & " OR ProjektgruppeId =" & strProjektgr8 & "" _
                                & " OR ProjektgruppeId =" & strProjektgr9 & "" _
                                & " OR ProjektgruppeId =" & strProjektgr10 & "" _
                                & ") GROUP BY MedarbejderId"

                                'Response.Write "strSQL "& strSQL & "<br><hr>"


                                objCmd = New OdbcCommand(strSQLpg, objConn2)
                                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                                While objDR2.Read() = True





                                    Dim strSQL3 As String = "INSERT INTO timereg_usejob (medarb, jobid, forvalgt, forvalgt_sortorder, forvalgt_af, forvalgt_dt) VALUES " _
                                    & " (" & objDR2("MedarbejderId") & ", " & lastID & ", 0, 0, 0, '" & dato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "')"

                                    objCmd = New OdbcCommand(strSQL3, objConn)
                                    objDR3 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                    objDR3.Close()



                                End While

                                objDR2.Close()


                        End Select ' lto





                        'Dim antalArbPakker As Array
                        'antalArbPakker = strantalArbPakker

                        '** Vælg gruppe pbg. af projetktype
                        Dim agforvalgtStamgrpKri As String = " ag.forvalgt = 1 "
                        Select Case lto
                            Case "oko"


                                Select Case projektKategori
                                    Case 4 'NAER
                                        agforvalgtStamgrpKri = " ag.id = 4 "


                                End Select
                            Case "wilke"

                                Select Case projektKategori
                                    Case 0, 2 'PROJEKT
                                        agforvalgtStamgrpKri = " ag.forvalgt = 1 "

                                    Case 1 'INTERN WILKE TID
                                        agforvalgtStamgrpKri = " ag.id = 6 "
                                    Case 3 ' FRAVÆR
                                        agforvalgtStamgrpKri = " ag.id = 7 "
                                    Case 4 'MARKETING
                                        agforvalgtStamgrpKri = " ag.id = 9 "
                                    Case 5 'Udviklingstid
                                        agforvalgtStamgrpKri = " ag.id = 10 "
                                    Case 6 'Uddanelse
                                        agforvalgtStamgrpKri = " ag.id = 11 "

                                    Case 7 'Salgstid
                                        agforvalgtStamgrpKri = " ag.id = 3 "
                                    Case 8 'ANDET
                                        agforvalgtStamgrpKri = " ag.id = 8 "

                                    Case Else
                                        agforvalgtStamgrpKri = " ag.forvalgt = 1 "
                                End Select

                        End Select


                        Select Case lto
                            Case "dencker", "dencker_test"
                            Case Else

                                '*** OPRETTER STAM AKTIVITETER ***'
                                Call opretStamAkt(lto, lastID, aktnavn, aktstdato, aktsldato, fomr, sort, aktvarenr, antalstk, agforvalgtStamgrpKri, objConn2, objConn, objCmd, objDR2, objDR3, objDR6, objDR4)

                        End Select



                    End If 'Opdater / opret


                    If importtype = "d1" Then


                        '** FOMR REL JOB ***********
                        Dim strSQLaktfomr As String = ("SELECT for_fomr, for_aktid FROM fomr_rel WHERE for_jobid = " & lastID & " AND for_aktid =  0")

                        objCmd = New OdbcCommand(strSQLaktfomr, objConn)
                        objDR4 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                        While objDR4.Read() = True



                            Dim strSQLaktinsfomr As String = ("INSERT INTO fomr_rel (for_fomr, for_aktid, for_jobid, for_faktor) VALUES  (" & fomr & ", 0, " & lastID & ", 100)")

                            objCmd = New OdbcCommand(strSQLaktinsfomr, objConn)
                            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                            objDR2.Close()


                        End While
                        objDR4.Close()


                    End If


                    'objConn2.Close() 20170707
                End If 'opretJobOk kundekundet



            End If 'CInt(errThis) = 0



            Select Case lto
                Case "dencker", "dencker_test"
                    '*** OPRETTER AKTIVITETER FRA EXCEL FIL FOR HVER LINJE // MONITOR ***'

                    If importtype = "d1" Then

                    Else



                        If (lastJobnr <> strjobnr) Then 'Sort = 10
                            '*** Finder jobid ***
                            lastID = 0
                            Dim strSQLlastJobID As String = "SELECT id FROM job WHERE id <> 0 AND jobnr = '" & strjobnr & "' ORDER BY id DESC LIMIT 1"
                            objCmd = New OdbcCommand(strSQLlastJobID, objConn)
                            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                            If objDR2.Read() = True Then

                                lastID = objDR2("id")

                            End If
                            objDR2.Close()


                            '** Opdaterer job start og slutdato med den første produktionsordre / aktivitet start og slut dato
                            Dim strSQLjobUpd As String = ("Update job SET editor = '" & editor & "', dato = '" & dato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', " _
                            & " jobstartdato = '" & aktstdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', jobslutdato = '" & aktsldato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "' WHERE jobnr = '" & strjobnr & "'")

                            objCmd = New OdbcCommand(strSQLjobUpd, objConn)
                            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                            objDR2.Close()


                            '*** Sætter alle aktiviteter til passiv ved første LOOP, hvis en aktivitet er blevet slettet i Monitor
                            If CDbl(lastID) <> 0 Then

                                Dim strSQLAktUpdStatus As String = "UPDATE aktiviteter SET aktstatus = 2" _
                                & " WHERE job = " & lastID & " AND easyreg = 0"
                                objCmd2 = New OdbcCommand(strSQLAktUpdStatus, objConn)
                                objDR6 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)
                                objDR6.Close()

                            End If


                            'If (lastID <> 0) Then
                            ' '** EASY Reg aktiviteten 20171206 ***'
                            'Dim agforvalgtStamgrpKriEa As String = ""
                            'Call opretStamAkt(lto, lastID, "Easyreg", aktstdato, aktsldato, fomr, sort, 901, 0, agforvalgtStamgrpKriEa, objConn2, objConn, objCmd, objDR2, objDR3, objDR6, objDR4)

                            'End If


                        End If

                        If (lastID <> 0) Then


                            '** ALM aktiviteter
                            Dim agforvalgtStamgrpKri As String = ""
                            Call opretStamAkt(lto, lastID, aktnavn, aktstdato, aktsldato, fomr, sort, aktvarenr, antalstk, agforvalgtStamgrpKri, objConn2, objConn, objCmd, objDR2, objDR3, objDR6, objDR4)

                        End If

                    End If

            End Select

            intCountInserted += 1

            lastJobnr = strjobnr


        End While
        objDR.Close()


        'Catch ex As Exception
        '   MsgBox("Can't load Web page" & vbCrLf & ex.Message)
        'End Try


        Dim strSQLjobToverfort As String = "UPDATE job_import_temp SET overfort = 1 WHERE id > 0 AND errid = 0"
        objCmd = New OdbcCommand(strSQLjobToverfort, objConn)
        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
        objDR2.Close()


        Dim childAgeAsInt As Integer
        If Integer.TryParse(jobnr, childAgeAsInt) Then
            ' childAge successfully parsed as Integer

            Dim strSQLjobnrUpd As String = "UPDATE licens SET jobnr = " & jobnr & " WHERE id = 1"
            objCmd = New OdbcCommand(strSQLjobnrUpd, objConn)
            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            objDR2.Close()

        Else
            ' childAge is not an Integer
        End If



        objConn.Close()
        objConn2.Close()






        'Dim errThisTOnoStr As String = errThisTOno.ToString()
        Return "Succes " + intCountInserted.ToString() + " linje(r) indlæst.<br><br>Du kan lukke denne side ned nu. [<a href=""Javascript:window.close();"">LUK</a>]"
        'jobnr: " + err_jobnr + " errid:" + errThisTOnoStr


    End Function



    'Public Function rtErrTxt(ByVal erTxt As String) As String
    ' Dim eTxt As String
    '    eTxt = erTxt

    '   Return eTxt
    'End Function



    Public Function opretStamAkt(ByVal lto As String, ByVal lastID As Integer, ByVal aktnavn As String, ByVal aktstdato As Date, ByVal aktsldato As Date, ByVal fomr As Integer, ByVal sort As Integer, ByVal aktvarenr As String, ByVal antalstk As Double, ByVal agforvalgtStamgrpKri As String, ByVal objConn2 As OdbcConnection, ByVal objConn As OdbcConnection, ByVal objCmd As OdbcCommand, ByVal objDR2 As OdbcDataReader, ByVal objDR3 As OdbcDataReader, ByVal objDR6 As OdbcDataReader, ByVal objDR4 As OdbcDataReader) As String



        '*** Tilføjer stam-aktiviteter ***'
        Select Case lto
            Case "dencker", "dencker_test"
                '*** INDLÆSER AKTIVITETER FRA MONITOR FIL
                '*** AKTIVTETER VIL ALTID KOMME IND UNDER REDIGER JOB STATE Da job bliver oprettet som første linje i filen / eller findes i forvejen

                '*** Finder aktid ***
                Dim aktFindes As Integer = 0
                Dim strSQLaktFindes As String = "SELECT id FROM aktiviteter WHERE job = " & lastID & " AND avarenr = " & aktvarenr & ""
                objCmd = New OdbcCommand(strSQLaktFindes, objConn)
                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                If objDR2.Read() = True Then

                    aktFindes = objDR2("id")

                End If
                objDR2.Close()


                '*** Finder TP på aktivitet **'
                '*** Finder aktid ***
                Dim fasttp As Double = 0
                Dim easyreg As Integer = 0
                Dim fakturerbar As Integer = 0
                Dim strSQLakttp As String = "SELECT fasttp, easyreg, fakturerbar FROM aktiviteter WHERE navn = '" & aktnavn.Replace("'", "") & "' AND job = 0 AND aktfavorit = 72"
                objCmd = New OdbcCommand(strSQLakttp, objConn)
                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                If objDR2.Read() = True Then

                    fasttp = objDR2("fasttp") '.replace(".", ",")
                    easyreg = objDR2("easyreg")
                    fakturerbar = objDR2("fakturerbar")

                End If
                objDR2.Close()

                'fasttp = fasttp.Replace(",", ".")


                strAktFase = strAktFase.Replace(",", "")
                strAktFase = strAktFase.Replace("'", "")

                '**** Findes aktivitet ***'
                If CInt(aktFindes) = 0 Then '** INSERT


                    If CInt(easyreg) = 1 Then
                        aktStatus = 2
                    Else
                        aktStatus = 1
                    End If

                    'If aktnavn = "Easyreg" Then

                    'Else

                    'End If

                    Dim strSQLaktins As String = ("INSERT INTO aktiviteter (navn, job, fakturerbar, " _
                    & "projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7," _
                    & "projektgruppe8, projektgruppe9, projektgruppe10, aktstatus, budgettimer, aktbudget, aktbudgetsum, aktstartdato, aktslutdato, aktkonto, fase, avarenr, fomr, sortorder, antalstk, bgr, " _
                    & " brug_fasttp, fasttp, fasttp_val, easyreg) VALUES " _
                    & " ('" & aktnavn.Replace("'", "") & "', " & lastID & ", " & fakturerbar & "," _
                    & " 10,1,1,1,1,1,1,1,1,1," & aktStatus & ",0,0,0,'" & aktstdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', " _
                    & "'" & aktsldato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', 0, '" & strAktFase & "', '" & aktvarenr & "', " & fomr & ", " & sort & ", " & antalstk & ", 2" _
                    & ", 1, " & fasttp & ", 1, " & easyreg & ")")

                    objCmd = New OdbcCommand(strSQLaktins, objConn)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    objDR2.Close()


                    '*** Finder aktid ***
                    Dim strSQLlastAktID As String = "SELECT id FROM aktiviteter WHERE id <> 0 ORDER BY id DESC"
                    objCmd = New OdbcCommand(strSQLlastAktID, objConn)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                    If objDR2.Read() = True Then

                        lastAktID = objDR2("id")

                    End If
                    objDR2.Close()


                Else '** UPDATE

                    Dim strSQLaktupd As String = ("UPDATE aktiviteter SET navn = '" & aktnavn.Replace("'", "") & "', fase = '" & strAktFase & "', aktstatus = 1, aktstartdato = '" & aktstdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', aktslutdato = '" & aktsldato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', fomr = " & fomr & ", sortorder = " & sort & ", antalstk = " & antalstk & ", fasttp = " & fasttp & " WHERE id = " & aktFindes & "")
                    objCmd = New OdbcCommand(strSQLaktupd, objConn)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    objDR2.Close()


                    lastAktID = aktFindes


                End If




                '** FOMR REL ***********
                Dim strSQLaktfomr As String = ("SELECT for_fomr, for_aktid FROM fomr_rel WHERE for_aktid =  " & lastAktID)

                objCmd = New OdbcCommand(strSQLaktfomr, objConn)
                objDR4 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                While objDR4.Read() = True



                    Dim strSQLaktinsfomr As String = ("INSERT INTO fomr_rel (for_fomr, for_aktid, for_jobid, for_faktor) VALUES  (" & fomr & ", " & lastAktID & ", " & lastID & ", 100)")

                    objCmd = New OdbcCommand(strSQLaktinsfomr, objConn)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    objDR2.Close()


                End While
                objDR4.Close()





            Case Else '"oko", "wilke"



                '*** Indlæser STAMaktiviteter ***'
                Dim strSQLStamAkt As String = "SELECT a.navn AS aktnavn, aktkonto, fase, a.id AS stamaktid FROM akt_gruppe AS ag" _
                & " LEFT JOIN aktiviteter AS a ON (a.aktfavorit = ag.id) WHERE " & agforvalgtStamgrpKri & " AND skabelontype = 0 ORDER BY a.sortorder DESC"
                objCmd = New OdbcCommand(strSQLStamAkt, objConn2)
                objDR3 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                While objDR3.Read() = True






                    Dim strSQLaktins As String = ("INSERT INTO aktiviteter (navn, job, fakturerbar, " _
                    & "projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7," _
                    & "projektgruppe8, projektgruppe9, projektgruppe10, aktstatus, budgettimer, aktbudget, aktbudgetsum, aktstartdato, aktslutdato, aktkonto, fase) VALUES " _
                    & " ('" & objDR3("aktnavn") & "', " & lastID & ", 1," _
                    & " 10,1,1,1,1,1,1,1,1,1,1,0,0,0,'" & jobstartdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', " _
                    & "'" & jobslutdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', " & objDR3("aktkonto") & ", '" & objDR3("fase") & "')")

                    objCmd = New OdbcCommand(strSQLaktins, objConn)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    objDR2.Close()


                    '*** Finder aktid ***
                    Dim strSQLlastAktID As String = "SELECT id FROM aktiviteter WHERE id <> 0 ORDER BY id DESC"
                    objCmd = New OdbcCommand(strSQLlastAktID, objConn)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                    If objDR2.Read() = True Then

                        lastAktID = objDR2("id")

                    End If
                    objDR2.Close()



                    '** FOMR REL ***********
                    Dim strSQLaktfomr As String = ("SELECT for_fomr, for_aktid FROM fomr_rel WHERE for_aktid =  " & objDR3("stamaktid"))

                    objCmd = New OdbcCommand(strSQLaktfomr, objConn)
                    objDR4 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    While objDR4.Read() = True



                        Dim strSQLaktinsfomr As String = ("INSERT INTO fomr_rel (for_fomr, for_aktid, for_jobid, for_faktor) VALUES  (" & objDR4("for_fomr") & ", " & lastAktID & ", " & lastID & ", 100)")

                        objCmd = New OdbcCommand(strSQLaktinsfomr, objConn)
                        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                        objDR2.Close()


                    End While
                    objDR4.Close()



                    '**************************************************************'
                    '*** Hvis timepris ikke findes på job bruges Gen. timepris fra '
                    '*** Fra medarbejdertype, og den oprettes på job              *'
                    '*** BLIVER ALTID HENTET FRA Medarb.type for ØKO              *'
                    '**************************************************************'
                    tprisGen = 0
                    valutaGen = 1
                    kostpris = 0
                    intTimepris = 0
                    intValuta = valutaGen

                    Dim SQLmedtpris As String = "SELECT medarbejdertype, timepris, timepris_a2, timepris_a3, tp0_valuta, kostpris, mnavn, mid FROM medarbejdere " _
                    & " LEFT JOIN medarbejdertyper ON (medarbejdertyper.id = medarbejdertype) " _
                    & " WHERE mid <> 0 AND mansat = 1 AND medarbejdertyper.id = medarbejdertype GROUP BY mid"

                    objCmd = New OdbcCommand(SQLmedtpris, objConn2)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                    While objDR2.Read() = True

                        If objDR2("kostpris") <> 0 Then
                            kostpris = objDR2("kostpris")
                        Else
                            kostpris = 0
                        End If


                        Select Case lto
                            Case "oko"


                                Select Case Left(Trim(beskrivelse), 8)
                                    Case "NaturErh"

                                        If objDR2("timepris_a2") <> 0 Then

                                            tprisGen = objDR2("timepris_a2")
                                        Else
                                            tprisGen = 0
                                        End If

                                        timeprisalt = 2

                                    Case "Direktor"

                                        If objDR2("timepris_a3") <> 0 Then

                                            tprisGen = objDR2("timepris_a3")
                                        Else
                                            tprisGen = 0
                                        End If

                                        timeprisalt = 3

                                    Case Else

                                        If objDR2("timepris") <> 0 Then

                                            tprisGen = objDR2("timepris")
                                        Else
                                            tprisGen = 0
                                        End If

                                        timeprisalt = 0

                                End Select


                            Case Else

                                If objDR2("timepris") <> 0 Then

                                    tprisGen = objDR2("timepris")
                                Else
                                    tprisGen = 0
                                End If

                                timeprisalt = 0




                        End Select




                        valutaGen = objDR2("tp0_valuta")


                        '**** Indlæser timepris på aktiviteter ***'
                        intTimepris = tprisGen
                        intTimepris = Replace(intTimepris, ",", ".")

                        intValuta = valutaGen

                        Dim strSQLtpris As String = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) " _
                        & " VALUES (" & lastID & ", " & lastAktID & ", " & objDR2("mid") & ", " & timeprisalt & ", " & intTimepris & ", " & intValuta & ")"

                        objCmd = New OdbcCommand(strSQLtpris, objConn)
                        objDR6 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)



                    End While

                    objDR2.Close()
                    '** Slut timepris **







                End While
                objDR3.Close()


        End Select 'lto STAMAKTIVITETER

    End Function 'Opret aktvitieter


End Class