<%@ WebService language="VB" class="oz_importmed_na" %>
Imports System
Imports System.Web.Services
Imports System.Data
Imports System.Data.Odbc


'***************************************************************************************
'*** DENNE SERVICE BENYTTES AF 
'*** TIA
'***************************************************************************************


Public Class oz_importmed_na


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
    Public dato, editorDate As Date

    Public origin As String

    Public importtype As String

    Public j, jhigh, jlow, LastID As Integer

    '*** TIA VAR
    Public lto, init, mnavn, email, expvendorno, costcenter, linemanager, countrycode, weblang, mednavn As String
    Public normtid As String
    Public ansatdato As Date
    Public opsagtdato As Date
    Public mansat As String

    Public loginnavn As String
    Public loginpw As String

    Public medarbejdertype As String
    Public errId As String = "0"

    Public sprog As Integer = 1
    Public LastMedID As String = "0"





    <WebMethod()> Public Function addmed(ByVal ds As DataSet) As String

        'Return "Webservice Msg dt: "  


        'On Error Resume Next



        'Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_intranet;User=outzource;Password=SKba200473;"

        Dim objConn As OdbcConnection
        Dim objConn2 As OdbcConnection
        'Dim objConn3 As OdbcConnection
        Dim objCmd As OdbcCommand
        Dim objCmd2 As OdbcCommand
        'Dim objCmd3 As OdbcCommand
        'Dim objDataSet As New DataSet
        Dim objDR As OdbcDataReader
        Dim objDR2 As OdbcDataReader
        'Dim objDR3 As OdbcDataReader
        'Dim objDR3 As OdbcDataReader

        Dim strConn As String
        Dim t As Double = 0
        Dim antalRecords As Integer = 0


        'Try

        lto = "tia"

        'strConn = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_epi_catitest;User=outzource;Password=SKba200473;"
        'strConn = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_oko;User=to_outzource2;Password=SKba200473;"
        'strConn = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=root;pwd=;database=timeout_oko; OPTION=32"
        'strConn = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_oko;User=to_outzource2;Password=SKba200473;"
        strConn = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" & lto & ";User=to_outzource2;Password=SKba200473;"


        '** Åbner Connection ***'
        objConn = New OdbcConnection(strConn)
        objConn.Open()

        objConn2 = New OdbcConnection(strConn)
        objConn2.Open()




        jlow = 0
        jhigh = 10000









        '*** HENTER med fra MED_IMPORT_TEMP '****

        Dim strSQLmedins As String = "SELECT id, dato, origin, Init, "
        strSQLmedins += "mnavn, "
        strSQLmedins += "email, "
        strSQLmedins += "normtid, "
        strSQLmedins += "ansatdato, "
        strSQLmedins += "opsagtdato, "
        strSQLmedins += "mansat, "
        strSQLmedins += "expvendorno, "
        strSQLmedins += "costcenter, "
        strSQLmedins += "linemanager, "
        strSQLmedins += "countrycode, "
        strSQLmedins += "weblang, lto, editor, overfort FROM med_import_temp_ds WHERE id > 0 AND overfort = 0 AND errid = 0 ORDER BY id LIMIT " & jlow & ", " & jhigh & ""
        objCmd = New OdbcCommand(strSQLmedins, objConn)
        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

        Dim a As Integer = 0

        While objDR.Read() = True



            'aktsumTilSum = 0
            'errThis = 0

            id = objDR("id")
            dato = objDR("dato")
            origin = objDR("origin")
            editor = objDR("editor")


            init = objDR("init")
            email = objDR("email")

            mednavn = objDR("mnavn")
            mednavn = EncodeUTF8(mednavn)
            mednavn = DecodeUTF8(mednavn)

            mednavn = Replace(mednavn, "'", "")

            lto = objDR("lto")
            normtid = objDR("normtid")
            normtid = Replace(normtid, ".", ",")

            Select Case normtid
                Case "37,5", "37.5"
                    medarbejdertype = 1
                Case "25"
                    medarbejdertype = 4
                Case "30"
                    medarbejdertype = 9
                Case "40"
                    medarbejdertype = 8
                Case "24"
                    medarbejdertype = 10
                Case Else
                    medarbejdertype = 2
            End Select

            'konsulent objDR("medarbejdertype")
            ansatdato = objDR("ansatdato")
            If CDate(ansatdato) <= "01-01-2002" Then
                ansatdato = "2002-01-01"
            End If

            opsagtdato = objDR("opsagtdato")

            If CDate(opsagtdato) <= "01-01-2002" Then
                opsagtdato = "2044-01-01"
            End If

            mansat = objDR("mansat")
            If mansat = "0" Then
                mansat = "2"
            End If

            'costcenter
            costcenter = objDR("costcenter")
            If costcenter = "" Then
                costcenter = "0"
            End If

            linemanager = objDR("linemanager")
            'countrycode
            'weblang

            '*** Opretter Medarb ****
            loginnavn = init
            loginpw = init + "1234"
            'ansatdato, opsagtdato

            '*** Findes med INIT ***
            Dim medidFindes As Integer = 0

            Dim strSQLlastJobID As String = "SELECT mid FROM medarbejdere WHERE init = '" & init & "' LIMIT 1"
            objCmd = New OdbcCommand(strSQLlastJobID, objConn)
            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

            If objDR2.Read() = True Then

                medidFindes = 1

            End If
            objDR2.Close()


            If LCase(init) = "sni" And lto = "tia" Then
                errId = 10
            End If

            Select Case weblang
                Case "DK"
                    sprog = 1
                Case Else
                    sprog = 2
            End Select


            If errId = 0 Then

                If CInt(medidFindes) = 0 Then

                    Dim strSQLaktins As String = "INSERT INTO medarbejdere (editor, dato, mnavn, mnr, init, login, pw, medarbejdertype, brugergruppe, mansat, email, ansatdato, opsagtdato, sprog, tsacrm) VALUES "
                    strSQLaktins += " ('TO_import-nav','" & Now.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', '" & mednavn & "', '" & id & "',"
                    strSQLaktins += "'" & init & "', '" & init & "', MD5('" & loginpw & "'), " & medarbejdertype & ", 6, " & mansat & ", '" & email & "', "
                    strSQLaktins += "'" & ansatdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', '" & opsagtdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', " & sprog & ", 9)"

                    objCmd2 = New OdbcCommand(strSQLaktins, objConn2)
                    objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)

                Else

                    'brugergruppe = 6,
                    Dim strSQLmedUpd As String = "UPDATE medarbejdere SET editor = 'TO_import-nav-upd', dato = '" & Now.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "',"
                    strSQLmedUpd += "mnavn = '" & mednavn & "', medarbejdertype = " & medarbejdertype & ", mansat = " & mansat & ", email = '" & email & "',"
                    strSQLmedUpd += "ansatdato = '" & ansatdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', opsagtdato = '" & opsagtdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', sprog = " & sprog & ""
                    strSQLmedUpd += " WHERE init = '" & init & "'"

                    objCmd2 = New OdbcCommand(strSQLmedUpd, objConn2)
                    objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)

                End If


                'Return strSQLaktins




                '*** Finder jobid ***
                Dim strSQLlastMedID As String = "SELECT mid FROM medarbejdere WHERE init = '" & init & "'"
                objCmd = New OdbcCommand(strSQLlastMedID, objConn)
                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                If objDR2.Read() = True Then

                    LastMedID = objDR2("mid")

                End If
                objDR2.Close()


                If CInt(medidFindes) = 0 Then

                    '*** Opretter Medarb i projektgrupper ****
                    Dim strSQLpgrel As String = ("INSERT INTO progrupperelationer (projektgruppeid, medarbejderid) VALUES " _
                    & " (10, " & LastMedID & ")")

                    'Return strSQLaktins

                    objCmd2 = New OdbcCommand(strSQLpgrel, objConn2)
                    objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)

                End If


                If costcenter <> "0" Then
                    '*** Opretter Medarb i COSTCENTER projektgrupper ****


                    Dim mapping As String = 0
                    Select Case costcenter.ToString

                        Case "100", "110", "120", "130", "140", "150"
                            mapping = 11
                        Case "300"
                            mapping = 12
                        Case "320"
                            mapping = 13
                        Case "330"
                            mapping = 14
                        Case "400"
                            mapping = 15
                        Case "500"
                            mapping = 60
                        Case "600", "610"
                            mapping = 16
                        Case "700"
                            mapping = 17
                    End Select


                    '*** sletter eksisterende projgrprel (hvis der er skiftet grupper) ***
                    Dim strSQLpgldel As String = "DELETE FROM progrupperelationer WHERE medarbejderid =  " & LastMedID & " AND "
                    strSQLpgldel += "(projektgruppeid = 11 Or projektgruppeid = 12 Or projektgruppeid = 13 Or projektgruppeid = 14 Or projektgruppeid = 15 Or projektgruppeid = 16 Or projektgruppeid = 17)"

                    objCmd2 = New OdbcCommand(strSQLpgldel, objConn2)
                    objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)

                    '** Indsætter **'
                    Dim strSQLpgrelc As String = ("INSERT INTO progrupperelationer (projektgruppeid, medarbejderid) VALUES " _
                        & " (" & mapping & ", " & LastMedID & ")")


                    objCmd2 = New OdbcCommand(strSQLpgrelc, objConn2)
                    objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)

                End If



                '*** Linemanager ********** (KRÆVER Linemaneger FINDES og er indlæst i TO)
                If Trim(linemanager) <> "" Then


                    '** Indsætter linemangger som TEAMleder
                    Dim LineMLastId As Integer = 0
                    Dim strSQLLMId As String = "SELECT mid FROM medarbejdere WHERE init = '" & linemanager & "' LIMIT 1"
                    objCmd = New OdbcCommand(strSQLLMId, objConn)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                    If objDR2.Read() = True Then

                        LineMLastId = objDR2("mid")

                    End If
                    objDR2.Close()

                    '** medidFindes Linemanager **
                    If CInt(LineMLastId) <> 0 Then

                        '*** Findes med INIT ***
                        Dim linemanagerFindes As Integer = 0
                        Dim HrgrpId As Integer = 0

                        Dim strSQLlinemanagerFindes As String = "SELECT id FROM projektgrupper WHERE navn = 'HR - " & linemanager & "' AND orgvir = 2 LIMIT 1"
                        objCmd = New OdbcCommand(strSQLlinemanagerFindes, objConn)
                        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                        If objDR2.Read() = True Then

                            linemanagerFindes = 1
                            HrgrpId = objDR2("id")

                        End If
                        objDR2.Close()

                        '*** Opretter HR Projektgruppe ***
                        If CInt(linemanagerFindes) = 0 Then

                            '** Indsætter **'
                            Dim strSQLoprHRgrp As String = ("INSERT INTO projektgrupper (navn, editor, dato, orgvir) VALUES " _
                            & " ('HR - " & linemanager & "', 'NAV import', '" & Now.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', 2)")


                            objCmd2 = New OdbcCommand(strSQLoprHRgrp, objConn2)
                            objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)

                            Dim HrgrpIdLastId As Integer = 0
                            Dim strSQLlinemanagerLastId As String = "SELECT id FROM projektgrupper WHERE navn = 'HR - " & linemanager & "' AND orgvir = 2 LIMIT 1"
                            objCmd = New OdbcCommand(strSQLlinemanagerLastId, objConn)
                            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

                            If objDR2.Read() = True Then


                                HrgrpIdLastId = objDR2("id")

                            End If
                            objDR2.Close()




                            If LineMLastId <> 0 Then

                                Dim strSQLpgrtl As String = ("INSERT INTO progrupperelationer (projektgruppeid, medarbejderid, teamleder) VALUES " _
                                & " (" & HrgrpIdLastId & ", " & LineMLastId & ", 1)")


                                objCmd2 = New OdbcCommand(strSQLpgrtl, objConn2)
                                objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)

                            End If



                        End If





                        '*** sletter eksisterende projgrprel til LM ***
                        Dim strSQLpgldelLM As String = "DELETE FROM progrupperelationer WHERE medarbejderid =  " & LastMedID & " AND projektgruppeid <> 10 AND "
                        strSQLpgldelLM += "(projektgruppeid <> 11 AND projektgruppeid <> 12 AND projektgruppeid <> 13 AND projektgruppeid <> 14 AND projektgruppeid <> 15 AND projektgruppeid <> 16 AND projektgruppeid <> 17)"


                        objCmd2 = New OdbcCommand(strSQLpgldelLM, objConn2)
                        objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)

                        '** Indsætter **'
                        Dim strSQLpgrelc As String = ("INSERT INTO progrupperelationer (projektgruppeid, medarbejderid) VALUES " _
                            & " (" & HrgrpId & ", " & LastMedID & ")")


                        objCmd2 = New OdbcCommand(strSQLpgrelc, objConn2)
                        objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)


                    End If 'If CInt(LineMLastId) <> 0 Then

                End If ' linemanager <> ""






            End If 'errID
            '**** Opdater overført ************

            Dim strSQLjobToverfort As String = "UPDATE med_import_temp_ds SET overfort = 1 WHERE overfort = 0 AND init = '" & init & "'"
            objCmd2 = New OdbcCommand(strSQLjobToverfort, objConn2)
            objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)



            a = a + 1


        End While
        objDR.Close()


        objConn.Close()
        objConn2.Close()








        '**** Opdater overført ************

        'objConn2 = New OdbcConnection(strConn)
        'objConn2.Open()
        'Dim strSQLjobToverfort As String = "UPDATE akt_import_temp SET overfort = 1 WHERE id <= " & id & ""
        'objCmd = New OdbcCommand(strSQLjobToverfort, objConn2)
        'objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
        'objDR2.Close()
        'objConn2.Close()

        '**** Delete overført ************

        'objConn2 = New OdbcConnection(strConn)
        'objConn2.Open()
        'Dim strSQLjobToverfort As String = "DELETE FROM akt_import_temp WHERE id > 0 AND errid = 0"
        'objCmd = New OdbcCommand(strSQLjobToverfort, objConn2)
        'objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
        'objDR2.Close()
        'objConn2.Close()




        'Next

        Return "<br>Antal medarbejdere opdateret: " & a & "<br>"

    End Function



    'Public Function rtErrTxt(ByVal erTxt As String) As String
    ' Dim eTxt As String
    '    eTxt = erTxt

    '   Return eTxt
    'End Function


End Class






