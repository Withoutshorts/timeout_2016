



<%@ WebService language="VB" class="oz_importmed" %>
Imports System
Imports System.Web.Services
Imports System.Data
Imports System.Data.Odbc


'***************************************************************************************
'*** DENNE SERVICE BENYTTES AF 
'*** TIA
'***************************************************************************************


Public Class oz_importmed


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
    Public dato, editorDate As Date

    Public origin As String
    Public mednavn As String
    Public aktnr As String = "0"
    Public jobans As Integer = 0
    Public lto As String
    Public beskrivelse As String
    Public errThis As Integer = 0

    Public jobNr As String = "0"

    Public akttimer As String = "0"
    Public akttpris As String = "0"
    Public aktsum As String = "0"
    Public aktsumTilSum As String = "0"

    Public jobId As Integer = 0
    Public kontoId As Integer = 0

    Public jobstartdato As Date
    Public jobslutdato As Date

    Public aktkonto As String = "0"
    Public akttype As String = ""
    Public varLobenr As String = "0"
    Public isJobWrt As String = "#0#"

    Public findesOmkostning As Integer = 0
    Public findesAkt As Integer = 0
    'Public findesBudgetpost As Integer = 0

    Public aktsumGrandBel As String = "0"

    Public aktsumGrandTotalJob(5000) As Double
    Public matsumGrandTotalJob(5000) As Double

    Public importtype As String

    Public j, jhigh, jlow As Integer
    Public lastID As Double

    '*** TIA VAR
    Public init, mnavn, email, expvendorno, costcenter, linemanager, countrycode, weblang As String
    Public normtid As Double
    Public ansatdato As Date
    Public opsagtdato As Date
    Public mansat As Integer

    Public loginnavn As String
    Public loginpw As String

    Public medarbejdertype As String




    <WebMethod()> Public Function addmed(ByVal ds As DataSet) As String

        'Return "Webservice Msg dt: "  


        'On Error Resume Next





        'Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_intranet;User=outzource;Password=SKba200473;"

        Dim objConn As OdbcConnection
        Dim objConn2 As OdbcConnection
        'Dim objConn3 As OdbcConnection
        Dim objCmd As OdbcCommand
        Dim objCmd2 As OdbcCommand
        Dim objCmd3 As OdbcCommand
        'Dim objDataSet As New DataSet
        Dim objDR As OdbcDataReader
        Dim objDR2 As OdbcDataReader
        Dim objDR3 As OdbcDataReader
        'Dim objDR3 As OdbcDataReader
        'Dim dt As DataTable
        'Dim dr As DataRow

        Dim strConn As String
        Dim t As Double = 0
        Dim antalRecords As Integer = 0


        'Try







        jlow = 0
        jhigh = 10000





        'Dim con As New SqlConnection(Application("MyDatabaseConnectionString"))
        'Dim daCust As New SqlDataAdapter("Select * From tblCustomer", con)
        'Dim cbCust As New SqlCommandBuilder(daCust)
        'daCust.Update(ds, "Cust")
        'Return ds

        lto = "demo"

        Dim dt As DataTable
        Dim dr As DataRow

        For Each dt In ds.Tables
            For Each dr In dt.Rows

                '** System variables
                lto = ds.Tables("tb_to_var").Rows(0).Item("lto")
                importtype = ds.Tables("tb_to_var").Rows(0).Item("importtype")
                editorDate = Year(Now) & "/" & Month(Now) & "/" & Day(Now)

                If (antalRecords = 0) Then


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


                End If



                '** TIA table 156 variables
                If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("init")) = False Then
                    init = ds.Tables("tb_to_var").Rows(0).Item("init")
                Else
                    init = ""
                End If

                If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("mname")) = False Then
                    mnavn = ds.Tables("tb_to_var").Rows(0).Item("mname")
                Else
                    mnavn = ""
                End If

                If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("email")) = False Then
                    email = ds.Tables("tb_to_var").Rows(0).Item("email")
                Else
                    email = ""
                End If

                If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("normtid")) = False Then
                    normtid = ds.Tables("tb_to_var").Rows(0).Item("normtid")
                Else
                    normtid = 0
                End If

                If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("empdate")) = False Then
                    ansatdato = ds.Tables("tb_to_var").Rows(0).Item("empdate")
                Else
                    ansatdato = "2002-01-01"
                End If

                If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("terdate")) = False Then
                    opsagtdato = ds.Tables("tb_to_var").Rows(0).Item("terdate")
                Else
                    opsagtdato = "2002-01-01"
                End If

                If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("blocked")) = False Then
                    mansat = ds.Tables("tb_to_var").Rows(0).Item("blocked")
                Else
                    mansat = 0
                End If

                If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("expvendorno")) = False Then
                    expvendorno = ds.Tables("tb_to_var").Rows(0).Item("expvendorno")
                Else
                    expvendorno = ""
                End If

                If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("costcenter")) = False Then
                    costcenter = ds.Tables("tb_to_var").Rows(0).Item("costcenter")
                Else
                    costcenter = ""
                End If

                If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("linemanager")) = False Then
                    linemanager = ds.Tables("tb_to_var").Rows(0).Item("linemanager")
                Else
                    linemanager = ""
                End If

                If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("countrycode")) = False Then
                    countrycode = ds.Tables("tb_to_var").Rows(0).Item("countrycode")
                Else
                    countrycode = ""
                End If

                If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("weblang")) = False Then
                    weblang = ds.Tables("tb_to_var").Rows(0).Item("weblang")
                Else
                    weblang = ""
                End If




                '*** Opretter Medarb ****

                Dim strSQLmedinsTemp As String = "INSERT INTO med_import_temp (dato, origin, Init, "
                strSQLmedinsTemp += "mnavn, "
                strSQLmedinsTemp += "email, "
                strSQLmedinsTemp += "normtid, "
                strSQLmedinsTemp += "ansatdato, "
                strSQLmedinsTemp += "opsagtdato, "
                strSQLmedinsTemp += "mansat, "
                strSQLmedinsTemp += "expvendorno, "
                strSQLmedinsTemp += "costcenter, "
                strSQLmedinsTemp += "linemanager, "
                strSQLmedinsTemp += "countrycode, "
                strSQLmedinsTemp += "weblang,  lto, editor, overfort) "
                strSQLmedinsTemp += " VALUES ('" + Year(Now) & "/" & Month(Now) & "/" & Day(Now) + "',914,'" + init + "','" + mnavn + "','" + email + "',"
                strSQLmedinsTemp += "" + normtid + ",'" & ansatdato  &"','" & opsagtdato & "'," + mansat + ", '" + expvendorno + "', '" + costcenter + "', '" + linemanager + "','" + countrycode + "', '" + weblang + "',"
                strSQLmedinsTemp += "'','tia','Timeout - ImportMedService ',0)"

                'Return strSQLaktins

                objCmd2 = New OdbcCommand(strSQLmedinsTemp, objConn2)
                objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)

                antalRecords = antalRecords + 1
            Next



        Next


        '*** SLUT indlæsning afdataset





        '*** HENTER med fra MED_IMPORT_TEMP '****

        Dim strSQLmedins As String = "SELECT id, origin, Init, "
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
        strSQLmedins += "weblang, lto, editor, overfort FROM med_import_temp WHERE id > 0 AND overfort = 0 AND errid = 0 ORDER BY id LIMIT " & jlow & ", " & jhigh & ""
        objCmd = New OdbcCommand(strSQLmedins, objConn)
        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

        Dim a As Integer = 0

        While objDR.Read() = True



            aktsumTilSum = 0
            errThis = 0

            id = objDR("id")
            dato = objDR("dato")
            origin = objDR("origin")
            editor = objDR("editor")


            init = objDR("init")
            email = objDR("email")

            mednavn = objDR("mnavn")
            mednavn = EncodeUTF8(mednavn)
            mednavn = DecodeUTF8(mednavn)

            lto = objDR("lto")
            medarbejdertype = 1 ' konsulent objDR("medarbejdertype")
            ansatdato = objDR("ansatdato")
            opsagtdato = objDR("opsagtdato")

            'costcenter
            'linemanager
            'countrycode
            'weblang

            '*** Opretter Medarb ****
            loginnavn = init
            loginpw = init + "1234"
            Dim strSQLaktins As String = "INSERT INTO medarbejdere (editor, dato, mnavn, mnr, init, medarbejdertype, brugergruppe, mansat, email, ansatdato, opsagtdato) VALUES "
            strSQLaktins += " ('TO_import-nav','" & Now.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', '" & mednavn & "', '" & id & "',"
            strSQLaktins += "'" & init & "','" & medarbejdertype & "', 6, 1, '" & email & "', '" & ansatdato & "', '" & opsagtdato & "')"

            'Return strSQLaktins

            objCmd2 = New OdbcCommand(strSQLaktins, objConn2)
            objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)


            '*** Finder jobid ***
            Dim strSQLlastJobID As String = "SELECT mid FROM medarbejdere WHERE mid <> 0 ORDER BY mid DESC LIMIT 1"
            objCmd = New OdbcCommand(strSQLlastJobID, objConn)
            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

            If objDR2.Read() = True Then

                lastID = objDR2("mid")

            End If
            objDR2.Close()


            '*** Opretter Medarb ****
            Dim strSQLpgrel As String = ("INSERT INTO progrupperelationer (projektgruppeid, medarbejderid) VALUES " _
            & " (10, " & lastID & ")")

            'Return strSQLaktins

            objCmd2 = New OdbcCommand(strSQLpgrel, objConn2)
            objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)
            'objDR2.Close()
            'objConn2.Close(



            '**** Opdater overført ************

            Dim strSQLjobToverfort As String = "UPDATE med_import_temp SET overfort = 1 WHERE id = " & id & ""
            objCmd2 = New OdbcCommand(strSQLjobToverfort, objConn2)
            objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)
            'objDR2.Close()
            'objConn2.Close()



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

        Return "<br>Last Webservice loop: " & j & "<br>"

    End Function



    'Public Function rtErrTxt(ByVal erTxt As String) As String
    ' Dim eTxt As String
    '    eTxt = erTxt

    '   Return eTxt
    'End Function


End Class






