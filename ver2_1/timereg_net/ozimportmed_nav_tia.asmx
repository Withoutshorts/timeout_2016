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
            'mednavn = EncodeUTF8(mednavn)
            'mednavn = DecodeUTF8(mednavn)

            lto = objDR("lto")
            medarbejdertype = 1 ' konsulent objDR("medarbejdertype")
            ansatdato = objDR("ansatdato")
            opsagtdato = objDR("opsagtdato")

            mansat = objDR("mansat")

            'costcenter
            'linemanager
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


            If CInt(medidFindes) = 0 Then

                Dim strSQLaktins As String = "INSERT INTO medarbejdere (editor, dato, mnavn, mnr, init, login, pw, medarbejdertype, brugergruppe, mansat, email, ansatdato, opsagtdato) VALUES "
                strSQLaktins += " ('TO_import-nav','" & Now.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', '" & mednavn & "', '" & id & "',"
                strSQLaktins += "'" & init & "', '" & init & "', MD5('" & loginpw & "'), " & medarbejdertype & ", 6, " & mansat & ", '" & email & "', "
                strSQLaktins += "'" & ansatdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', '" & opsagtdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "')"

                objCmd2 = New OdbcCommand(strSQLaktins, objConn2)
                objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)

            Else

                Dim strSQLmedUpd As String = "UPDATE medarbejdere SET editor = 'TO_import-nav-upd', dato = '" & Now.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "',"
                strSQLmedUpd += "mnavn = '" & mednavn & "', medarbejdertype = " & medarbejdertype & ", brugergruppe = 6, mansat = " & mansat & ", email = '" & email & "',"
                strSQLmedUpd += "ansatdato = '" & ansatdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', opsagtdato = '" & opsagtdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "'"
                strSQLmedUpd += " WHERE init = '" & init & "'"

                objCmd2 = New OdbcCommand(strSQLmedUpd, objConn2)
                objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)

            End If


            'Return strSQLaktins




            '*** Finder jobid ***
            'Dim strSQLlastJobID As String = "SELECT mid FROM medarbejdere WHERE mid <> 0 ORDER BY mid DESC LIMIT 1"
            'objCmd = New OdbcCommand(strSQLlastJobID, objConn)
            'objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

            'If objDR2.Read() = True Then

            'LastID = objDR2("mid")

            'End If
            'objDR2.Close()


            '*** Opretter Medarb i projektgrupper ****
            Dim strSQLpgrel As String = ("INSERT INTO progrupperelationer (projektgruppeid, medarbejderid) VALUES " _
            & " (10, " & LastID & ")")

            'Return strSQLaktins

            objCmd2 = New OdbcCommand(strSQLpgrel, objConn2)
            objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)


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

        Return "<br>Last Webservice loop: " & j & "<br>"

    End Function



    'Public Function rtErrTxt(ByVal erTxt As String) As String
    ' Dim eTxt As String
    '    eTxt = erTxt

    '   Return eTxt
    'End Function


End Class






