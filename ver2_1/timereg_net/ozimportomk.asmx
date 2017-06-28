



<%@ WebService language="VB" class="oz_importomk" %>
Imports System
Imports System.Web.Services
Imports System.Data
Imports System.Data.Odbc


'***************************************************************************************
'*** DENNE SERVICE BENYTTES AF 
'*** OKO
'*** Standard Excel uplaod
'***************************************************************************************


Public Class oz_importomk


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

    Public origin As String
    Public mednavn As String
    Public aktnr As String = "0"
    Public jobans As Integer = 0
    Public lto As String
    Public beskrivelse As String
    Public errThis As Integer = 0

    Public jobNr As String = "0"



    '*** NAV import OMK
    Public strNavn As String = ""
    Public dblKobsPris As Double = 0
    Public dblSalgsPris As Double = 0
    Public belob As Double = 0
    Public usemrn As String = "0"
    Public useMid As String = "0"
    Public bogforingsdato As Date
    Public aftid As String = "0"
    Public intkode As Integer = 0
    Public intValuta As Integer = 1
    Public bilagsnr As String = "0"
    Public dblKurs As Double = 100
    Public personlig As Integer = 0
    Public extsysid As String = "9900"
    '**** SLUT **


    Public jobId As Integer = 0
    Public kontoId As Integer = 0


    Public varLobenr As String = "0"
    Public isJobWrt As String = "#0#"

    Public findesOmkostning As Integer = 0
    Public findesAkt As Integer = 0
    'Public findesBudgetpost As Integer = 0

    Public aktsumGrandBel As String = "0"

    Public aktsumGrandTotalJob(5000) As Double
    Public matsumGrandTotalJob(5000) As Double

    Public importtype As String

    Public j, jhigh, jlow, a As Integer
    Public lastID As Double



    <WebMethod()> Public Function addomk(ByVal ds As DataSet) As String

        'Return "Webservice Msg dt: "  


        'On Error Resume Next


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


                lto = ds.Tables("tb_to_var").Rows(0).Item("lto")
                importtype = ds.Tables("tb_to_var").Rows(0).Item("importtype")

            Next

        Next


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

        'Try







        jlow = 0
        jhigh = 10000


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


        '*** HENTER AKT fra MED_IMPORT_TEMP '****
        Dim strSQLakts As String = "SELECT pwc.dato, pwc.konto, pwc.postext, pwc.belob, pwc.jobnr, pwc.init, pwc.extsysid, pwc.id, "
        strSQLakts += " j.id AS jid, serviceaft, m.mid, pwc.valutakode, pwc.bogforingsdato FROM mat_import_temp AS pwc "
        strSQLakts += " LEFT JOIN job AS j ON (j.jobnr = pwc.jobnr) "
        strSQLakts += " Left JOIN medarbejdere As m On (m.init = pwc.init) WHERE pwc.postext <> '' AND pwc.belob <> 0 ORDER BY pwc.id"
        objCmd = New OdbcCommand(strSQLakts, objConn)
        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

        Dim a As Integer = 0

        While objDR.Read() = True

            id = objDR("id")

            If IsDBNull(objDR("postext")) <> True Then
                strNavn = Replace(objDR("postext"), "'", "")
            Else
                strNavn = "-"
            End If

            strNavn = EncodeUTF8(strNavn)
            strNavn = DecodeUTF8(strNavn)

            belob = Replace(objDR("belob"), ".", "")
            belob = Replace(belob, ",", ".")

            'belob = objDR("belob")

            dblKobsPris = belob
            dblSalgsPris = belob

            If IsDBNull(objDR("jid")) <> True Then
                jobId = objDR("jid")
            Else
                jobId = 3 '** Intern tid
            End If

            editor = "PWC Import"

            dato = Year(Now) & "/" & Month(Now) & "/" & Day(Now)

            If IsDBNull(objDR("mid")) = False Then


                If Len(Trim(objDR("mid"))) > 4 And (Left(LCase(objDR("mid")), 3) <> "itw" And Left(LCase(objDR("mid")), 4) <> "intw") Then
                    useMid = 620
                Else
                    useMid = objDR("mid")
                End If
            Else
                useMid = 621 'materiale forbrug ekstern / ikke fundet
            End If

            usemrn = useMid


            bogforingsdato = objDR("bogforingsdato")

            If IsDBNull(objDR("serviceaft")) <> True Then
                aftid = objDR("serviceaft")
            Else
                aftid = 0
            End If

            If useMid <> 620 Then
                intkode = 0
            Else
                intkode = 2
            End If


            '*** Finder Valuta ***
            Dim strSQLValuta As String = "SELECT id, kurs FROM valutaer WHERE valutakode = '" & objDR("valutakode") & "' LIMIT 1"
            objCmd = New OdbcCommand(strSQLValuta, objConn)
            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

            If objDR2.Read() = True Then

                intValuta = objDR2("id")
                dblKurs = objDR2("kurs")

            End If
            objDR2.Close()




            bilagsnr = "9900" & objDR("extsysid")

            personlig = 0

            extsysid = objDR("extsysid")

            extsysid = "9900" & extsysid

            Dim strSQLMatIndlas As String = "INSERT INTO materiale_forbrug "
            strSQLMatIndlas += " (matid, matantal, matnavn, matvarenr, matkobspris, matsalgspris, matenhed, jobid, "
            strSQLMatIndlas += " editor, dato, usrid, matgrp, forbrugsdato, serviceaft, valuta, intkode, bilagsnr, kurs, personlig, extsysid) VALUES "
            strSQLMatIndlas += " (0, 1, '" & strNavn & "', '0', "
            strSQLMatIndlas += " " & dblKobsPris & ", " & dblSalgsPris & ", 'Stk.', " & jobId & ", "
            strSQLMatIndlas += " '" & editor & "', '" & dato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', " & usemrn & ", "
            strSQLMatIndlas += " 0, '" & bogforingsdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', " & aftid & ", " & intValuta & ", "
            strSQLMatIndlas += " " & intkode & ", '" & bilagsnr & "', " & dblKurs & ", " & personlig & ", " & extsysid & ")"
            objCmd2 = New OdbcCommand(strSQLMatIndlas, objConn2)
            objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)



            'oConn.execute(strSQL)

            '**** Opdater overført ************

            Dim strSQLjobToverfort As String = "UPDATE mat_import_temp SET overfort = 1 WHERE id = " & id & ""
            objCmd2 = New OdbcCommand(strSQLjobToverfort, objConn2)
            objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)
            'objDR2.Close()
            'objConn2.Close()



            a = a + 1


        End While
        objDR.Close()
        objConn.Close()

        objConn2.Close()







        'Next

        Return "<br>Last Webservice loop: " & a & "<br>"

    End Function



    'Public Function rtErrTxt(ByVal erTxt As String) As String
    ' Dim eTxt As String
    '    eTxt = erTxt

    '   Return eTxt
    'End Function


End Class






