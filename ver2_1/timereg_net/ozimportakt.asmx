



<%@ WebService language="VB" class="oz_importakt" %>
Imports System
Imports System.Web.Services
Imports System.Data
Imports System.Data.Odbc


'***************************************************************************************
'*** DENNE SERVICE BENYTTES AF OKO
'***************************************************************************************


Public Class oz_importakt

    
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
    Public aktnavn As String
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
    Public varLobenr As string = "0"
    
    Public findesOmkostning As Integer = 0
    Public findesAkt As Integer = 0
    Public findesBudgetpost As Integer = 0
    
    Public aktsumGrandBel As String = "0"
    
    Public aktsumGrandTotalJob(5000) As Double
    Public matsumGrandTotalJob(5000) As Double
   
    
    <WebMethod()> Public Function addakt(ByVal ds As DataSet) As String
        
         'Return "Webservice Msg dt: "  
        
        
        'On Error Resume Next
    
        
        'Dim con As New SqlConnection(Application("MyDatabaseConnectionString"))
        'Dim daCust As New SqlDataAdapter("Select * From tblCustomer", con)
        'Dim cbCust As New SqlCommandBuilder(daCust)
        'daCust.Update(ds, "Cust")
        'Return ds
        
        
        
        
        'Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_intranet;User=outzource;Password=SKba200473;"

        Dim objConn As OdbcConnection
        Dim objConn2 As OdbcConnection
        Dim objCmd As OdbcCommand
        'Dim objDataSet As New DataSet
        Dim objDR As OdbcDataReader
        Dim objDR2 As OdbcDataReader
        Dim objDR3 As OdbcDataReader
        'Dim dt As DataTable
        'Dim dr As DataRow
        
        Dim strConn As String
        Dim t As Double = 0
        
        'Try
        
       
        
        'strConn = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_epi_catitest;User=outzource;Password=SKba200473;"
        'strConn = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_oko;User=to_outzource2;Password=SKba200473;"
        'strConn = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=root;pwd=;database=timeout_oko; OPTION=32"
        strConn = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_oko;User=to_outzource2;Password=SKba200473;"
                  
                    
      
        
        '** Åbner Connection ***'
        objConn = New OdbcConnection(strConn)
        objConn.Open()
        
        '*** HENTER AKT fra AKT_IMPORT_TEMP '****
        Dim strSQLakts As String = "SELECT id, dato, editor, origin, jobnr, aktnavn, aktnr, akttimer, akttpris, aktsum, lto, beskrivelse, aktkonto, akttype FROM akt_import_temp WHERE id > 0 AND overfort = 0 AND errid = 0"
        objCmd = New OdbcCommand(strSQLakts, objConn)
        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
        Dim a As Integer = 0
         
        
        
        While objDR.Read() = True
            
            
            
            aktsumTilSum = 0
            
            errThis = 0

            id = objDR("id")
            dato = objDR("dato")
            origin = objDR("origin")
            editor = objDR("editor")
            
            jobNr = objDR("jobnr")
            
            aktnavn = objDR("aktnavn")
            aktnavn = EncodeUTF8(aktnavn)
            aktnavn = DecodeUTF8(aktnavn)
            
            aktnr = objDR("aktnr")
           
            
            akttimer = objDR("akttimer")
            akttimer = Replace(akttimer, ",", ".")
            
            akttpris = objDR("akttpris")
            akttpris = Replace(akttpris, ",", ".")
            
            aktsum = objDR("aktsum")
            aktsum = Replace(aktsum, ",", ".")
            
            aktkonto = objDR("aktkonto")
            akttype = objDR("akttype")
            
            lto = objDR("lto")
            
            'beskrivelse = IsDBNull(objDR("beskrivelse")) '** VARCHAR FOR IKKE AT FEJLE (TEXT fejler)
            'If beskrivelse <> True Then
            ' beskrivelse = objDR("beskrivelse")
            'beskrivelse = beskrivelse.ToString()
            'Else
            'beskrivelse = ""
            'End If
            
            
          
            '**** Henter JobID OG Tjekker om jobnr findes '****
            If Len(Trim(jobNr)) <> 0 Then
                jobNr = jobNr
            Else
                jobNr = 0
            End If
            
            objConn2 = New OdbcConnection(strConn)
            objConn2.Open()
            Dim strSQLjobans As String = "SELECT id, jobstartdato, jobslutdato FROM job WHERE jobnr = '" + jobNr + "'"
            objCmd = New OdbcCommand(strSQLjobans, objConn2)
            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            jobId = 0
            If objDR2.Read() = True Then
                
                jobId = objDR2("id")
                jobstartdato = objDR2("jobstartdato")
                jobslutdato = objDR2("jobslutdato")
                
            End If
            objDR2.Close()
            objConn2.Close()
            
            If jobId = 0 Then
                errThis = 21
            End If
            
            
            '**** Henter kontonr og ID frakontoplan'****
            If Len(Trim(aktkonto)) <> 0 Then
                aktkonto = aktkonto
            Else
                aktkonto = 0
            End If
            
            objConn2 = New OdbcConnection(strConn)
            objConn2.Open()
            Dim strSQLkontonr As String = "SELECT id FROM kontoplan WHERE kontonr = " + aktkonto
            objCmd = New OdbcCommand(strSQLkontonr, objConn2)
            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            kontoId = 0
            If objDR2.Read() = True Then
                
                kontoId = objDR2("id")
                
            End If
            objDR2.Close()
            objConn2.Close()
            
            If kontoId = 0 Then
                errThis = 22
            End If
            
   
            
              
                                                 
                '*** Opretter aktiviteter ***'
            If CInt(errThis) = 0 Then
                
               varLobenr = jobId+"0"+aktkonto+"0"+aktnr
                        
                '*** Konnr > 130 indlæses som SalgsOmkostninger
                   
                If Trim(akttype) = "Budget" Then
                
                    If CInt(aktkonto) >= 101 And CInt(aktkonto) <= 199 Then
                        'Dim acc As Integer = 100
                        'If CInt(acc) = 100 Then
                    
                        '** Løn kat BUDGET Opdater / Insert ***'
                
                        '*** Hvis findes -_> UPDATE
                        objConn2 = New OdbcConnection(strConn)
                        objConn2.Open()
                        Dim strSQLaktfindes As String = "SELECT extsysid FROM aktiviteter WHERE extsysid = " & varLobenr & " AND job = " & jobId
                        objCmd = New OdbcCommand(strSQLaktfindes, objConn2)
                        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                        findesAkt = 0
                        If objDR2.Read() = True Then
                
                            findesAkt = 1
                
                        End If
                        objDR2.Close()
                        objConn2.Close()
                        
                        If aktsum <> 0 Then
                        
                             If CInt(findesAkt) = 0 then
                            
                                objConn2 = New OdbcConnection(strConn)
                                objConn2.Open()
                                
                                Dim strSQLaktins As String = ("INSERT INTO aktiviteter (editor, dato, navn, aktnr, job, fakturerbar, " _
                                  & "projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7," _
                                  & "projektgruppe8, projektgruppe9, projektgruppe10, aktstatus, budgettimer, aktbudget, aktbudgetsum, aktstartdato, aktslutdato, aktkonto, fase, extsysid) VALUES " _
                                  & " ('TO_import-nav','" & Now.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', '" & aktnavn & "', " & aktkonto & "," & jobId & ", 90," _
                                  & " 10,1,1,1,1,1,1,1,1,1,1," & akttimer & ", " & akttpris & ", " & aktsum & ",'" & jobstartdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "'" _
                                  & ",'" & jobslutdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', " & kontoId & ", '01. NAV-SUM', " & varLobenr & ")")
                    
                                'Return strSQLaktins
                        
                                objCmd = New OdbcCommand(strSQLaktins, objConn2)
                                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                objDR2.Close()
                                objConn2.Close()
                                
                                
                            Else
                                
                                Dim strSQLaktupd As String = ("UPDATE aktiviteter SET editor = 'TO_import-nav', dato = '" & Now.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', " _
                                & " navn = '" & aktnavn & "', aktnr = " & aktkonto & ", job = " & jobId & ",  " _
                                & " budgettimer = "& akttimer & ", aktbudget = " & akttpris & ", aktbudgetsum = " & aktsum & ", aktkonto = " & kontoId & ", fase = '01. NAV-SUM' WHERE extsysid = " & varLobenr & " AND job = " & jobId)
                    
                                'Return strSQLaktins
                        
                                objConn2 = New OdbcConnection(strConn)
                                objConn2.Open()
                                objCmd = New OdbcCommand(strSQLaktupd, objConn2)
                                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                objDR2.Close()
                                objConn2.Close()
                                
                            End if
                            
                        
                            'aktsumGrandTotalJob = aktsumGrandTotalJob + aktsum
                            aktsumTilSum = Replace(aktsum, ".", ",")
                            aktsumGrandTotalJob(jobId) = aktsumGrandTotalJob(jobId) + aktsumTilSum/1  
                           
                        
                        End If 'aktsum
                    
                    
                           
                    Else
                    
                        'Eksternbistand / Salgsomkostninger BUDGET
                       
                        '*** Hvis findes -_> UPDATE
                        
                        objConn2 = New OdbcConnection(strConn)
                        objConn2.Open()
                        Dim strSQLaktfindes As String = "SELECT extsysid FROM job_ulev_ju WHERE extsysid = " & varLobenr & " AND ju_jobid = " & jobId
                        objCmd = New OdbcCommand(strSQLaktfindes, objConn2)
                        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                        findesBudgetpost = 0
                        If objDR2.Read() = True Then
                
                            findesBudgetpost = 1
                
                        End If
                        objDR2.Close()
                        objConn2.Close()
                        
                        If aktsum <> 0 Then
                            
                            
                            If CInt(findesBudgetpost) = 1 Then 'Update
                                
                                Dim strSQLsalgsomkupd As String = ("UPDATE job_ulev_ju SET ju_navn = '" & aktnavn & "', ju_ipris = " & akttpris & ", ju_faktor = 1, ju_belob = " & aktsum & ", ju_jobid = "& jobId &"" _
                                &", ju_favorit = 0, ju_fase = '', ju_stk = " & akttimer & ", ju_stkpris = " & akttpris & ", ju_fravalgt = 0, ju_konto = "& kontoId &" WHERE extsysid = " & varLobenr &"  AND ju_jobid = " & jobId) 
                        
                                'return strSQLsalgsomkins
                                
                                objConn2 = New OdbcConnection(strConn)
                                objConn2.Open()
                                
                                objCmd = New OdbcCommand(strSQLsalgsomkupd, objConn2)
                                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                objDR2.Close()
                                objConn2.Close()
                                
                            Else
                                
                            
                            
                                Dim strSQLsalgsomkins As String = ("INSERT INTO job_ulev_ju (ju_navn, ju_ipris, ju_faktor, ju_belob, ju_jobid, ju_favorit, ju_fase, ju_stk, ju_stkpris, ju_fravalgt, extsysid, ju_konto) VALUES " _
                                & " ('" & aktnavn & "', " & akttpris & ", 1, " & aktsum & ", " & jobId & ", 0,'', " & akttimer & ", " & akttpris & ", 0, " & varLobenr & ", "& kontoId &")")
                        
                                'return strSQLsalgsomkins
                        
                                objConn2 = New OdbcConnection(strConn)
                                objConn2.Open()
                                objCmd = New OdbcCommand(strSQLsalgsomkins, objConn2)
                                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                objDR2.Close()
                                objConn2.Close()
                                
                                End If
                            
                            End If
                    
                       
                    
                        'matsumGrandTotalJob = matsumGrandTotalJob + aktsum
                        aktsumTilSum = Replace(aktsum, ".", ",")
                        matsumGrandTotalJob(jobId) = matsumGrandTotalJob(jobId) + aktsumTilSum/1 
                        
                    
                    End If 'aktsum
                    
                Else '** Kontrakt
                    
                    'Omkostninger 'Kun indlæs
                    
                    objConn2 = New OdbcConnection(strConn)
                    objConn2.Open()
                    
                    Dim strSQLomkostninger As String = "SELECT extsysid FROM materiale_forbrug WHERE extsysid = " & varLobenr & " AND jobid = " & jobId
                    objCmd = New OdbcCommand(strSQLomkostninger, objConn2)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    findesOmkostning = 0
                    If objDR2.Read() = True Then
                
                        findesOmkostning = 1
                
                    End If
                    objDR2.Close()
                    objConn2.Close()
            
                    If CInt(findesOmkostning) = 0 And aktsum <> 0 Then 'SKAL TILFØJES NÅR vi begyndner at indlæse fra TiemOut til NAV i feb: And CInt(aktkonto) >= 199
                    
                        Dim strSQLsalgsomkins As String = ("INSERT INTO materiale_forbrug (matantal, matnavn, matvarenr, matkobspris, matsalgspris, matenhed, jobid, " _
                        & " dato, editor, usrid, matgrp, forbrugsdato, extsysid, mf_konto) VALUES " _
                        & " (" & akttimer & ", '" & aktnavn & "', '" & aktkonto & "', " & akttpris & ", " & akttpris & ", 'Stk.', " & jobId & ", " _
                        & "'" & Now.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "','TO_import-nav', 71, 1, '" & Now.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', " & varLobenr & ", "& kontoId &")")
                        
                        'Return "XX HER XX errThis SQL: " & strSQLsalgsomkins
                        objConn2 = New OdbcConnection(strConn)
                        objConn2.Open()
                        objCmd = New OdbcCommand(strSQLsalgsomkins, objConn2)
                        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                        objDR2.Close()
                        objConn2.Close()
                        
                        
                    End If
                    
                    
                End If
                    
              
                
                  
                                                    
                
                
                
                                                
               
                    
            End If ' 21 jobId findes IKKE i forvejen
            
            
       
        
            a = a + 1
            
            
        End While
        objDR.Close()
        objConn.Close() 
        
        
        
            
        '************* Overfører beløb til GT på job *********************
        
        'For Each element As Integer In aktsumGrandTotalJob
        Dim i As Integer = 0
        
        For i = 0 To UBound(aktsumGrandTotalJob)
            
           
        
            Dim jobBruttoOms As Double = (aktsumGrandTotalJob(i) / 1 + matsumGrandTotalJob(i) / 1)
            Dim jobBruttoOmsBel As String = jobBruttoOms
            
            If jobBruttoOms <> 0 Then
            
                'jobBruttoOms = Replace(jobBruttoOms, ".", "")
                jobBruttoOmsBel = Replace(jobBruttoOmsBel, ",", ".")
        
        
                
                
                
                'aktsumGrandTotalJob(i) = Replace(aktsumGrandTotalJob(i), ".", "")
                'aktsumGrandTotalJob(i) = aktsumGrandTotalJob(i) / 1000
                aktsumGrandTotalJob(i) = aktsumGrandTotalJob(i) 
                aktsumGrandBel = Replace(aktsumGrandTotalJob(i), ",", ".")
        
                objConn2 = New OdbcConnection(strConn)
                objConn2.Open()
                Dim strSQLjobbruttoOms As String = "UPDATE job SET jo_bruttooms = " & jobBruttoOmsBel & ", jobTpris = " & aktsumGrandBel & ", jo_gnsbelob = " & aktsumGrandBel & ", jo_udgifter_intern = " & aktsumGrandBel & ", job_internbesk = '222-" & aktsumGrandBel & "', budgettimer = 1, jo_gnsfaktor = 1 WHERE id = " & i
                'Dim strSQLjobbruttoOms As String = "UPDATE job SET jo_bruttooms = " & jobBruttoOmsBel & ", jobTpris = 0, jo_gnsbelob = 0, job_internbesk = '" & aktsumGrandBel & "', budgettimer = 1, jo_gnsfaktor = 1 WHERE id = " & i
                objCmd = New OdbcCommand(strSQLjobbruttoOms, objConn2)
                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                objDR2.Close()
                objConn2.Close()
            
                
               end if
            
        Next
        '*****************************************************************
        
        
        
        
        
        '**** Opdater overført ************
        
        objConn2 = New OdbcConnection(strConn)
        objConn2.Open()
        Dim strSQLjobToverfort As String = "UPDATE akt_import_temp SET overfort = 1 WHERE id > 0 AND errid = 0"
        objCmd = New OdbcCommand(strSQLjobToverfort, objConn2)
        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
        objDR2.Close()
        objConn2.Close()
                
                                                    
                   
        
        
    End Function

    
    
    'Public Function rtErrTxt(ByVal erTxt As String) As String
    ' Dim eTxt As String
    '    eTxt = erTxt
        
    '   Return eTxt
    'End Function
   
    
End Class

        




