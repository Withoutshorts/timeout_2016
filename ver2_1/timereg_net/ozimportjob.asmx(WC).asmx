



<%@ WebService language="VB" class="oz_importjob" %>
Imports System
Imports System.Web.Services
Imports System.Data
Imports System.Data.Odbc





Public Class oz_importjob : 

    
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
    Public beskrivelse As String
    Public errThis As Integer = 0
    Public kid As Integer
     
    Public kunderef As Integer = 0
    Public rekvisitionsnr As String
    Public bruttooms As Double
    Public internnote As String = ""
    Public jobansInit As String = ""
    Public strjobnr As String = ""
     
    Public strantalArbPakker As String = "0"
    
    <WebMethod()> Public Function xcreatejob(ByVal ds As DataSet) As String
        
         
        
        
        'On Error Resume Next
    
        
        'Dim con As New SqlConnection(Application("MyDatabaseConnectionString"))
        'Dim daCust As New SqlDataAdapter("Select * From tblCustomer", con)
        'Dim cbCust As New SqlCommandBuilder(daCust)
        'daCust.Update(ds, "Cust")
        'Return ds
        
        
        
        
        'Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_intranet;User=outzource;Password=SKba200473;"

        Dim objConn As OdbcConnection
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
        strConn = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_oko;User=to_outzource2;Password=SKba200473;"
        'strConn = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=root;pwd=;database=timeout_oko; OPTION=32"
                  
                    
        '** Åbner Connection ***'
        objConn = New OdbcConnection(strConn)
        objConn.Open()
        
        '*** HENTER JOB FRA JOB_IMPORT_TEMP '****
        Dim strSQLjnj As String = "SELECT id, dato, editor, origin, jobnavn, jobnr, jobstartdato, jobslutdato, jobans, lto, beskrivelse FROM job_import_temp WHERE id > 0 AND overfort = 0 AND errid = 0"
        objCmd = New OdbcCommand(strSQLjnj, objConn)
        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
        While objDR.Read() = True
            
            errThis = 0

            id = objDR("id")
            dato = objDR("dato")
            origin = objDR("origin")
            editor = objDR("editor")
            jobnavn = objDR("jobnavn")
            jobnr = objDR("jobnr")
            jobstartdato = objDR("jobstartdato")
            jobslutdato = objDR("jobslutdato")
            jobansInit = objDR("jobans")
            lto = objDR("lto")
            beskrivelse = objDR("beskrivelse")
            
            
            
           

            
            '**** Henter Jobans '****
            Dim strSQLjobans As String = "SELECT mid FROM medarbejdere WHERE init = '" + jobansInit + "'"
            objCmd = New OdbcCommand(strSQLjobans, objConn)
            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            jobans = 0
            If objDR2.Read() = True Then
                
                jobans = objDR2("mid")
                
            End If
            objDR2.Close()
            
            
            strjobnr = jobnr.ToString
            
           
            '*** Tjekker om jobnr findes ***
            Dim strSQLjobnrfindes As String = "SELECT jobnr FROM job WHERE jobnr = '" + strjobnr + "'"
            objCmd = New OdbcCommand(strSQLjobnrfindes, objConn)
            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            jobans = 0
            If objDR2.Read() = True Then
                
                errThis = 11
                
            End If
            objDR2.Close()
            
            
                                                 
            '*** Opretter Job ***'
            If CInt(errThis) = 0 Then
                                                    
                        
                kid = 1 ' ALTID ØKOLOGISK
                kunderef = 0
                rekvisitionsnr = 0
                bruttooms = 0
                internnote = ""

                Dim strSQLjob As String = ("INSERT INTO job (jobnavn, jobnr, jobknr, jobTpris, jobstatus, jobstartdato," _
                & " jobslutdato, editor, dato, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, " _
                & " projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, " _
                & " fakturerbart, budgettimer, fastpris, kundeok, beskrivelse, " _
                & " ikkeBudgettimer, tilbudsnr, jobans1, " _
                & " serviceaft, kundekpers, valuta, rekvnr, " _
                & " risiko, job_internbesk, " _
                & " jo_bruttooms, fomr_konto) VALUES " _
                & "('" & jobnavn & "', " _
                & "'" & jobnr & "', " _
                & "" & kid & ", " _
                & "0, " _
                & "1, " _
                & "'" & jobstartdato & "', " _
                & "'" & jobslutdato & "', " _
                & "'" & editor & "', " _
                & "'" & dato & "', " _
                & "10, " _
                & "1,1,1,1,1,1,1,1,1," _
                & "1,0,1,0," _
                & "'" & beskrivelse & "', " _
                & "0,0, " _
                & "" & jobans & "," _
                & "0," & kunderef & ", " _
                & "1, '" & rekvisitionsnr & "', " _
                & "100,'" & internnote & "'," _
                & "" & bruttooms & ", 0)")
    							
                    
        
                objCmd = New OdbcCommand(strSQLjob, objConn)
                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                objDR2.Close()
                                                    
                
                
                '*** Finder jobid ***
                Dim strSQLlastJobID As String = "SELECT id FROM job WHERE id <> 0 ORDER BY id DESC"
                objCmd = New OdbcCommand(strSQLlastJobID, objConn)
                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                Dim lastID As Integer = 0
                If objDR2.Read() = True Then
                
                    lastID = objDR2("id")
                
                End If
                objDR2.Close()
                
                
                'Dim antalArbPakker As Array
                'antalArbPakker = strantalArbPakker
                
                '*** Indlæser aktiviteter (arbejdspakker) - aktiviteter fra Sagsopgavelinjer indlæses separat ***'
                Dim strSQLStamAkt As String = "SELECT a.navn AS aktnavn FROM akt_gruppe AS ag" _
                & " LEFT JOIN aktiviteter AS a ON (a.aktfavorit = ag.id) WHERE ag.forvalgt = 1 ORDER BY a.sortorder DESC"
                objCmd = New OdbcCommand(strSQLStamAkt, objConn)
                objDR3 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                While objDR3.Read() = True
                
                    Dim strSQLaktins As String = ("INSERT INTO aktiviteter (navn, job, fakturerbar, " _
                    & "projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7," _
                    & "projektgruppe8, projektgruppe9, projektgruppe10, aktstatus, budgettimer, aktbudget, aktbudgetsum, aktstartdato, aktslutdato) VALUES " _
                    & " ('" & objDR3("aktnavn") & "', " & lastID & ", 1," _
                    & " 10,1,1,1,1,1,1,1,1,1,1,20,5,400,'2015-11-01', '2015-12-20')")
                                               
                    objCmd = New OdbcCommand(strSQLaktins, objConn)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    objDR2.Close()
                    
                    
                End While
                objDR3.Close()
                                                
               
                    
            End If ' 6 findes i forvejen
            
            
       
            
          
        End While
        objDR.Close()
            
            
        'Catch ex As Exception
        '    MsgBox("Can't load Web page" & vbCrLf & ex.Message)
        'End Try
        
        
        Dim strSQLjobToverfort As String = "UPDATE job_import_temp SET overfort = 1 WHERE id > 0 AND errid = 0"
        objCmd = New OdbcCommand(strSQLjobToverfort, objConn)
        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
        objDR2.Close()
                
                                                    
        Dim strSQLjobnrUpd As String = "UPDATE licens SET jobnr = " & jobnr & " WHERE id = 1"
        objCmd = New OdbcCommand(strSQLjobnrUpd, objConn)
        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
        objDR2.Close()
                                           
        
        
    End Function

    
    
    'Public Function rtErrTxt(ByVal erTxt As String) As String
    ' Dim eTxt As String
    '    eTxt = erTxt
        
    '   Return eTxt
    'End Function
   
    
End Class

        




