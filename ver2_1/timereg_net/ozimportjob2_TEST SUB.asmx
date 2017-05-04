



<%@ WebService language="VB" class="oz_importjob2" %>
Imports System
Imports System.Web.Services
Imports System.Data
Imports System.Data.Odbc

'***************************************************************************************
'*** DENNE SERVICE BENYTTES BL.A AF 
'*** OKO
'*** Wilke (bruger ozimportjob2_ds)
'*** DENCKER
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
    Public jobnrTjk As Integer = 0
    Public jobID As Integer = 0
    Public timereguseJobFindes As Integer = 0
    Public projektKategori As Integer = 0
    
    <WebMethod()> Public Function createjob2(ByVal ds As dataset) As String
        
         
         
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
               
            Next
            
        Next
        
        
                
   
        Dim objConn As OdbcConnection
        Dim objConn2 As OdbcConnection
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
        
        Dim objCmd10 As OdbcCommand
        'Dim objCmd11 As OdbcCommand
        Dim objDR10 As OdbcDataReader
        Dim objDR11 As OdbcDataReader
        
            Dim strConn As String
            Dim t As Double = 0
        
            'Try
           
        
       
        
            strConn = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" & lto & ";User=to_outzource2;Password=SKba200473;"
            'strConn = "Driver={MySQL ODBC 3.51 Driver};Server=195.189.130.210;Database=timeout_oko;User=to_outzource2;Password=SKba200473;"
            'strConn = "Provider=MSDASQL;driver={MySQL ODBC 3.51 Driver}; Server=195.189.130.210; Port=3306; User=outzource; Password=SKba200473; Database=timeout_oko; Option=3;"
       
            'strConn = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=root;pwd=;database=timeout_oko; OPTION=32"
            'strConn = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=root;pwd=;database=timeout_wilke; OPTION=32"
                  
       
        
       
        
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
            
                'jobnavn = EncodeUTF8(objDR("jobnavn"))
                'jobnavn = DecodeUTF8(jobnavn)
                'jobnavn = jq_format(objDR("jobnavn"))
                'jobnavn = jq_formatTxt
                jobnavn = Replace(objDR("jobnavn"), "'", "")
                jobnavn = Replace(jobnavn, "''", "")
                jobnavn = Replace(jobnavn, ";", "")
            
                jobnr = objDR("jobnr")
            
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
            
            
                jobnrTjk = objDR("jobnr")
                kundenavnTxt = beskrivelse
            
           

            
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
                Dim opdaterJob As Integer = 0
                Dim strSQLjobnrfindes As String = "SELECT jobnr FROM job WHERE jobnr = '" + strjobnr + "'"
                objCmd = New OdbcCommand(strSQLjobnrfindes, objConn)
                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                jobans = 0
                If objDR2.Read() = True Then
                
                    errThis = 0
                    opdaterJob = 1
                
                End If
                objDR2.Close()
            
            
                                                 
                '*** Opretter Job ***'
              
                        
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
                            Dim strSQKkundeins As String = "INSERT INTO kunder SET kkundenavn = '" + kundenavnTxt + "', kkundenr = " + jobnr + ", ketype = 'ke', ktype = 0"
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
                            
                        End If
                        
                      
                End Select
                    
                
                kunderef = 0
                rekvisitionsnr = 0
                bruttooms = 0
                internnote = ""
                
               
                
                'Return "Webservice Msg dt: " & jobstartdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture)
                
                'Call DecodeUTF8(jobnavn)
                'jobnavn = 
            
            
                If CInt(errThis) = 0 Then
                
            
                  
               
                    
                    'Return "Webservice Msg dt jobnrTjk: "+ jobnrTjk +" opretJobOk: "+ opretJobOk 
                opretJobOk = 1
                If CInt(opretJobOk) = 1 Then
                    
                    objConn2 = New OdbcConnection(strConn)
                     objConn2.Open()
                
                    If CInt(opdaterJob) = 1 Then 'opdater
                    
                    
                        Dim strSQLjobUpd As String = ("Update job SET jobnavn = '" & jobnavn & "', jobnr = '" & jobnr & "', jobstatus = 1, " _
                        & " jobstartdato = '" & jobstartdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "'," _
                        & " jobslutdato = '" & jobslutdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', editor = '" & editor & "', " _
                        & " dato = '" & dato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', beskrivelse = '" & beskrivelse & "', jobans1 = " & jobans & ", " _
                        & " kundekpers = " & kunderef & " WHERE jobnr = '" & jobnr & "'")
                  
                        objCmd = New OdbcCommand(strSQLjobUpd, objConn)
                        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                        objDR2.Close()
                    
                        
                        '*** Tilføjer til timereg_usejob (for en sikkerhedsskyld)
                        
                        '*** Finder jobid ***
                        jobID = 0
                        
                        
                        '**************************************************************
                        '**** Sikrer timereg usejob bliver udfyldt ved rediger job ****
                        Dim strSQLlastJobID As String = "SELECT id FROM job WHERE jobnr = '" & jobnr & "'"
                        objCmd = New OdbcCommand(strSQLlastJobID, objConn)
                        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                
                        If objDR2.Read() = True Then
                
                            jobID = objDR2("id")
                
                        End If
                        objDR2.Close()
                        
                        
                       
                        
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
                        
                        
                        
                        '**************************************************************'
                        '*** Hvis timepris ikke findes på job bruges Gen. timepris fra '
                        '*** Fra medarbejdertype, og den oprettes på job              *'
                        '*** BLIVER ALTID HENTET FRA Medarb.type for ØKO              *'
                        '**************************************************************'
                        
                        Select Case lto
                            Case "oko"
                       
                        
                        
                            'Call opdaterTimerpriser(lto, opdaterJob, strConn)
                                
                        End Select
                            
                        
                    
                    Else 'opret
                
                    
                        
                
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
                                & "'" & jobstartdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', " _
                                & "'" & jobslutdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', " _
                                & "'" & editor & "', " _
                                & "'" & dato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', " _
                                & "10, " _
                                & "1,1,1,1,1,1,1,1,1," _
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
                                    Case "wilke"
                       
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
                            & " WHERE mid <> 0 AND mansat = 1 AND medarbejdertyper.id = medarbejdertype"

                            objCmd10 = New OdbcCommand(SQLmedtpris, objConn2)
                            objDR10 = objCmd10.ExecuteReader '(CommandBehavior.closeConnection)
            
                            While objDR10.Read() = True
		 	
                                If objDR10("kostpris") <> 0 Then
                                    kostpris = objDR10("kostpris")
                                Else
                                    kostpris = 0
                                End If
		
                            
                                Select Case lto
                                    Case "oko"
                                    
                                    
                                        Select Case Left(Trim(beskrivelse), 8)
                                            Case "NaturErh"
                                    
                                                If objDR10("timepris_a2") <> 0 Then
                                      
                                                    tprisGen = objDR10("timepris_a2")
                                                Else
                                                    tprisGen = 0
                                                End If
                                    
                                                timeprisalt = 2
                                    
                                            Case "Direktor"
                                    
                                                If objDR10("timepris_a3") <> 0 Then
                                      
                                                    tprisGen = objDR10("timepris_a3")
                                                Else
                                                    tprisGen = 0
                                                End If
                                    
                                                timeprisalt = 3
                                    
                                            Case Else
                                
                                                If objDR10("timepris") <> 0 Then
                                     
                                                    tprisGen = objDR10("timepris")
                                                Else
                                                    tprisGen = 0
                                                End If
                                    
                                                timeprisalt = 0
                                    
                                        End Select
                                    
                                    
                                    Case Else
                                    
                                        If objDR10("timepris") <> 0 Then
                                     
                                            tprisGen = objDR10("timepris")
                                        Else
                                            tprisGen = 0
                                        End If
                                    
                                        timeprisalt = 0
                           
                            
                                   
                                    
                                End Select
                            
                            
                        
                        
                                valutaGen = objDR10("tp0_valuta")
                        
                        
                                '**** Indlæser timepris på aktiviteter ***'
                                intTimepris = tprisGen
                                intTimepris = Replace(intTimepris, ",", ".")
                            
                                intValuta = valutaGen
            
                                If opdaterJob = 1 Then 'opdater
                
                                Else
                
                                    Dim strSQLtpris As String = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) " _
                                    & " VALUES (" & lastID & ", " & lastAktID & ", " & objDR10("mid") & ", " & timeprisalt & ", " & intTimepris & ", " & intValuta & ")"
							
                                    objCmd10 = New OdbcCommand(strSQLtpris, objConn2)
                                    objDR11 = objCmd10.ExecuteReader '(CommandBehavior.closeConnection)
                        
                                End If
            
                            End While

                            objDR10.Close()
                    
                                    '** Slut timepris **
                
                
                   
							
                    
                    
                    
                                End While
                                objDR3.Close()
                
                
                
                                                
                    End If 'Opdater / opret
                
                    objConn2.Close()
                End If 'opretJobOk kundekundet
                    
                
                
                End If 'CInt(errThis) = 0
            
            
                
       
            
          
            End While
            objDR.Close()
            
            
            'Catch ex As Exception
            '   MsgBox("Can't load Web page" & vbCrLf & ex.Message)
            'End Try
        
        
            Dim strSQLjobToverfort As String = "UPDATE job_import_temp SET overfort = 1 WHERE id > 0 AND errid = 0"
            objCmd = New OdbcCommand(strSQLjobToverfort, objConn)
            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            objDR2.Close()
                
                                                    
            Dim strSQLjobnrUpd As String = "UPDATE licens SET jobnr = " & jobnr & " WHERE id = 1"
            objCmd = New OdbcCommand(strSQLjobnrUpd, objConn)
            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            objDR2.Close()
                        
        objConn.Close()
        
        
        
      
        
        
    End Function

    
     
    'Sub opdaterTimepriser(lto As String, func As String)
    Public Sub XopdaterTimerpriser(lto As String, opdaterJob As Integer, strConn As string, objConn As OdbcConnection)
        'Throw New NotImplementedException
        
      
        Dim objCmd10 As OdbcCommand
        'Dim objCmd11 As OdbcCommand
        Dim objDR10 As OdbcDataReader
        Dim objDR11 As OdbcDataReader
        
        
        tprisGen = 0
        valutaGen = 1
        kostpris = 0
        intTimepris = 0
        intValuta = valutaGen
       
        
   
                    
        Dim SQLmedtpris As String = "SELECT medarbejdertype, timepris, timepris_a2, timepris_a3, tp0_valuta, kostpris, mnavn, mid FROM medarbejdere " _
        & " LEFT JOIN medarbejdertyper ON (medarbejdertyper.id = medarbejdertype) " _
        & " WHERE mid <> 0 AND mansat = 1 AND medarbejdertyper.id = medarbejdertype"

        objCmd10 = New OdbcCommand(SQLmedtpris, objConn)
        objDR10 = objCmd10.ExecuteReader '(CommandBehavior.closeConnection)
            
        While objDR10.Read() = True
		 	
            If objDR10("kostpris") <> 0 Then
                kostpris = objDR10("kostpris")
            Else
                kostpris = 0
            End If
		
                            
            Select Case lto
                Case "oko"
                                    
                                    
                    Select Case Left(Trim(beskrivelse), 8)
                        Case "NaturErh"
                                    
                            If objDR10("timepris_a2") <> 0 Then
                                      
                                tprisGen = objDR10("timepris_a2")
                            Else
                                tprisGen = 0
                            End If
                                    
                            timeprisalt = 2
                                    
                        Case "Direktor"
                                    
                            If objDR10("timepris_a3") <> 0 Then
                                      
                                tprisGen = objDR10("timepris_a3")
                            Else
                                tprisGen = 0
                            End If
                                    
                            timeprisalt = 3
                                    
                        Case Else
                                
                            If objDR10("timepris") <> 0 Then
                                     
                                tprisGen = objDR10("timepris")
                            Else
                                tprisGen = 0
                            End If
                                    
                            timeprisalt = 0
                                    
                    End Select
                                    
                                    
                Case Else
                                    
                    If objDR10("timepris") <> 0 Then
                                     
                        tprisGen = objDR10("timepris")
                    Else
                        tprisGen = 0
                    End If
                                    
                    timeprisalt = 0
                           
                            
                                   
                                    
            End Select
                            
                            
                        
                        
            valutaGen = objDR10("tp0_valuta")
                        
                        
            '**** Indlæser timepris på aktiviteter ***'
            intTimepris = tprisGen
            intTimepris = Replace(intTimepris, ",", ".")
                            
            intValuta = valutaGen
            
            If opdaterJob = 1 Then 'opdater
                
            Else
                
                Dim strSQLtpris As String = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) " _
                & " VALUES (" & lastID & ", " & lastAktID & ", " & objDR10("mid") & ", " & timeprisalt & ", " & intTimepris & ", " & intValuta & ")"
							
                objCmd10 = New OdbcCommand(strSQLtpris, objConn)
                objDR11 = objCmd10.ExecuteReader '(CommandBehavior.closeConnection)
                        
            End If
            
        End While

        objDR10.Close()
        '** Slut timepris **
        
        
    End Sub
   
    
    
    'Public Function rtErrTxt(ByVal erTxt As String) As String
    ' Dim eTxt As String
    '    eTxt = erTxt
        
    '   Return eTxt
    'End Function

    Private Sub opdaterTimepriser(lto As String, opdaterJob As Integer)
        Throw New NotImplementedException
    End Sub

    Private Sub opdaterTimepriser(lto As String, opdaterJob As Integer, strConn As String)
        Throw New NotImplementedException
    End Sub


   
    
    

   
    
End Class

        




