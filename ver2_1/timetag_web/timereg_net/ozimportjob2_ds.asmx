



<%@ WebService language="VB" class="oz_importjob2_ds" %>
Imports System
Imports System.Web.Services
Imports System.Data
Imports System.Data.Odbc



'***************************************************************************************
'*** DENNE SERVICE BENYTTES BL.A AF 
'*** Wilke
'***************************************************************************************



Public Class oz_importjob2_ds

    
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
    
    Public jobtimepris As string
    Public kundenr, valuta As string
    Public clientMid, andetMid, fieldMid, itMid, researchMid, salgsansvarligMid As Integer
    Public clientInit, andetInit, fieldInit, itInit, researchInit, salgsansvarligInit As String
    
    Public fieldTimer, researchTimer, itTimer, andetTimer, clientTimer, budgetkr, budgetTimer As String
    public aktidPaaJob, fcFindesPaaJob As integer
    
    
    
    <WebMethod()> Public Function createjob2_ds(ByVal ds As dataset) As String
        
         
        'HttpResponse.RemoveOutputCacheItem
         
        'On Error Resume Next
    
        
        'Dim con As New SqlConnection(Application("MyDatabaseConnectionString"))
        'Dim daCust As New SqlDataAdapter("Select * From tblCustomer", con)
        'Dim cbCust As New SqlCommandBuilder(daCust)
        'daCust.Update(ds, "Cust")
        'Return ds
        
        
        'Fra EXCEL / timer_import_temp
        'lto = ds.Tables("tb_to_var").Rows(t).Item("lto")
       
        
        Dim objConn As OdbcConnection
        Dim objConn2 As OdbcConnection
        Dim objConn3 As OdbcConnection
        'Dim objConn4 As OdbcConnection
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
        
        Dim strConn As String
        Dim t As Double = 0
        
        
        'lto = "demo"
        
        Dim dt As DataTable
        Dim dr As DataRow
        
        
        
        For Each dt In ds.Tables
            For Each dr In dt.Rows
                
                'If t < 10 then
              
                '** ORIGIN SKAL WEBSERVICE hente data fra timer_imp_temp / eller fra dataset
                '** 600 fra timer_imp_temp
                '** 601 fra dataset Wilke
                
                If t = 0 Then
                
                    
                    'If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("origin")) = False Then
                        
                   
                    'If ds.Tables("tb_to_var").Columns.Contains("origin") Then
                    'origin = ds.Tables("tb_to_var").Rows(t).Item("origin")
                    'Else
                    origin = 601
                    'End If
                    
                        
                    'Else
                    'origin = 699
                    'End If
              
                    
                    
                    lto = "wilke" '"wilke_webservicetest" '"" ' 'ds.Tables("tb_to_var").Rows(t).Item("lto")
             
                    strConn = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" & lto & ";User=to_outzource2;Password=SKba200473;"
                    'strConn = "timeout_wilke_webservicetest"
                    
                    'strConn = "Driver={MySQL ODBC 3.51 Driver};Server=195.189.130.210;Database=timeout_oko;User=to_outzource2;Password=SKba200473;"
                    'strConn = "Provider=MSDASQL;driver={MySQL ODBC 3.51 Driver}; Server=195.189.130.210; Port=3306; User=outzource; Password=SKba200473; Database=timeout_oko; Option=3;"
       
                    'strConn = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=root;pwd=;database=timeout_oko; OPTION=32"
                    'strConn = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=root;pwd=;database=timeout_wilke; OPTION=32"
                  
       
                    '** Åbner Connection ***'
                    objConn = New OdbcConnection(strConn)
                    objConn.Open()
                    
                   
                    
                    
                Else
                    
                    origin = origin
                    
                End If
                
                objConn2 = New OdbcConnection(strConn)
                objConn2.Open()
                    
                objConn3 = New OdbcConnection(strConn)
                objConn3.Open()
                
                
                '*******************************************************
                '** INPUT VARIABLE FRA DATASET 
                'jobnr                      = OK
                'jobnavn                    = OK
                'jobansInit                 = OK
                'jobstartdato               = OK
                'jobslutdato                = OK
                'Kundenavn                  = OK
                'Projekttype                = OK
                'lto                        = OK
                
                '**************************  HENTET fra dataset | Fundet i DB | Indlæst i DB incl. datamodel 
                'field                      = OK                | OK
                'fieldtimer                 = OK
                'research                   = OK                | OK
                'researchtimer              = OK
                'client                     = OK                | OK
                'clientimer                 = OK
                'it                         = OK                | OK
                'ittimer                    = OK
                'andet2' = xmif user                            | OK
                'andettimer                 = OK
                'kundenr                    = OK                |               | OK
                'budgetkr                   = OK
                'budgettimer                = OK
                'valuta                     = OK
                'salgsansvarlig             = OK
                
                

                
                '********************************************************
            
                errThis = 0

                'id = NOT IN USE
                    
                dato = Date.Now  '** dagsdato
                editor = lto + " webservice"
            
                
                
                '*** Medarb. IDs på fordeling af timer
                
                If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("salgsansvarlig")) = False Then
                    salgsansvarligInit = Trim(ds.Tables("tb_to_var").Rows(t).Item("salgsansvarlig"))
                Else
                    salgsansvarligInit = ""
                End If
                
                If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("field")) = False Then
                    fieldInit = Trim(ds.Tables("tb_to_var").Rows(t).Item("field"))
                Else
                    fieldInit = ""
                End If
                
                If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("research")) = False Then
                    researchInit = Trim(ds.Tables("tb_to_var").Rows(t).Item("research"))
                Else
                    researchInit = ""
                End If
                
                If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("client")) = False Then
                    clientInit = Trim(ds.Tables("tb_to_var").Rows(t).Item("client"))
                Else
                    clientInit = ""
                End If
                
                If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("it")) = False Then
                    itInit = Trim(ds.Tables("tb_to_var").Rows(t).Item("it"))
                Else
                    itInit = ""
                End If
                
                If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("andet2")) = False Then
                    andetInit = Trim(ds.Tables("tb_to_var").Rows(t).Item("andet2"))
                Else
                    andetInit = ""
                End If
                
                
                
                 '***************** Medarbejdere INIT *****************
                
                
                
                '**** Henter salgsansvarlig '****
                If salgsansvarligInit <> "" Then
                
                    Dim strSQLsalgsansvarlig As String = "SELECT mid FROM medarbejdere WHERE init = '" + salgsansvarligInit + "'"
                    objCmd = New OdbcCommand(strSQLsalgsansvarlig, objConn2)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    salgsansvarligMid = 0
                    If objDR2.Read() = True Then
                
                        salgsansvarligMid = objDR2("mid")
                
                    End If
                    objDR2.Close()
                
                End If
                
                
                
                '**** Henter client '****
                If clientInit <> "" Then
                
                    Dim strSQLclient As String = "SELECT mid FROM medarbejdere WHERE init = '" + clientInit + "'"
                    objCmd = New OdbcCommand(strSQLclient, objConn2)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    clientMid = 0
                    If objDR2.Read() = True Then
                
                        clientMid = objDR2("mid")
                
                    End If
                    objDR2.Close()
                    
                End If
                
                
                
                '**** Henter research '****
                If researchInit <> "" Then
                    
                    Dim strSQLresearch As String = "SELECT mid FROM medarbejdere WHERE init = '" + researchInit + "'"
                    objCmd = New OdbcCommand(strSQLresearch, objConn2)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    researchMid = 0
                    If objDR2.Read() = True Then
                
                        researchMid = objDR2("mid")
                
                    End If
                    objDR2.Close()
                    
                End If
                
                
                
                '**** Henter it '****
                If itInit <> "" then
                    Dim strSQLit As String = "SELECT mid FROM medarbejdere WHERE init = '" + itInit + "'"
                    objCmd = New OdbcCommand(strSQLit, objConn2)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    itMid = 0
                    If objDR2.Read() = True Then
                
                        itMid = objDR2("mid")
                
                    End If
                    objDR2.Close()
                    
                 end if
                
                
                
                
                '**** Henter field '****
                If fieldInit <> "" then
                
                    Dim strSQLfield As String = "SELECT mid FROM medarbejdere WHERE init = '" + fieldInit + "'"
                    objCmd = New OdbcCommand(strSQLfield, objConn2)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    fieldMid = 0
                    If objDR2.Read() = True Then
                
                        fieldMid = objDR2("mid")
                
                    End If
                    objDR2.Close()
                    
                  end if
                
                
                
                
                   
                '**** Henter andet '****
                If andetInit <> "" Then
                
                    Dim strSQLandet As String = "SELECT mid FROM medarbejdere WHERE init = '" + andetInit + "'"
                    objCmd = New OdbcCommand(strSQLandet, objConn2)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    andetMid = 0
                    If objDR2.Read() = True Then
                
                        andetMid = objDR2("mid")
                
                    End If
                    objDR2.Close()
                
                End If
                
                
                    'Timer fordelt på INIT (grupper)
                
                
                    If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("fieldtimer")) = False Then
                        fieldTimer = Trim(ds.Tables("tb_to_var").Rows(t).Item("fieldtimer"))
                    Else
                    fieldTimer = "0"
                    End If
                
                 
                    If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("researchtimer")) = False Then
                        researchTimer = Trim(ds.Tables("tb_to_var").Rows(t).Item("researchtimer"))
                    Else
                    researchTimer = "0"
                    End If
                
                    If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("ittimer")) = False Then
                        itTimer = Trim(ds.Tables("tb_to_var").Rows(t).Item("ittimer"))
                    Else
                    itTimer = "0"
                    End If
                
                
                    If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("andettimer")) = False Then
                        andetTimer = Trim(ds.Tables("tb_to_var").Rows(t).Item("andettimer"))
                    Else
                    andetTimer = "0"
                    End If
                
                    If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("clienttimer")) = False Then
                        clientTimer = Trim(ds.Tables("tb_to_var").Rows(t).Item("clienttimer"))
                    Else
                    clientTimer = "0"
                    End If
                
                
                If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("budgetkr")) = False Then
                    budgetkr = Trim(ds.Tables("tb_to_var").Rows(t).Item("budgetkr"))
                Else
                    budgetkr = 0
                End If
                
                    If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("budgettimer")) = False Then
                        budgetTimer = Trim(ds.Tables("tb_to_var").Rows(t).Item("budgettimer"))
                    Else
                        budgetTimer = 0
                    End If
                
                  
                
                    If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("valuta")) = False Then
                        valuta = Trim(ds.Tables("tb_to_var").Rows(t).Item("valuta"))
                    
                    
                        '**** Henter VALUTA '****
                        Dim strSQLvaluta As String = "SELECT id FROM valutaer WHERE valutakode = '" + valuta + "'"
                        objCmd = New OdbcCommand(strSQLvaluta, objConn2)
                        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                        intValuta = 0
                        If objDR2.Read() = True Then
                
                            intValuta = objDR2("id")
                
                        End If
                        objDR2.Close()
                
                        If intValuta <> 0 Then
                            intValuta = intValuta
                        Else
                            intValuta = 1
                        End If
                    
                    
                    Else
                        intValuta = intValuta 'valutaGen = 1
                    End If
                
                
                
                
                
                    If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("kundenr")) = False Then
                        kundenr = Trim(ds.Tables("tb_to_var").Rows(t).Item("kundenr"))
                    Else
                        kundenr = 0
                    End If
                
                    '******** SLUT EKSTRA FELTER
                
                
                
                
                    If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("jobnavn")) = False Then
                        jobnavn = Trim(ds.Tables("tb_to_var").Rows(t).Item("jobnavn"))
                    
                        jobnavn = Replace(jobnavn, "'", "")
                        jobnavn = Replace(jobnavn, "''", "")
                        jobnavn = Replace(jobnavn, ";", "")
                    
                        jobnavn = EncodeUTF8(jobnavn)
                        jobnavn = DecodeUTF8(jobnavn)
                        'jobnavn = jq_format(jobnavn)
                        'jobnavn = jq_formatTxt
                    
                    Else
                        jobnavn = "-TOM-"
                    End If
                
                    'jobnavn = EncodeUTF8(ds.Tables("tb_to_var").Rows(t).Item("jobnavn"))
                    'jobnavn = DecodeUTF8(jobnavn)
                    'jobnavn = jq_format(jobnavn)
                    'jobnavn = jq_formatTxt
                 
                    If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("jobnr")) = False Then
                        jobnr = Trim(ds.Tables("tb_to_var").Rows(t).Item("jobnr"))
                    Else
                        jobnr = 0
                    End If
             
                    jobnrTjk = jobnr
                
                    If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("jobstartdato")) = False Then
                        jobstartdato = ds.Tables("tb_to_var").Rows(t).Item("jobstartdato")
                    Else
                        jobstartdato = "2002-01-01"
                    End If
                
                    If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("jobslutdato")) = False Then
                        jobslutdato = ds.Tables("tb_to_var").Rows(t).Item("jobslutdato")
                    Else
                        jobslutdato = "2044-01-01"
                    End If
            
                
                    If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("jobansinit")) = False Then
                        jobansInit = ds.Tables("tb_to_var").Rows(t).Item("jobansinit")
                    Else
                        jobansInit = "XXX"
                    End If

                
                    If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("kundenavn")) = False Then
                        kundenavnTxt = ds.Tables("tb_to_var").Rows(t).Item("kundenavn")
                    
                        kundenavnTxt = Replace(kundenavnTxt, "'", "")
                        kundenavnTxt = Replace(kundenavnTxt, "''", "")
                        kundenavnTxt = Replace(kundenavnTxt, ";", "")
                    
                    Else
                        kundenavnTxt = "-TOM-"
                    End If
                
                    If String.IsNullOrEmpty(ds.Tables("tb_to_var").Rows(t).Item("kategori")) = False Then
                        projektKategori = ds.Tables("tb_to_var").Rows(t).Item("kategori")
                    Else
                        projektKategori = 0
                    End If
                    
              
                

            
                    '**** Henter Jobans '****
                
                    If jobansInit <> "" Then
                        Dim strSQLjobans As String = "SELECT mid FROM medarbejdere WHERE init = '" + jobansInit + "'"
                        objCmd = New OdbcCommand(strSQLjobans, objConn2)
                        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                        jobans = 0
                        If objDR2.Read() = True Then
                
                            jobans = objDR2("mid")
                
                        End If
                        objDR2.Close()
                    
                    End If
                
            
            
                    strjobnr = jobnr.ToString
            
           
                    '*** Tjekker om jobnr findes ***
                    Dim opdaterJob As Integer = 0
                    Dim strSQLjobnrfindes As String = "SELECT jobnr FROM job WHERE jobnr = '" + strjobnr + "'"
                    objCmd = New OdbcCommand(strSQLjobnrfindes, objConn2)
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
                        
                            Dim strSQKfindkunde As String = "SELECT kid FROM kunder WHERE kkundenavn = '" + kundenavnTxt + "'" 'SKAL ændres til kundernr når alle kunder er indlæst 27.6.2016
                            objCmd = New OdbcCommand(strSQKfindkunde, objConn2)
                            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                        
                            If objDR2.Read() = True Then
                
                                errThis = 0
                                kid = objDR2("kid")
                         
                            End If
                            objDR2.Close()
                        
                        
                        
                            If kid = 0 Then
                            
                                '*** OPRETTER KUNDE **
                                Dim strSQKkundeins As String = "INSERT INTO kunder SET kkundenavn = '" + kundenavnTxt + "', kkundenr = " + kundenr + ", ketype = 'ke', ktype = 0"
                                objCmd = New OdbcCommand(strSQKkundeins, objConn2)
                                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                objDR2.Close()
                            
                        
                        
                                '** HENTER ID PÅ NETOP oprettede ***'
                                Dim strSQKfindkundeNy As String = "SELECT kid FROM kunder WHERE kid <> 0 ORDER BY kid DESC LIMIT 1"
                                objCmd = New OdbcCommand(strSQKfindkundeNy, objConn2)
                                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                        
                                If objDR2.Read() = True Then
                
                          
                                    kid = objDR2("kid")
                         
                                End If
                                objDR2.Close()
                           
                         
                            Else
                            
                                '*** OPDATERER KUNDE **
                                Dim strSQKkundeupd As String = "UPDATE kunder SET kkundenavn = '" + kundenavnTxt + "', kkundenr = '" + kundenr + "' WHERE kid = " & kid
                                objCmd = New OdbcCommand(strSQKkundeupd, objConn2)
                                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                objDR2.Close()
                            
                            
                            End If
                        
                      
                    End Select
                    
                
                    kunderef = 0
                rekvisitionsnr = 0
                
                If budgetTimer <> 0 AND budgetkr <> 0 Then
                    jobtimepris = FormatNumber((budgetkr / budgetTimer), 2)
                    Else
                    jobtimepris = 0
                End If
                
                jobtimepris.ToString
                
                    budgetkr = Replace(budgetkr, ".", "")
                    budgetkr = Replace(budgetkr, ",", ".")
                    bruttooms = budgetkr
                
                    budgetTimer = Replace(budgetTimer, ".", "")
                    budgetTimer = Replace(budgetTimer, ",", ".")
                    internnote = ""
                
                    
                jobtimepris = Replace(jobtimepris, ".", "")
                jobtimepris = Replace(jobtimepris, ",", ".")
                    
                
               
                
                    'Return "Webservice Msg dt: " & jobstartdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture)
                
                    'Call DecodeUTF8(jobnavn)
                    'jobnavn = 
            
            
                    If CInt(errThis) = 0 Then
                
            
                        '*** Opret / Rediger job OK. Findes kunde *******************
                        
               
                    
                        'Return "Webservice Msg dt jobnrTjk: "+ jobnrTjk +" opretJobOk: "+ opretJobOk 
                        opretJobOk = 1
                        If CInt(opretJobOk) = 1 Then
            
                
                            If CInt(opdaterJob) = 1 Then 'opdater
                    
                           
                            Dim strSQLjobUpd As String = ("Update job SET jobnavn = '" & jobnavn & "', jobnr = '" & jobnr & "', jobstatus = 1, " _
                            & " jobstartdato = '" & jobstartdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "'," _
                            & " jobslutdato = '" & jobslutdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', editor = '" & editor & "', " _
                            & " dato = '" & dato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', beskrivelse = '" & beskrivelse & "', jobans1 = " & jobans & ", " _
                            & " kundekpers = " & kunderef & ", budgettimer = " & budgetTimer & ", jo_bruttooms = " & bruttooms & ", jo_gnstpris = " & jobtimepris & ", jo_gnsbelob = " & bruttooms & ", jobTpris =  "& bruttooms &", valuta = " & intValuta & ", salgsans1 = " & salgsansvarligMid & " WHERE jobnr = '" & jobnr & "'")
                  
                                objCmd = New OdbcCommand(strSQLjobUpd, objConn2)
                                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                objDR2.Close()
                            
                            
                    
                        
                                '*** Tilføjer til timereg_usejob (for en sikkerhedsskyld)
                        
                                '*** Finder jobid ***
                                jobID = 0
                        
                        
                        
                                '**** Sikrer timereg usejob bliver udfyldt ved rediger job ****
                                Dim strSQLlastJobID As String = "SELECT id FROM job WHERE jobnr = '" & jobnr & "'"
                                objCmd = New OdbcCommand(strSQLlastJobID, objConn2)
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
                                    objCmd = New OdbcCommand(strSQLtreguseFindes, objConn2)
                                    objDR3 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                
                            
                            
                                    timereguseJobFindes = 0
                                    If objDR3.Read() = True Then
                            
                                        timereguseJobFindes = 1
                            
                                    End If
                                    objDR3.Close()
                                
                                
                            
                                    If CInt(timereguseJobFindes) = 0 Then
                                    
                                        Dim strSQL3 As String = "INSERT INTO timereg_usejob (medarb, jobid, forvalgt, forvalgt_sortorder, forvalgt_af, forvalgt_dt) VALUES " _
                                        & " (" & objDR2("mid") & ", " & jobID & ", 0, 0, 0, '" & dato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "')"

                                        objCmd = New OdbcCommand(strSQL3, objConn2)
                                        objDR3 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                        objDR3.Close()
                                    
                                    End If
                            
                                
                                
                                End While
                                objDR2.Close()
                          
                            
                            
                            '*****************************************************************************************************************
                            '** FORACAST BLIVER IKKE OPDATERET DA DET KAN VÆRE ÆNDRET OG FORDELT I TO
                            '** UNDT hvis der ikke findes FORECAST i forvejen på dette id, så bliver der indlæst
                        
                            
                            '** FORACAST BLIVER INDLÆST på en medarbejder tilhørende afdeling og dermed bliver forecast fordelt på afdeling
                            '*****************************************************************************************************************    
                            
                            Dim strSQLakt As String = "SELECT id FROM aktiviteter WHERE job = " & jobID & " AND fakturerbar = 1 LIMIT 1"
                            objCmd = New OdbcCommand(strSQLakt, objConn2)
                            objDR3 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                
                            
                            
                            aktidPaaJob = 0
                            If objDR3.Read() = True Then
                            
                                aktidPaaJob = objDR3("id")
                            
                            End If
                            objDR3.Close()
                            
                            
                            Dim strSQLfc As String = "SELECT id FROM ressourcer_md WHERE jobid = " & jobID
                            objCmd = New OdbcCommand(strSQLfc, objConn2)
                            objDR3 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                
                            
                            
                            fcFindesPaaJob = 0
                            If objDR3.Read() = True Then
                            
                                fcFindesPaaJob = 1
                            
                            End If
                            objDR3.Close()
                            
                            
                            If aktidPaaJob <> 0 AND fcFindesPaaJob = 0 then
                            
                            'Client
                            'Research
                            'It
                            'FIELD
                            'Andet
                            
                            
                            clientTimer = Replace(clientTimer, ".", "")
                            clientTimer = Replace(clientTimer, ",", ".")
                            
                            researchTimer = Replace(researchTimer, ".", "")
                            researchTimer = Replace(researchTimer, ",", ".")
                            
                            itTimer = Replace(itTimer, ".", "")
                            itTimer = Replace(itTimer, ",", ".")
                            
                            fieldTimer = Replace(fieldTimer, ".", "")
                            fieldTimer = Replace(fieldTimer, ",", ".")
                            
                            andetTimer = Replace(andetTimer, ".", "")
                                andetTimer = Replace(andetTimer, ",", ".")
                                
                                Dim resAar As Integer = Year(jobstartdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture))
                            
                                If clientMid <> 0 And clientTimer <> "0" Then
                                
                                    
                                    Dim strSQLfc1 As String = "INSERT INTO ressourcer_md (jobid, aktid, medid, aar, md, timer) VALUES (" & jobID & ", " & aktidPaaJob & ", " & clientMid & ", " & resAar & ", 1, " & clientTimer & ")"
                                    objCmd2 = New OdbcCommand(strSQLfc1, objConn2)
                                    objDR6 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)
                                
                                End If
                            
                            
                                If researchMid <> 0 And researchTimer <> "0" Then
                            
                                    Dim strSQLfc2 As String = "INSERT INTO ressourcer_md (jobid, aktid, medid, aar, md, timer) VALUES (" & jobID & ", " & aktidPaaJob & ", " & researchMid & ", " & resAar & ", 1, " & researchTimer & ")"
                                    objCmd2 = New OdbcCommand(strSQLfc2, objConn2)
                                    objDR6 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)
                            
                                End If
                                
                            
                                If itMid <> 0 And itTimer <> "0" Then
                                
                                    Dim strSQLfc3 As String = "INSERT INTO ressourcer_md (jobid, aktid, medid, aar, md, timer) VALUES (" & jobID & ", " & aktidPaaJob & ", " & itMid & ", " & resAar & ", 1, " & itTimer & ")"
                                    objCmd2 = New OdbcCommand(strSQLfc3, objConn2)
                                    objDR6 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)
                                
                                End If
                            
                            
                                If fieldMid <> 0 And fieldTimer <> "0" Then
                                
                                    Dim strSQLfc4 As String = "INSERT INTO ressourcer_md (jobid, aktid, medid, aar, md, timer) VALUES (" & jobID & ", " & aktidPaaJob & ", " & fieldMid & ", " & resAar & ", 1, " & fieldTimer & ")"
                                    objCmd2 = New OdbcCommand(strSQLfc4, objConn2)
                                    objDR6 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)
                                
                                End If
                            
                            
                            If andetMid <> 0 And andetTimer <> "0" Then
                                Dim strSQLfc5 As String = "INSERT INTO ressourcer_md (jobid, aktid, medid, aar, md, timer) VALUES (" & jobID & ", " & aktidPaaJob & ", " & andetMid & ", " & resAar & ", 1, " & andetTimer & ")"
                                objCmd2 = New OdbcCommand(strSQLfc5, objConn2)
                                objDR6 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)
                                End If
                                
                                
                              End If 'aktidPaaJob
                            '**** END FORECAST ****'
                            
                            
                            
                            
                            
                            '************************
                            Else 'opret
                            '***********************
                    
                        
                
                            Dim strSQLjob As String = ("INSERT INTO job (jobnavn, jobnr, jobknr, jobTpris, jobstatus, jobstartdato," _
                            & " jobslutdato, editor, dato, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, " _
                            & " projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, " _
                            & " fakturerbart, budgettimer, fastpris, kundeok, beskrivelse, " _
                            & " ikkeBudgettimer, tilbudsnr, jobans1, " _
                            & " serviceaft, kundekpers, valuta, rekvnr, " _
                            & " risiko, job_internbesk, " _
                            & " jo_bruttooms, fomr_konto, salgsans1, jo_gnstpris, jo_gnsbelob) VALUES " _
                            & "('" & jobnavn & "', " _
                            & "'" & jobnr & "', " _
                            & "" & kid & ", " _
                            & "" & bruttooms & ", " _
                            & "1, " _
                            & "'" & jobstartdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', " _
                            & "'" & jobslutdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', " _
                            & "'" & editor & "', " _
                            & "'" & dato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "', " _
                            & "10, " _
                            & "1,1,1,1,1,1,1,1,1," _
                            & "1, " & budgetTimer & ",0,0," _
                            & "'" & beskrivelse & "', " _
                            & "0,0, " _
                            & "" & jobans & "," _
                            & "0," & kunderef & ", " _
                            & intValuta & ", '" & rekvisitionsnr & "', " _
                            & "100,'" & internnote & "'," _
                            & "" & bruttooms & ", 0, " & salgsansvarligMid & ", " & jobtimepris & ", " & bruttooms & ")")
    							
                                'return "Webservice SQL: " & strSQLjob
                          
                                objCmd = New OdbcCommand(strSQLjob, objConn2)
                                objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                objDR2.Close()
                                                     
                
                
                                '*** Finder jobid ***
                                Dim strSQLlastJobID As String = "SELECT id FROM job WHERE id <> 0 ORDER BY id DESC LIMIT 1"
                                objCmd = New OdbcCommand(strSQLlastJobID, objConn2)
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

                                            objCmd = New OdbcCommand(strSQL3, objConn2)
                                            objDR3 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                            objDR3.Close()
                                       
                                                
                                       
									
                                        End While
                                
                                        objDR2.Close()
                                   
                                
                                End Select ' lto
                        
                        
                
                                'Dim antalArbPakker As Array
                                'antalArbPakker = strantalArbPakker
                        
                                '** Vælg gruppe pbg. af projetktype
                                Dim agforvalgtStamgrpKri As String = " ag.forvalgt = 1 "
                            
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
                                          
                                    objCmd = New OdbcCommand(strSQLaktins, objConn2)
                                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                    objDR2.Close()
                                
                                
                    
                                    '*** Finder aktid ***
                                    Dim strSQLlastAktID As String = "SELECT id FROM aktiviteter WHERE id <> 0 ORDER BY id DESC"
                                    objCmd = New OdbcCommand(strSQLlastAktID, objConn2)
                                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    
                                    If objDR2.Read() = True Then
                
                                        lastAktID = objDR2("id")
                
                                    End If
                                    objDR2.Close()
                                
                    
                    
                                    '** FOMR REL ***********
                                    Dim strSQLaktfomr As String = ("SELECT for_fomr, for_aktid FROM fomr_rel WHERE for_aktid =  " & objDR3("stamaktid"))
                                             
                                    objCmd = New OdbcCommand(strSQLaktfomr, objConn2)
                                    objDR4 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                    While objDR4.Read() = True
                                
                                
                                
                                        Dim strSQLaktinsfomr As String = ("INSERT INTO fomr_rel (for_fomr, for_aktid, for_jobid, for_faktor) VALUES  (" & objDR4("for_fomr") & ", " & lastAktID & ", " & lastID & ", 100)")
                                               
                                        objCmd = New OdbcCommand(strSQLaktinsfomr, objConn2)
                                        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                        objDR2.Close()
                                
                                
                                    End While
                                    objDR4.Close()
                               
                            
                    
                                    '**************************************************************'
                                    '*** Hvis timepris ikke findes på job bruges Gen. timepris fra '
                                    '*** Fra medarbejdertype, og den oprettes på job              *'
                                    '*** BLIVER ALTID HENTET FRA Medarb.type                       *'
                                    '**** VALUTA ALTID TP 1
                                    '**************************************************************'
                                    tprisGen = 0
                                    valutaGen = 1
                                    kostpris = 0
                                    intTimepris = 0
                                    intValuta = valutaGen
                                
                               
                                    Dim SQLmedtpris As String = "SELECT medarbejdertype, timepris, timepris_a2, timepris_a3, tp0_valuta, kostpris, mnavn, mid FROM medarbejdere " _
                                    & " LEFT JOIN medarbejdertyper ON (medarbejdertyper.id = medarbejdertype) " _
                                    & " WHERE mid <> 0 AND mansat = 1 AND medarbejdertyper.id = medarbejdertype"

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
                           
                            
                                   
                                    
                                        End Select
                            
                            
                        
                        
                                        valutaGen = objDR2("tp0_valuta")
                        
                        
                                        '**** Indlæser timepris på aktiviteter ***'
                                        intTimepris = tprisGen
                                        intTimepris = Replace(intTimepris, ",", ".")
                            
                                        intValuta = valutaGen
                                        Dim strSQLtpris As String = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) " _
                                        & " VALUES (" & lastID & ", " & lastAktID & ", " & objDR2("mid") & ", " & timeprisalt & ", " & intTimepris & ", " & intValuta & ")"
							
                                        objCmd2 = New OdbcCommand(strSQLtpris, objConn2)
                                        objDR6 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)
                                    
                        
		
                                    End While

                                    objDR2.Close()
                              
                                    '** Slut timepris **
                
                
                   
							
                    
                    
                    
                                End While
                                objDR3.Close()
                          
                            
                
                            '*****************************************************************************************************************
                            '** FORACAST BLIVER INDLÆST på en medarbejder tilhørende afdeling og dermed bliver forecast fordelt på afdeling
                            '*****************************************************************************************************************    
                            'Client
                                'Research
                                'It
                                'FIELD
                                'Andet
                                clientTimer = Replace(clientTimer, ".", "")
                                clientTimer = Replace(clientTimer, ",", ".")
                            
                                researchTimer = Replace(researchTimer, ".", "")
                                researchTimer = Replace(researchTimer, ",", ".")
                            
                                itTimer = Replace(itTimer, ".", "")
                                itTimer = Replace(itTimer, ",", ".")
                            
                                fieldTimer = Replace(fieldTimer, ".", "")
                                fieldTimer = Replace(fieldTimer, ",", ".")
                            
                                andetTimer = Replace(andetTimer, ".", "")
                                andetTimer = Replace(andetTimer, ",", ".")
                            
                            Dim resAar As Integer = Year(jobstartdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture))
                                
                            
                            If clientMid <> 0 And clientTimer <> "0" Then
                                
                                Dim strSQLfc1 As String = "INSERT INTO ressourcer_md (jobid, aktid, medid, aar, md, timer) VALUES (" & lastID & ", " & lastAktID & ", " & clientMid & ", " & resAar & ", 1, " & clientTimer & ")"
                                objCmd2 = New OdbcCommand(strSQLfc1, objConn2)
                                objDR6 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)
                                
                            End If
                            
                            
                            If researchMid <> 0 And researchTimer <> "0" Then
                            
                                Dim strSQLfc2 As String = "INSERT INTO ressourcer_md (jobid, aktid, medid, aar, md, timer) VALUES (" & lastID & ", " & lastAktID & ", " & researchMid & ", " & resAar & ", 1, " & researchTimer & ")"
                                objCmd2 = New OdbcCommand(strSQLfc2, objConn2)
                                objDR6 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)
                            
                            End If
                                
                            
                            If itMid <> 0 And itTimer <> "0" Then
                                
                                Dim strSQLfc3 As String = "INSERT INTO ressourcer_md (jobid, aktid, medid, aar, md, timer) VALUES (" & lastID & ", " & lastAktID & ", " & itMid & ", " & resAar & ", 1, " & itTimer & ")"
                                objCmd2 = New OdbcCommand(strSQLfc3, objConn2)
                                objDR6 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)
                                
                            End If
                            
                            
                            If fieldMid <> 0 And fieldTimer <> "0" Then
                                
                                Dim strSQLfc4 As String = "INSERT INTO ressourcer_md (jobid, aktid, medid, aar, md, timer) VALUES (" & lastID & ", " & lastAktID & ", " & fieldMid & ", " & resAar & ", 1, " & fieldTimer & ")"
                                objCmd2 = New OdbcCommand(strSQLfc4, objConn2)
                                objDR6 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)
                                
                            End If
                            
                            
                            If andetMid <> 0 And andetTimer <> "0" Then
                                Dim strSQLfc5 As String = "INSERT INTO ressourcer_md (jobid, aktid, medid, aar, md, timer) VALUES (" & lastID & ", " & lastAktID & ", " & andetMid & ", " & resAar & ", 1, " & andetTimer & ")"
                                objCmd2 = New OdbcCommand(strSQLfc5, objConn2)
                                objDR6 = objCmd2.ExecuteReader '(CommandBehavior.closeConnection)
                            End If
                            '**** END FORECAST ****'
                
                                                
                            End If 'Opdater / opret
                
                        End If 'opretJobOk 
                    
                    End If 'CInt(errThis) = 0
            
            
                    objConn2.Close()
                    objConn3.Close()
                
                
          
                    t = t + 1
                    'Catch ex As Exception
                    '   MsgBox("Can't load Web page" & vbCrLf & ex.Message)
                    'End Try
        
              
                    'end if
                    'Select Case Right(t, 1)
                    '   Case 0
                    'Response.flush()
                    'End Select
                
                    
                    
            Next
            
        Next
        
       
                
                                                    
        Dim strSQLjobnrUpd As String = "UPDATE licens SET jobnr = " & jobnr & " WHERE id = 1"
        objCmd = New OdbcCommand(strSQLjobnrUpd, objConn)
        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
        objDR2.Close()
        
        objConn.Close()
       
        
        'objConn.Dispose()
        'objConn2.Dispose()
        'objConn3.Dispose()
                                           
        
        
    End Function

    
    
    'Public Function rtErrTxt(ByVal erTxt As String) As String
    ' Dim eTxt As String
    '    eTxt = erTxt
        
    '   Return eTxt
    'End Function

    Private Function Response() As Object
        Throw New NotImplementedException
    End Function

   
    
End Class

        




