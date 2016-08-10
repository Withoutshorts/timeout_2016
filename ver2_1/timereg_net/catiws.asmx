﻿


<%@ WebService language="VB" class="CATI" %>
Imports System
Imports System.Web.Services
Imports System.Data
Imports System.Data.Odbc





Public Class CATI : 

    public testValue As string 
    Public testmode As Integer = 0
    
    Public timerkom As string
    Public intMedarbId As String
    Public intCatiId, intJobNr As String
    Public dlbTimer As String
   
    Public cdDato As String
    public importFrom AS String
    
    Public errThis As Integer = 0
        
    Public meNavn As String
    Public meNr As String
    Public meID As Integer
    Public mTypeSQL As String
    
    Public jobNavn As String
    Public jobId As Double
    Public jobFastPris As Integer
    Public jobSeraft As Double
    Public jobKid As Double
    Public jobValuta As Integer
    
    Public aktNavn, aktnavnUse As String
    Public aktId As Double = 0
    Public aktFakturerbar As Integer
    
    Public kurs As String
    
    Public kNavn As String
    Public kNr As Double
    
    Public extsysid As Double
    
    public intJobNrTxt As String

    Public intTimepris As String = 0
    Public timeprisAlt As Double = 0
    Public intValuta As Double
    Public tpid As Double
    Public foundone As String = "n"
    Public timeprisalernativ, valutaAlt As String
    Public mtp As Integer
    Public aktIdUse As Double
    
    Public kostpris As String
    Public tprisGen As String
    Public valutaGen As Integer = 1
    
    Public emx As String = 100
    Public tempM As String 
    
    Public errThisAll As String = "ErrorMSG: "
    Public len_tempM As String
    Public akttypeUse As String
    
    Public addBonus10 As Integer = 0
    
    'Public errThisAll As Integer = 0
    
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
    
    
    Public Weekdaynb As integer
    Public Function Weekday(ByVal DateValue As DateTime) As Integer
        
        
      
        Weekdaynb = DateValue.DayOfWeek
        Return Weekdaynb
        
        
    
    End Function
   

   
   
    
    <WebMethod()> Public Function upDateCatiTimer(ByVal ds As DataSet) As String
        
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
        Dim dt As DataTable
        Dim dr As DataRow
        
        Dim strConn As String
        

        'return "ok"
        
        'Dim objTable As DataTable
        
        Dim t As Double = 0
        
        'Try
       
        
        For Each dt In ds.Tables
            For Each dr In dt.Rows
                
              
                errThis = 0
                meID = 0
                meNavn = "-"
                meNr = "0"
                extsysid = 0
                aktId = 0
                intJobNr = "0"
                'intjobId = 0
                intValuta = 0
                intMedarbId = ""
                jobKid = 0
                jobId = 0
                intJobNrTxt = ""
                foundone = "n"
                
                kNr = 0
                kNavn = ""
                
                intTimepris = 0
                timeprisAlt = 0
                kurs = 100
                kostpris = 0
                mTypeSQL = ""
                addBonus10 = 0
                
                 
                
                If ds.Tables("CATI_TIME").Columns.Contains("origin") Then
                    importFrom = ds.Tables("CATI_TIME").Rows(t).Item("origin")
                Else
                    importFrom = 0
                End If

                If t = 0 Then
                    
                    If importFrom = "1" Then 'Test: Vietnam
                        strConn = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_epi_catitest;User=to_outzource2;Password=SKba200473;"
                    Else 'Prod: CATI
                        If importFrom = "3" Then 'Test: IBM US
                            strConn = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_epi;User=to_outzource2;Password=SKba200473;"
                            
                        Else
                            'strConn = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_epi_catitest;User=to_outzource2;Password=SKba200473;"
                            strConn = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_epi;User=outzource;Password=SKba200473;"
                        End If
                        
                    End If
            
                    'Else 'Prod: CATI
                    '**** TESTER *****'
                    'strConn = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_epi_catitest;User=to_outzource2;Password=SKba200473;"
                    'End If
                    
                    '** Åbner Connection ***'
                    objConn = New OdbcConnection(strConn)
                    objConn.Open()
                    
                End If

                Select Case importFrom
                    Case "1"
                        'Vietnam
                        aktnavnUse = "Epinion Asia (Vietnam) time"
                        akttypeUse = 1
                        mTypeSQL = " AND (medarbejdertype = 18 OR medarbejdertype = 20 OR medarbejdertype = 21 OR medarbejdertype = 22)"
                        timerkom = ""
                    Case Else
                        'CATI / SPSS
                        aktnavnUse = "CATI interviewtimer import"
                        akttypeUse = 91
                        timerkom = ""
                        mTypeSQL = " AND medarbejdertype = 14"
                End Select
                
                'intCatiId = 999
                
                'If String.IsNullOrEmpty(ds.Tables("CATI_TIME").Rows(t).Item("id")) = True Then
                intCatiId = ds.Tables("CATI_TIME").Rows(t).Item("id")
                'Else
                'intCatiId = 0
                'End If
                
                    
                'If String.IsNullOrEmpty(ds.Tables("CATI_TIME").Rows(t).Item("medarbejderid")) = False Then
                'tempM = ds.Tables("CATI_TIME").Rows(t).Item("medarbejderid")
                'tempM = tempM.ToString()
                'Else
                'tempM = ""
                'End If
                    
                tempM = IsDBNull(ds.Tables("CATI_TIME").Rows(t).Item("medarbejderid"))
                If tempM <> True Then
                    tempM = ds.Tables("CATI_TIME").Rows(t).Item("medarbejderid")
                    tempM = tempM.ToString()
                Else
                    tempM = ""
                End If
                
                intMedarbId = tempM
                
               
                
                'intJobNr = 14022
                'If tempM <> "CLOUD\iisanon01" Then                
                
                If String.IsNullOrEmpty(ds.Tables("CATI_TIME").Rows(t).Item("jobid")) = False Then
                    intJobNr = ds.Tables("CATI_TIME").Rows(t).Item("jobid").ToString
                    
                Else
                    intJobNr = "0"
                End If
                
                If Len(Trim(intJobNr)) <> 0 Then
                    intJobNr = intJobNr
                Else
                    intJobNr = "0"
                End If
                
                'Else
                'intJobNr = "0"
                'End If
                    
                ' **** Import af valgkamp 2011 fra special server **''
                If importFrom = "0" Then
                    If intJobNr = "121212" Then
                        intJobNr = "14520" '** valgkamp
                    End If
                End If
                
                If importFrom = "3" Then
                    If intJobNr = "TU2014" Then
                        intJobNr = "17687"
                    End If
                   
                End If
                
                intJobNrTxt = intJobNr
                
             

                'dlbTimer = 1
                'If String.IsNullOrEmpty(ds.Tables("CATI_TIME").Rows(t).Item("timer")) = True Then
                dlbTimer = ds.Tables("CATI_TIME").Rows(t).Item("timer")
                'Else
                'dlbTimer = 0
                'End If
                
                
                'cdDato = "2010-12-01"
                'If String.IsNullOrEmpty(ds.Tables("CATI_TIME").Rows(t).Item("dato").ToString()) = True Then
                cdDato = ds.Tables("CATI_TIME").Rows(t).Item("dato")
                'Else
                'cdDato = "2001-10-01"
                'End If
                
               
                    
                '*** Test indlæser ellae records i err db ***'
               
                'Dim strSQLer As String = "INSERT INTO timer_imp_err (dato, extsysid, errid, jobid, jobnr, med_init, timeregdato, timer, origin) " _
                '& " VALUES " _
                '& " ('" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "', 0, 99, 0,'" & intJobNrTxt & " t:" & dlbTimer & " dt: " & Year(cdDato) & "/" & Month(cdDato) & "/" & Day(cdDato) & " sysID: " & intCatiId & " medid: " & intMedarbId & "', '0', '2000-01-01', 0, 1)"

                
                'Dim strSQLerrTest As String = "INSERT INTO timer_imp_err (dato, jobnr, timer, med_init) VALUES ('" & cdDato & "', '" & intJobNr & "', '" & dlbTimer & "', '" & intMedarbId & " : "& dlbTimer &"')"

           
                'objCmd = New OdbcCommand(strSQLerrTest, objConn)
                'objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                'objDR.Close()
               
                'errThis = 6
                '*** end **'
                 
                
                '*** if record findes i timereg i forvejen ****'
                
                If CInt(errThis) = 0 Then
                    
                    '*** tjekker om extsysid findes i forvejen ***'
                    Dim strSQLext As String = "SELECT extsysid FROM Timer WHERE extsysid = '" & intCatiId & "' AND origin = " & importFrom
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
                
                
                '*** if record findes i timereg i forvejen ****'
                If CInt(errThis) <> 6 Then
                
                       
                
                    '*** Er timer angive korrekt ***'
                    If Len(Trim(dlbTimer)) = 0 Or InStr(dlbTimer, "-") <> 0 Then
                    
                        dlbTimer = 0
                        errThis = 7
                    
                    End If
       
                    '*** Finder medarbejder oplysninger ***'
                    'If  String.IsNullOrEmpty(intMedarbId) <> true Then
                    If intMedarbId <> "" Then
                        intMedarbId = intMedarbId
                    Else
                        intMedarbId = ""
                        errThis = 1
                    End If
        
                
       
                    If CInt(errThis) <> 1 Then
            
        
                        Dim strSQLme As String = "SELECT mid, mnavn, mnr FROM medarbejdere WHERE init = '" & intMedarbId & "'" & mTypeSQL 'AND mansat <> 2" & mTypeSQL
                        
                        objCmd = New OdbcCommand(strSQLme, objConn)
                        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                        If objDR.Read() = True Then

                            meNavn = objDR("mnavn")
                            If Len(Trim(meNavn)) <> 0 Then
                                meNavn = DecodeUTF8(meNavn)
                            End If
                        
                            meNr = objDR("mnr")
                            meID = objDR("mid")

                        End If
            
                        objDR.Close()
                    
                            
                        '**** Jobn Doe befoere 2011 **'
                        If meID = 0 And importFrom = "1" And CDate(cdDato) < CDate("01-01-2011") Then
                            
                            meNavn = "John Doe Asia"
                            meNr = "1000006"
                            meID = 604
                            
                        End If
                 
                    
            
                    End If
                    '***'
                
                
                
                
                
                    
                
                    If meID <> 0 Then
                        meID = meID
                    Else
                        meID = 0
                        errThis = 11
                    End If
                        
                    
                        
      
                    '** Trimmer jobnr **'
                    If Len(Trim(intJobNr)) <> 0 Then
                        
                 
                        Select Case importFrom
                            Case "1"
                                intJobNr = Replace(intJobNr, "_", "")
                            Case "3"
                                
                                
                                If (intJobNr.Substring(0, 2)) = "P_" Then
                                    
                                    intJobNr = Replace(intJobNr, "P_", "")
                                    intJobNr = Left(intJobNr, 5)
                                    
                                    
                                Else
                                
                                    If (intJobNr.Substring(0, 1)) = "P" Then
                                        intJobNr = Replace(intJobNr, "P", "")
                                        intJobNr = Left(intJobNr, 5)
                                  
                                    Else
                                    
                                        intJobNr = intJobNr
                                     
                                    End If
                                End If
                                    
                                    
                                 
                                
                            Case Else
                                intJobNr = intJobNr
                        End Select
                        
                 
                    
                    Else
                        intJobNr = "0"
                        errThis = 2
                    End If
                
                    
                
                    If CInt(errThis) = 0 Then
                    
                    
            

                        Dim strSQLjob As String = "SELECT id, jobnavn, fastpris, serviceaft, jobknr, valuta FROM job WHERE jobnr = '" & intJobNr & "'"
                        objCmd = New OdbcCommand(strSQLjob, objConn)
                        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                        If objDR.Read() = True Then

                            jobNavn = Replace(objDR("jobnavn"), "'", "''")
                            jobNavn = EncodeUTF8(jobNavn)
                            jobNavn = DecodeUTF8(jobNavn)
                            jobId = objDR("id")
                            jobFastPris = objDR("fastpris")
                            jobSeraft = objDR("serviceaft")
                            jobKid = objDR("jobknr")
                            'jobValuta = objDR("valuta")
            
                    

                        End If
            
                        objDR.Close()
                        '***'
                    
                
                
                
                        If CInt(jobId) = 0 Then
                                
                            'Hvis Vietnam og der ikke er fundet et job indlæs på Epinion Asian --> Lokale projekter
                            If importFrom = "1" Then
                                jobNavn = "Diverse Vietnam projekter (lokale/herreløse)"
                                jobId = 1681
                                jobFastPris = 0
                                jobSeraft = 0
                                jobKid = 1048
                                errThis = 0
                                timerkom = intJobNr
                                timerkom = EncodeUTF8(timerkom)
                                timerkom = DecodeUTF8(timerkom)
                                intJobNr = "100" '14554
                                
                            Else
                                jobNavn = ""
                                jobId = 0
                                jobFastPris = 0
                                jobSeraft = 0
                                jobKid = 0
                                errThis = 21
                               
                            End If
                               
                    
                        End If
                
                    

                    End If

                
                
                    If CInt(errThis) = 0 Then
                    
                        '*** Finder akt. oplysninger ****'
                        '*** Fase ??? ***'
                        Dim strSQLa As String
                        
                        Select Case importFrom
                           
                            Case "1" '*** Vietnam **'
                                strSQLa = "SELECT id, fakturerbar, navn FROM aktiviteter WHERE job = " & jobId & " AND navn = '" & aktnavnUse & "'" ' AND fakturerbar = " & akttypeUse & "" 'Interviewer 'AND aktstatus = 1
                            Case Else
                                strSQLa = "SELECT id, fakturerbar, navn FROM aktiviteter WHERE job = " & jobId & " AND navn = '" & aktnavnUse & "' AND fakturerbar = " & akttypeUse & "" 'Interviewer 'AND aktstatus = 1
                         
                        End Select
                        
                        objCmd = New OdbcCommand(strSQLa, objConn)
                        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                        If objDR.Read() = True Then

                            aktNavn = objDR("navn")
                            aktNavn = EncodeUTF8(aktNavn)
                            aktNavn = DecodeUTF8(aktNavn)
                            aktId = objDR("id")
                            aktFakturerbar = objDR("fakturerbar")
               

                        End If
            
                        objDR.Close()
        
                        '** Blev der fundet en aktivitet
                        If aktId <> 0 Then
                            aktId = aktId
                        Else
                            aktId = 0
                            errThis = 3
                        End If
        
                        '***'
                    
                  
                    
                    End If
        
               
        
        
                    If CInt(errThis) = 0 Then
        
                        '*** Finder medarb timepris og kostpris ***'
                        '*** Først prøves aktivitet derefter job ***'
                
                
                        For mtp = 0 To 1
                    
                            If foundone = "n" Then
                        
                    
                    
                                If mtp = 0 Then
                                    aktIdUse = aktId
                                Else
                                    aktIdUse = 0
                                End If
                            
                                'Dim strSQLer As String = "INSERT INTO timer_imp_err (dato, extsysid, errid, jobid, jobnr, med_init, timeregdato, timer, origin) " _
                                '          & " VALUES " _
                                '           & " ('" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "', '" & intCatiId & "', "& errThis &", 0,'" & intJobNrTxt & "_" & intJobNr & "', '" & intMedarbId & "_" & tempM & "_" & meID & "', '" & Year(cdDato) & "/" & Month(cdDato) & "/" & Day(cdDato) & "', 5, " & importFrom & ")"

           
                                'objCmd = New OdbcCommand(strSQLer, objConn)
                                'objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                'objDR.Close()
                
                                Dim strSQLmtp As String = "SELECT id AS tpid, jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta FROM timepriser WHERE jobid = " & jobId & " AND aktid = " & aktIdUse & " AND medarbid =  " & meID
								
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
                    
                        
                                 
                                        If Len(Trim(timeprisalernativ)) <> 0 And Len(Trim(valutaAlt)) <> 0 Then
                        
                                            Dim strSQL3 As String = "SELECT mid, " & timeprisalernativ & " AS useTimepris, " & valutaAlt & " AS useValuta, " _
                                            & " medarbejdertype FROM medarbejdere, medarbejdertyper WHERE mid =" & meID & " AND medarbejdertyper.id = medarbejdertype"
									
                                            objCmd = New OdbcCommand(strSQL3, objConn)
                                            objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                                            If objDR2.Read() = True Then
                                                intTimepris = objDR2("useTimepris")
                                                intValuta = objDR2("useValuta")
                                            End If
                                            objDR2.Close()
									
                                        
                                            If Len(Trim(intTimepris)) <> 0 Then
                                                intTimepris = Replace(intTimepris, ",", ".")
									
                                            End If
									
                                    
                                        
                                        End If
                                        
                
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
                        tprisGen = 0
                        valutaGen = 1
                    
                        Dim SQLmedtpris As String = "SELECT medarbejdertype, timepris, tp0_valuta, kostpris, mnavn FROM medarbejdere, medarbejdertyper " _
                        & " WHERE Mid = " & meID & " AND medarbejdertyper.id = medarbejdertype"

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
                            & " VALUES (" & jobId & ", 0, " & meID & ", 0, " & intTimepris & ", " & intValuta & ")"
							
                            objCmd = New OdbcCommand(strSQLtpris, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
							
                        End If
                        
                        
                        '**** Tjekker om akt. er sat til fast timepris for alle medarbjedere ***'
                        '**** overruler medarbejer timepris på akt. og job *********************'
                        
                        
                        Dim brug_fasttp As Integer = 0
                        Dim brug_fastkp As Integer = 0
                        Dim fasttp As Double = 0
                        Dim fastkp As Double = 0
                        Dim fasttp_val As Integer = 0
                        '*** Tjekker om aktiviteten er sat til ens timpris for alle medarbejdere (overskriver medarbejderens egen timepris)
                        Dim strSQLtjkAktTp As String = "SELECT brug_fasttp, brug_fastkp, fasttp, fasttp_val, fastkp, fastkp_val FROM aktiviteter WHERE id = " & aktIdUse
                        objCmd = New OdbcCommand(strSQLtjkAktTp, objConn)
                        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                        If objDR.Read() = True Then

                            If CInt(objDR("brug_fasttp")) = 1 Then
                                brug_fasttp = 1
                                fasttp = objDR("fasttp")
                                fasttp = Replace(fasttp, ".", "")
                                fasttp = Replace(fasttp, ",", ".")

                                fasttp_val = objDR("fasttp_val")
                            End If


                            If CInt(objDR("brug_fastkp")) = 1 Then
                                brug_fastkp = 1
                                fastkp = objDR("fastkp")
                                fastkp = Replace(fastkp, ".", "")
                                fastkp = Replace(fastkp, ",", ".")
                            End If

                        End If
                        objDR.Close()


                       
                        If CInt(brug_fasttp) = 1 Then
                            intTimepris = fasttp
                            intValuta = fasttp_val
                        End If

                        If CInt(brug_fastkp) = 1 Then
                            kostpris = fastkp
                        End If
                            
                           
                        
                
                
                    End If
                
                
                   
							            
                       
                        
                    If CInt(errThis) = 0 Then
                    
                        '** Finder valuta og kurs **'
                        If Len(Trim(intValuta)) <> 0 And intValuta <> 0 Then
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
                        
                    End If
                    '***'		            
					
				
              
                
                
                
        
                
                
                    If CInt(errThis) = 0 Then
                    
                        '*** Finder kunde oplysninger ****'
                        If Len(Trim(jobKid)) <> 0 And jobKid <> 0 Then
                            jobKid = jobKid
                        Else
                            jobKid = 0
                            errThis = 5
                        End If
      
                        Dim strSQLk As String = "SELECT kkundenavn, kkundenr FROM kunder WHERE kid = " & jobKid
                        objCmd = New OdbcCommand(strSQLk, objConn)
                        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                        If objDR.Read() = True Then

                            kNavn = Replace(objDR("kkundenavn"), "'", "")
                            kNavn = EncodeUTF8(kNavn)
                            kNavn = DecodeUTF8(kNavn)
                            kNr = jobKid 'objDR("kkundenr")
                            
                            If jobKid = 1012 Then 'DK Statistik
                                addBonus10 = 1
                                
                            End If
            
                            objDR.Close()
                            '***'
                
                            If Len(Trim(kNavn)) = 0 Then
                                errThis = 51
                            End If
                    
                        End If
                
                
                    End If
                    
                
                    
                

                    'errThis = 7
                    'dlbTimer = Replace(dlbTimer, ".", ",")
                    'dlbTimer = FormatNumber(dlbTimer, 2)
                    'dlbTimer = Replace(dlbTimer, ",", ".")
                    dlbTimer = FormatNumber(dlbTimer, 2)
                    dlbTimer = Replace(dlbTimer, ",", ".")
                    If CInt(errThis) = 0 Then
                        
                        'If importFrom = "1" Then 'Test: Vietnam
                        'dlbTimer = FormatNumber(dlbTimer, 2)
                        'dlbTimer = Replace(dlbTimer, ",", ".")
                        'Else
                        'dlbTimer = dlbTimer
                        
                       
                        'End If
                                                        
                    
                        'Call jq_format(meNavn)
                        'meNavn = jq_formatTxt
                    
                        'Call jq_format(kNavn)
                        'kNavn = jq_formatTxt
                    
                        'Call jq_format(jobNavn)
                        'jobNavn = jq_formatTxt
                    
                        'Call DecodeUTF8(jobNavn)
                    
                        'testValue = testValue &"<br>"& meNavn &" "& jobnavn &" "& dlbTimer & "" & cdDato
                          
                        '*** Udelader Supervisorer ***'
                        Select Case intMedarbId
                            Case "intw_244", "intw_480", "intw_465", "intw_1014", "intw_305", "intw_768", "Gunzo", "itw_1472", "itw_1475", "intw_753"

                            Case Else
                                
                                '**** Udelader jobnr ***'
                                Select Case intJobNr
                                    Case "9010"
                                    Case Else
                                
                    
                                        '*** Indlæser Timer ***'
                                        
                                        
                                        Dim strSQL As String = "INSERT INTO timer " _
                                        & "(" _
                                        & " timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, timerkom, " _
                                        & " TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, " _
                                        & " editor, kostpris, seraft, " _
                                        & " valuta, kurs, extSysId, sttid, sltid, origin" _
                                        & ") " _
                                        & " VALUES " _
                                        & " (" _
                                        & dlbTimer & ", " & aktFakturerbar & ", " _
                                        & "'" & Year(cdDato) & "/" & Month(cdDato) & "/" & Day(cdDato) & "', " _
                                        & "'" & meNavn & "', " _
                                        & meID & ", " _
                                        & "'" & jobNavn & "', " _
                                        & "'" & intJobNr & "', " _
                                        & "'" & kNavn & "'," _
                                        & kNr & ", " _
                                        & "'" & timerkom & "', " _
                                        & aktId & ", " _
                                        & "'" & aktNavn & "', " _
                                        & Year(Now) & ", " _
                                        & intTimepris & ", " _
                                        & "'" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "', " _
                                        & jobFastPris & ", " _
                                        & "'00:00:01', " _
                                        & "'CaTi Webs Import', " _
                                        & Replace(kostpris, ",", ".") & ", " _
                                        & jobSeraft & ", " _
                                        & intValuta & ", " _
                                        & Replace(kurs, ",", ".") & ", '" & intCatiId & "', '00:00:00', '00:00:00', " & importFrom & ")"

           
                                        '** Manger salgs og kost priser ***'

                                        'Response.write(strSQL)
        
                                        objCmd = New OdbcCommand(strSQL, objConn)
                                        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                        'objDR.Close() FORBUDT ELLLERS INDLÆSES BONUS IKKE !!!
                                            
                                            
                                            
                                        '*** tilføj Bonus 10 i Weekend *****
                                        If CInt(addBonus10) = 1 Then
                                                
                                               
                                               
                                                    
                                            '** Hvis søndag ell. lørdag
                                            Dim DateValue As DateTime = Year(cdDato) & "/" & Month(cdDato) & "/" & Day(cdDato)
                                                
                                            Call Weekday(DateValue)
                                                
                                            If CInt(Weekdaynb) = 0 Then 'Or CInt(Weekdaynb) = 6
                                                    
                                                ''** Finder Bonus10 Aktivitet på job - Bonus 10 CATIindlaest **'
                                                Dim strSQLaktbo As String = "SELECT id, navn FROM aktiviteter WHERE fakturerbar = 54 AND navn LIKE 'Bonus 10 CATIindlaest%' AND job = " & jobId
                                                Dim baktnavn As String = ""
                                                Dim baktid As Double = 0
                                                    
                                                objCmd = New OdbcCommand(strSQLaktbo, objConn)
                                                objDR3 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                                                If objDR3.Read() = True Then

                                                     
                                                            
                                                    baktnavn = objDR3("navn")
                                                    baktnavn = EncodeUTF8(baktnavn)
                                                    baktnavn = DecodeUTF8(baktnavn)
                                                    baktid = objDR3("id")
                                                       
                                                End If
            
                                                objDR3.Close()
        
                                                '** Blev der fundet en aktivitet
                                                If baktid <> 0 Then
                                                    baktid = baktid
                                                Else
                                                    baktid = 0
                                                    errThis = 13
                                                End If
                                                    
                                                
                                                
                                                
                                                
                                                
                                                '*** Indlæser Timer ***'
                                                If CInt(errThis) = 0 Then
                                                    
                                                    
                                                    
                                                    Dim strSQLbonus As String = "INSERT INTO timer " _
                                                    & "(" _
                                                    & " timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, timerkom, " _
                                                    & " TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, " _
                                                    & " editor, kostpris, seraft, " _
                                                    & " valuta, kurs, extSysId, sttid, sltid, origin" _
                                                    & ") " _
                                                    & " VALUES " _
                                                    & " (" & dlbTimer & ", 54, " _
                                                    & "'" & Year(cdDato) & "/" & Month(cdDato) & "/" & Day(cdDato) & "', " _
                                                    & "'" & meNavn & "', " _
                                                    & meID & ", " _
                                                    & "'" & jobNavn & "', " _
                                                    & "'" & intJobNr & "', " _
                                                    & "'" & kNavn & "'," _
                                                    & kNr & ", " _
                                                    & "'', " _
                                                    & baktid & ", " _
                                                    & "'" & baktnavn & "', " _
                                                    & Year(Now) & ", 0, " _
                                                    & "'" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "', " _
                                                    & jobFastPris & ", " _
                                                    & "'00:00:01', " _
                                                    & "'CaTi Webs Import', 0, " _
                                                    & jobSeraft & ", " _
                                                    & intValuta & ", " _
                                                    & Replace(kurs, ",", ".") & ", '" & intCatiId & "', '00:00:00', '00:00:00', " & importFrom & ")"

           
                                                   

                                                    'Response.write(strSQL)
        
                                                    objCmd = New OdbcCommand(strSQLbonus, objConn)
                                                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                                    'objDR.Close()
                                                    
                                                    
                                                    
                                                  
                                                        
                                                Else 'Hvis bonus10 akt ikke findes på job
                                                    
                                                    Dim strSQLer As String = "INSERT INTO timer_imp_err (dato, extsysid, errid, jobid, jobnr, med_init, timeregdato, timer, origin) " _
                                                        & " VALUES " _
                                                        & " ('" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "', '" & intCatiId & "', " & errThis & ", " & jobId & ",'" & intJobNr & "', '" & intMedarbId & "', '" & Year(cdDato) & "/" & Month(cdDato) & "/" & Day(cdDato) & "', " & dlbTimer & ", " & importFrom & ")"

           
                                                    objCmd = New OdbcCommand(strSQLer, objConn)
                                                    objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                                    'objDR.Close()
                                                
                                                End If
                                                        
                                                        
                                            End If 'err
                                            
                                            
                                                
                                            
                                                
                                        End If 'bnonus10 = 1
                                            
                                            
                    
                                        Dim strSQL2 As String = "DELETE FROM timer_imp_err WHERE extsysid = '" & intCatiId & "' AND origin = " & importFrom
                    
                                        objCmd = New OdbcCommand(strSQL2, objConn)
                                        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                        objDR.Close()
                                        
                                End Select
                    
                        End Select
                    
                        'errThisAll = errThisAll
                       

                    
                    Else '** Indlæser ind til ErrLog
                        
                        
                        '*** Udelader Supervisorer ***'
                        Select Case intMedarbId
                            Case "intw_244", "intw_480", "intw_465", "intw_1014", "intw_305", "intw_768", "Gunzo"

                            Case Else
                                
                                '**** Udelader jobnr ***'
                                Select Case intJobNr
                                    Case "9010"
                                    Case Else
                                
                                        Dim strSQLer As String = "INSERT INTO timer_imp_err (dato, extsysid, errid, jobid, jobnr, med_init, timeregdato, timer, origin) " _
                                        & " VALUES " _
                                        & " ('" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "', '" & intCatiId & "', " & errThis & ", " & jobId & ",'" & intJobNr & "', '" & intMedarbId & "', '" & Year(cdDato) & "/" & Month(cdDato) & "/" & Day(cdDato) & "', " & dlbTimer & ", " & importFrom & ")"

           
                                        objCmd = New OdbcCommand(strSQLer, objConn)
                                        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                        objDR.Close()
            
                                        'Response.write(strSQLer)
                        
                                        'errThisAll = errThisAll & ";" & errThis & "(" & intCatiId & ")"
                    
                                        'errThisAll = errThisAll + 1    
                                        
                                End Select
                                
                        End Select
            
                    End If
        
                    
                
                
                    
                
                    
                    'End Select '** udelad supervisor

                    'Return errThis & ", "
                    
                    'End If 'testmode
                    'Else
                    
                    
                    '*** Update function ***'
                    'If errThis = 6 And importFrom = "1" Then '*** Update timer Vietnam Only, last 7 days..? depends
                        
                    'dlbTimer = FormatNumber(dlbTimer, 2)
                    'dlbTimer = Replace(dlbTimer, ",", ".")
                    
                    'Dim strSQL As String = "UPDATE timer SET timer = " & dlbTimer & " WHERE extSysId = " & intCatiId & " AND origin = 1"
                        
                    'objCmd = New OdbcCommand(strSQL, objConn)
                    'objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    'objDR.Close()
                    
                    'End If
                    
                End If ' 6 findes i forvejen
                
                
                t = t + 1
                    
            Next
            
        Next
        
        'Return "HER " 'testValue
      
            
        'Catch When Err.Number <> 0
            
        'Return "<br>ErrNO: " & Err.Number & " ErrDesc: " & Err.Description & " # " & errThis & " # CATI recordID: (" & intCatiId & ")"
        'Exit Try
            
        '''Finally
            
           
            
        'End Try
        
        
        'rtErrTxt(errThisAll)
        'errThisAll = ""
        'Return errThisAll
           
        '*** Mail ***'

         return true
        
    End Function

    
    
    'Public Function rtErrTxt(ByVal erTxt As String) As String
    ' Dim eTxt As String
    '    eTxt = erTxt
        
    '   Return eTxt
    'End Function
   
   
    
End Class

        



