



<%@ WebService language="VB" class="CATI" %>
Imports System
Imports System.Web.Services
Imports System.Data
Imports System.Data.Odbc





Public Class CATI
    
    Public errThisAll As string = ""
    
    Public intMedarbId As String
    Public intCatiId, intJobNr As String
    Public dlbTimer As String
    Public cdDato As String
    
    Public errThis As Integer = 0
        
    Public meNavn As String
    Public meNr As String
    Public meID As Integer
    
    Public jobNavn As String
    Public jobId As Double
    Public jobFastPris As Integer
    Public jobSeraft As Double
    Public jobKid As Double
    Public jobValuta As Integer
    
    Public aktNavn As String
    Public aktId As Double = 0
    Public aktFakturerbar As Integer
    
    Public kurs As String
    
    Public kNavn As String
    Public kNr As Double
    
    Public extsysid As Double
    
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
    
    
    <WebMethod()> _
    Public Function upDateCatiTimer(ByVal ds As DataSet) As String
        
        'Dim con As New SqlConnection(Application("MyDatabaseConnectionString"))
        'Dim daCust As New SqlDataAdapter("Select * From tblCustomer", con)
        'Dim cbCust As New SqlCommandBuilder(daCust)
        'daCust.Update(ds, "Cust")
        'Return ds
        
        Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_epi_catitest;User=outzource;Password=SKba200473;"
        'Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_intranet;User=outzource;Password=SKba200473;"

        Dim objConn As OdbcConnection
        Dim objCmd As OdbcCommand
        'Dim objDataSet As New DataSet
        Dim objDR As OdbcDataReader
        Dim objDR2 As OdbcDataReader
        Dim dt As DataTable
        Dim dr As DataRow
        
        

        '** Åbner Connection ***'
        objConn = New OdbcConnection(strConn)
        objConn.Open()
        
        'Dim objTable As DataTable
        
        Dim t As double = 0
        
       
        
        For Each dt In ds.Tables
            For Each dr In dt.Rows
                
                
              
                errThis = 0
                meID = 0
                extsysid = 0
                aktId = 0
                intJobNr = 0
                'intjobId = 0
                intValuta = 0
                intMedarbId = ""
                jobKid = 0
                
                'intCatiId = 1
                
               
                intCatiId = ds.Tables("CATI_TIME").Rows(t).Item("id")
                intMedarbId = ds.Tables("CATI_TIME").Rows(t).Item("medarbejderid")
                
                'intJobNr = 14022
                intJobNr = ds.Tables("CATI_TIME").Rows(t).Item("jobid")
            
                'dlbTimer = 1
                dlbTimer = ds.Tables("CATI_TIME").Rows(t).Item("timer")
                
                'cdDato = "2010-08-22"
                cdDato = ds.Tables("CATI_TIME").Rows(t).Item("dato")
                 
               
                
                errThisAll = errThisAll & ";" & errThis & "(" & intCatiId & ")"
        
                'Dim indtastningfindes AS integer = 0 
                
                '*** Er timer angive korrekt ***'
                If Len(trim(dlbTimer)) = 0 OR InStr(dlbTimer, "-") <> 0 Then
                    
                    dlbTimer = 0
                    errThis = 7
                    
                End If
       
                '*** Finder medarbejder oplysninger ***'
                If intMedarbId <> "" Then
                    intMedarbId = intMedarbId
                Else
                    intMedarbId = ""
                    errThis = 1
                End If
        
                
       
                If errThis <> 1 Then
            
        
                    Dim strSQLme As String = "SELECT mid, mnavn, mnr FROM medarbejdere WHERE init = '" & intMedarbId & "'"
                    objCmd = New OdbcCommand(strSQLme, objConn)
                    objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                    If objDR.Read() = True Then

                        meNavn = objDR("mnavn")
                        meNr = objDR("mnr")
                        meID = objDR("mid")

                    End If
            
                    objDR.Close()
            
                End If
                '***'
                
                
                If meID <> 0 Then
                    meID = meID
                Else
                    meID = 0
                    errThis = 11
                End If
            
      
                '*** Finder job oplysninger ****'
                If intJobNr <> "" Then
                    intJobNr = intJobNr
                Else
                    intJobNr = 0
                    errThis = 2
                End If
                
                
                jobId = 0
       
                Dim strSQLjob As String = "SELECT id, jobnavn, fastpris, serviceaft, jobknr, valuta FROM job WHERE jobnr = " & intJobNr
                objCmd = New OdbcCommand(strSQLjob, objConn)
                objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                If objDR.Read() = True Then

                    jobNavn = Replace(objDR("jobnavn"), "'", "''")
                    jobId = objDR("id")
                    jobFastPris = objDR("fastpris")
                    jobSeraft = objDR("serviceaft")
                    jobKid = objDR("jobknr")
                    'jobValuta = objDR("valuta")
            
               

                End If
            
                objDR.Close()
                '***'
                
                
                If jobId = 0 Then
                    
                    jobNavn = ""
                    jobId = 0
                    jobFastPris = 0
                    jobSeraft = 0
                    jobKid = 0
                    intJobNr = 0
                    errThis = 21
                    
                End If
                
                
                
                
                If errThis = 0 Then
                    
                    '*** Finder akt. oplysninger ****'
                    '*** Fase ??? ***'
                    Dim strSQLa As String = "SELECT id, fakturerbar, navn FROM aktiviteter WHERE job = " & jobId & " AND navn = 'Interviewer' AND aktstatus = 1"
                    objCmd = New OdbcCommand(strSQLa, objConn)
                    objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                    If objDR.Read() = True Then

                        aktNavn = objDR("navn")
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
        
               
        
        
                If errThis = 0 Then
        
                    '*** Finder medarb timepris og kostpris ***'
                    '*** Først prøves aktivitet derefter job ***'
                
                
                    For mtp = 0 To 1
                    
                        If foundone = "n" Then
                        
                    
                    
                            If mtp = 0 Then
                                aktIdUse = aktId
                            Else
                                aktIdUse = 0
                            End If
                        
                
                            Dim strSQLmtp As String = "SELECT id AS tpid, jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta FROM timepriser WHERE jobid = " & jobId & " AND aktid = " & aktId & " AND medarbid =  " & meID
								
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
                    
                        
                        
                        
                        
                                    Dim strSQL3 As String = "SELECT mid, " & timeprisalernativ & " AS useTimepris, " & valutaAlt & " AS useValuta, " _
                                    & " medarbejdertype FROM medarbejdere, medarbejdertyper WHERE mid =" & meID & " AND medarbejdertyper.id = medarbejdertype"
									
                                    objCmd = New OdbcCommand(strSQL3, objConn)
                                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                                    If objDR2.Read() = True Then
                                        intTimepris = objDR2("useTimepris")
                                        intValuta = objDR2("useValuta")
                                    End If
                                    objDR2.Close()
									
                        
                                    intTimepris = Replace(intTimepris, ",", ".")
									
									
									
                      
                
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
                
                
                End If
                
                
                If errThis = 0 Then
                
							            
                    '** Finder valuta og kurs **'
                    If intValuta <> 0 Then
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
                    '***'		            
					
				
                End If
			    
                
                
                If errThis = 0 Then
        
                    '*** Finder kunde oplysninger ****'
                    If jobKid <> 0 Then
                        jobKid = jobKid
                    Else
                        jobKid = 0
                        errThis = 5
                    End If
        
      
                    Dim strSQLk As String = "SELECT kkundenavn, kkundenr FROM kunder WHERE kid = " & jobKid
                    objCmd = New OdbcCommand(strSQLk, objConn)
                    objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                    If objDR.Read() = True Then

                        kNavn = objDR("kkundenavn")
                        kNr = objDR("kkundenr")
               

                    End If
            
                    objDR.Close()
                    '***'
                
                End If
                
                
                If errThis = 0 Then
                    
                    '*** tjekker om extsysid findes i forvejen ***'
                    Dim strSQLext As String = "SELECT extsysid FROM Timer WHERE extsysid = " & intCatiId
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
                    
                    
        
                If CInt(errThis) = 0 Then
                        
                    
                    dlbTimer = FormatNumber(dlbTimer, 2)
                    
                    '*** Indlæser Timer ***'
                    Dim strSQL As String = "INSERT INTO timer " _
                    & "(" _
                    & " timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, " _
                    & " TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, " _
                    & " editor, kostpris, seraft, " _
                    & " valuta, kurs, extSysId, sttid, sltid " _
                    & ") " _
                    & " VALUES " _
                    & " (" _
                    & Replace(dlbTimer, ",", ".") & ", " & aktFakturerbar & ", " _
                    & "'" & Year(cdDato) & "/" & Month(cdDato) & "/" & Day(cdDato) & "', " _
                    & "'" & meNavn & "', " _
                    & meID & ", " _
                    & "'" & jobNavn & "', " _
                    & intJobNr & ", " _
                    & "'" & kNavn & "', " _
                    & kNr & ", " _
                    & aktId & ", " _
                    & "'" & aktNavn & "', " _
                    & Year(Now) & ", " _
                    & intTimepris & ", " _
                    & "'" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "', " _
                    & jobFastPris & ", " _
                    & "'00:00:01', " _
                    & "'CaTi Import', " _
                    & Replace(kostpris, ",", ".") & ", " _
                    & jobSeraft & ", " _
                    & intValuta & ", " _
                    & Replace(kurs, ",", ".") & ", " & intCatiId & ", '00:00:00', '00:00:00')"

           
                    '** Manger salgs og kost priser ***'

                    'Response.write(strSQL)
        
                    objCmd = New OdbcCommand(strSQL, objConn)
                    objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    objDR.Close()
                    
                    errThisAll = errThisAll
                    
                Else '** Indlæser ind til ErrLog
                        
                        
                       
                        
                    Dim strSQLer As String = "INSERT INTO timer_imp_err (dato, extsysid, errid) " _
                    & " VALUES " _
                    & " ('" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "', " & intCatiId & ", " & errThis & ")"

           
                    objCmd = New OdbcCommand(strSQLer, objConn)
                    objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    objDR.Close()
            
                    'Response.write(strSQLer)
                        
                    errThisAll = errThisAll & ";" & errThis & "(" & intCatiId & ")"
                    
                        
            
                End If
        
                    
                
                
                t = t + 1
                
            Next
            
        Next
        
      
        
        Return errThisAll
        
        
        '*** Mail ***'

    End Function

    
End Class

        




