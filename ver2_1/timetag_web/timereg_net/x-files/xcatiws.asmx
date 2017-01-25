

<%@ WebService language="VB" class="CATI" %>
Imports System
Imports System.Web.Services
Imports System.Data
Imports System.Data.Odbc




Public Class CATI
    
    Public intMedarbId As String
    Public intCatiId, intJobNr As String
    Public dlbTimer As String
    Public cdDato As String
    
    Public errThis As Integer = 0
        
    Public meNavn As String
    Public meNr As Integer
    
    Public jobNavn As String
    Public jobId As Integer
    Public jobFastPris As Integer
    Public jobSeraft As Integer
    Public jobKid As Integer
    Public jobValuta As Integer
    
    Public aktNavn As String
    Public aktId As Integer = 0
    Public aktFakturerbar As Integer
    
    Public kurs As Double
    
    Public kNavn As String
    Public kNr As Integer
    
    <WebMethod()> _
    Public Function upDateCatiTimer(ByVal ds As DataSet) As DataSet()
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
        Dim dt As DataTable
        Dim dr As DataRow
        
        

        '** Åbner Connection ***'
        objConn = New OdbcConnection(strConn)
        objConn.Open()
        
        'Dim objTable As DataTable
        
        Dim t As Integer = 0
        
        For Each dt In ds.Tables
            For Each dr In dt.Rows
                
           
      
        
                intCatiId = ds.Tables("CATI_TIME").Rows(t).Item("id")
                intMedarbId = ds.Tables("CATI_TIME").Rows(t).Item("medarbejderid")
        
                intJobNr = ds.Tables("CATI_TIME").Rows(t).Item("jobid")
        
        
                dlbTimer = ds.Tables("CATI_TIME").Rows(t).Item("timer")
                cdDato = ds.Tables("CATI_TIME").Rows(t).Item("dato")
        
        
                'Dim indtastningfindes AS integer = 0 
       
                '*** Finder medarbejder oplysninger ***'
                If intMedarbId <> "" Then
                    intMedarbId = intMedarbId
                Else
                    intMedarbId = 0
                    errThis = 1
                End If
        
        
       
                If errThis <> 1 Then
            
        
                    Dim strSQLme As String = "SELECT mnavn, mnr FROM medarbejdere WHERE mid = " & intMedarbId
                    objCmd = New OdbcCommand(strSQLme, objConn)
                    objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                    If objDR.Read() = True Then

                        meNavn = objDR("mnavn")
                        meNr = objDR("mnr")
               

                    End If
            
                    objDR.Close()
            
                End If
                '***'
       
        
      
                '*** Finder job oplysninger ****'
                If intJobNr <> "" Then
                    intJobNr = intJobNr
                Else
                    intJobNr = 0
                    errThis = 2
                End If
        
       
                Dim strSQLjob As String = "SELECT id, jobnavn, fastpris, serviceaft, jobknr, valuta FROM job WHERE jobnr = " & intJobNr
                objCmd = New OdbcCommand(strSQLjob, objConn)
                objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                If objDR.Read() = True Then

                    jobNavn = objDR("jobnavn")
                    jobId = objDR("id")
                    jobFastPris = objDR("fastpris")
                    jobSeraft = objDR("serviceaft")
                    jobKid = objDR("jobknr")
                    jobValuta = objDR("valuta")
            
               

                End If
            
                objDR.Close()
                '***'
        
        
        
                '*** Finder akt. oplysninger ****'
                '*** Fase ??? ***'
                Dim strSQLa As String = "SELECT id, fakturerbar, navn FROM aktiviteter WHERE job = " & intJobNr & " AND navn = 'Interviewer' AND aktstatus = 1"
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
        
                '** Finder valuta og kurs **'
                If jobValuta <> 0 Then
                    jobValuta = jobValuta
                Else
                    jobValuta = 0
                    errThis = 4
                End If
        
        
        
        
                Dim strSQLv As String = "SELECT kurs FROM valutaer WHERE id = " & jobValuta
                objCmd = New OdbcCommand(strSQLv, objConn)
                objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                If objDR.Read() = True Then

                    kurs = objDR("kurs")
          
                End If
            
                objDR.Close()
                '***'
        
        
        
        
                '*** Finder medarb timepris og kostpris ***'
        
        
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
        
                '*** tjekker om extsysid findes i forvejen ***'
        
        
                If CInt(errThis) = 0 Then
                        
                    '*** Indlæser Timer ***'
                    Dim strSQL As String = "INSERT INTO timer " _
                    & "(" _
                    & " timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, " _
                    & " TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, " _
                    & " editor, kostpris, seraft, " _
                    & " valuta, kurs, extSysId " _
                    & ") " _
                    & " VALUES " _
                    & " (" _
                    & Replace(dlbTimer, ",", ".") & ", " & aktFakturerbar & ", " _
                    & "'" & Year(cdDato) & "/" & Month(cdDato) & "/" & Day(cdDato) & "', " _
                    & "'" & meNavn & "', " _
                    & meNr & ", " _
                    & "'" & jobNavn & "', " _
                    & intJobNr & ", " _
                    & "'" & kNavn & "', " _
                    & kNr & ", " _
                    & aktId & ", " _
                    & "'" & aktNavn & "', " _
                    & Year(Now) & ", " _
                    & 100 & ", " _
                    & "'" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "', " _
                    & jobFastPris & ", " _
                    & "'00:00:01', " _
                    & "'CaTi Import', " _
                    & Replace(100, ",", ".") & ", " _
                    & jobSeraft & ", " _
                    & jobValuta & ", " _
                    & Replace(kurs, ",", ".") & ", " & intCatiId & ")"

           
                    '** Manger salgs og kost priser ***'

                    'Response.write(strSQL)
        
                    objCmd = New OdbcCommand(strSQL, objConn)
                    objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    objDR.Close()
            
                Else '** Indlæser ind til ErrLog
            
            
                    Dim strSQLer As String = "INSERT INTO timer_imp_err (dato, extsysid, errid) " _
                    & " VALUES " _
                    & " ('" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "', " & intCatiId & ", " & errThis & ")"

           
                    objCmd = New OdbcCommand(strSQLer, objConn)
                    objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                    objDR.Close()
            
                    'Response.write(strSQLer)
            
            
                End If
        
            
                t = t + 1
                
            Next
            
        Next

    End Function

    
End Class

        




