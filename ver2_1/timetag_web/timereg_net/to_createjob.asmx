



<%@ WebService language="VB" class="TO_CREATEJOB" %>
Imports System
Imports System.Web.Services
Imports System.Data
Imports System.Data.Odbc





Public Class TO_CREATEJOB : 

    
 
    
    
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
    
    
    Public kundenr As String
    Public kid As Integer = 0
    Public kunderef As Integer = 0
    Public jobnavn As String
    Public jobansvarlig As Integer
    Public rekvisitionsnr As String
    Public internnote As String
    Public jobbeskrivelse As String
    Public jbstdato As Date
    Public jbsldato As Date
    Public bruttooms As Double
    Public aktorfase As Integer
    Public errThis As Integer = 0
    
    Public lto As String
    Public cdDato As Date = DateTime.Now
    Public cdDatoSQL As Date
    Public editor As String
    
    
    Public strjnr As String

   
   
    
    <WebMethod()> Public Function createjob(ByVal ds As DataSet) As String
        
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
        Dim t As Double = 0
        
        'Try
       
        
     
        
        For Each dt In ds.Tables
            For Each dr In dt.Rows
                
              
                errThis = 0
                
                '** Opret job t = 0 **'
                'to_lto_c      'fra link
                'to_editor_c   'fra link
                
                'to_jobnavn_c
                'to_kundeid_c
                'to_kunderef_c
                'to_jobstdato_c
                'to_jobsldato_c
                
                'to_jobansvarlig_c
                'to_rekvisitionsnr_c
                'to_note_c
                'to_jobbeskrivelse_c
               
                'to_jobbruttooms_c
                'to_aktorfase_c
                'to_aktorfase_line1_antal_c
                'to_aktorfase_line1_navn_c
                'to_aktorfase_line1_stkpris_c
                'to_aktorfase_line1_pris_c
                'to_aktorfase_line1_dato_c
                
                'to_akt_list

               
             
                
                
                '** Opret job og Tilføj aktiviteter / Faser *

                If t = 0 Then
                    
                    
                    
                    If ds.Tables("TO_CREATEJOB_TB").Columns.Contains("to_lto_c") Then
                        lto = ds.Tables("TO_CREATEJOB_TB").Rows(t).Item("to_lto_c")
                    Else
                        lto = 0
                    End If
                    
                    If ds.Tables("TO_CREATEJOB_TB").Columns.Contains("to_editor_c") Then
                        editor = ds.Tables("TO_CREATEJOB_TB").Rows(t).Item("to_editor_c")
                    Else
                        editor = "TimeOut System"
                    End If
                    
                    Dim cdDatoSQL As String =  "2013-12-01" '""& Day(cdDato) & " " & Left(MonthName(Month(cdDato)), 3) & ". " & year(cdDato) & ""
                    
                    
                    If lto = "outz" OR lto = "intranet - local" Then
                        lto = "intranet"
                    End If
                 
                'strConn = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_epi_catitest;User=outzource;Password=SKba200473;"
                strConn = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_" & lto & ";User=outzource;Password=SKba200473;"
                  
                    
                '** Åbner Connection ***'
                objConn = New OdbcConnection(strConn)
                objConn.Open()
                    
                    
                    '*** Finder nyt jobnr '****
                    Dim strSQLjn As String = "SELECT jobnr FROM licens WHERE id = 1"
                    objCmd = New OdbcCommand(strSQLjn, objConn)
                    objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                    If objDR.Read() = True Then

                        strjnr = objDR("jobnr") + 1
                    
                    End If
            
                    objDR.Close()
                    
                    
                    
                End If

                
                
                
                If ds.Tables("TO_CREATEJOB_TB").Columns.Contains("to_jobnavn_c") Then
                    jobnavn = ds.Tables("TO_CREATEJOB_TB").Rows(t).Item("to_jobnavn_c")
                Else
                    jobnavn = "Jobnavn.."
                End If
                
              
                If ds.Tables("TO_CREATEJOB_TB").Columns.Contains("to_kundeid_c") Then
                    kundenr = ds.Tables("TO_CREATEJOB_TB").Rows(t).Item("to_kundeid_c")
                Else
                    kundenr = 0
                    errThis = 1
                End If
                
                        
                If CInt(errThis) = 0 Then
                    
                    '*** Finder nyt jobnr '****
                    Dim strSQLjn As String = "SELECT kid FROM kunder WHERE kkundenr = '" & kundenr & "'"
                    objCmd = New OdbcCommand(strSQLjn, objConn)
                    objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                    If objDR.Read() = True Then

                        kid = objDR("kid")
                    
                    End If
            
                    objDR.Close()
                
                    If kid <> 0 Then
                        errThis = 2
                    End If
                
                    '***'
                
                End If
                
                
                If ds.Tables("TO_CREATEJOB_TB").Columns.Contains("to_kunderef_c") Then
                    kunderef = ds.Tables("TO_CREATEJOB_TB").Rows(t).Item("to_kunderef_c")
                Else
                    kunderef = 0
                End If
                
                If ds.Tables("TO_CREATEJOB_TB").Columns.Contains("to_jobstdato_c") Then
                    jbstdato = ds.Tables("TO_CREATEJOB_TB").Rows(t).Item("to_jobstdato_c")
                    jbstdato.ToString("yyyy-mm-dd")
                Else
                    jbstdato = cdDato
                    jbstdato.ToString("yyyy-mm-dd")
                End If
           
                
                If ds.Tables("TO_CREATEJOB_TB").Columns.Contains("to_jobsldato_c") Then
                    jbsldato = ds.Tables("TO_CREATEJOB_TB").Rows(t).Item("to_jobsldato_c")
                    jbstdato.ToString("yyyy-mm-dd")
                Else
                    jbsldato = "2044-01-01"
                End If
                
                    
                
                    'If ds.Tables("TO_CREATEJOB_TB").Columns.Contains("to_jobbeskrivelse_c") Then
                    If String.IsNullOrWhiteSpace(ds.Tables("TO_CREATEJOB_TB").Rows(t).Item("to_jobbeskrivelse_c").ToString()) = False Then
                        jobbeskrivelse = ds.Tables("TO_CREATEJOB_TB").Rows(t).Item("to_jobbeskrivelse_c")
                    Else
                        jobbeskrivelse = ".."
                    End If
                
                    If ds.Tables("TO_CREATEJOB_TB").Columns.Contains("to_rekvisitionsnr_c") Then
                        rekvisitionsnr = ds.Tables("TO_CREATEJOB_TB").Rows(t).Item("to_rekvisitionsnr_c")
                    Else
                        rekvisitionsnr = ""
                    End If
                
                    'If ds.Tables("TO_CREATEJOB_TB").Columns.Contains("to_note_c") Then
                    'internnote = ds.Tables("TO_CREATEJOB_TB").Rows(t).Item("to_note_c")
                    'Else
                    internnote = ""
                    'End If
                
                    'If ds.Tables("TO_CREATEJOB_TB").Columns.Contains("to_jobbruttooms_c") Then
                    'bruttooms = ds.Tables("TO_CREATEJOB_TB").Rows(t).Item("to_jobbruttooms_c")
                    'Else
                    bruttooms = "0"
                    'End If
                
                

                
                    'Call jq_format(kNavn)
                    'kNavn = jq_formatTxt
                    
                        
                    'Call DecodeUTF8(jobNavn)
                    
                                                    
                                                
                                                
                                                
                                                
                    errThis = 0
                                                 
                    '*** Opretter Job ***'
                    If CInt(errThis) = 0 And t = 0 Then
                                                    
                                                    
  

                        Dim strSQLjob As String = ("INSERT INTO job (jobnavn, jobnr, jobknr, jobTpris, jobstatus, jobstartdato," _
                       & " jobslutdato, editor, dato, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, " _
                       & " projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, " _
                       & " fakturerbart, budgettimer, fastpris, kundeok, beskrivelse, " _
                       & " ikkeBudgettimer, tilbudsnr, jobans1, " _
                       & " serviceaft, kundekpers, valuta, rekvnr, " _
                       & " risiko, job_internbesk, " _
                       & " jo_bruttooms) VALUES " _
                       & "('" & jobnavn & "', " _
                       & "'" & strjnr & "', " _
                       & "" & kid & ", " _
                       & "0, " _
                       & "1, " _
                       & "'" & jbstdato & "', " _
                       & "'" & jbsldato & "', " _
                       & "'" & editor & "', " _
                       & "'" & cdDatoSQL & "', " _
                       & "10, " _
                       & "1,1,1,1,1,1,1,1,1," _
                       & "1,0,1,0," _
                       & "'" & jobbeskrivelse & "', " _
                       & "0,0, " _
                       & "" & jobansvarlig & "," _
                       & "0," & kunderef & ", " _
                       & "1, '" & rekvisitionsnr & "', " _
                       & "100,'" & internnote & "'," _
                       & "" & bruttooms & ")")
    							
                    
        
                        objCmd = New OdbcCommand(strSQLjob, objConn)
                        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                        'objDR.Close()
                                                    
                                                    
                    Dim strSQLjobnr As string = "UPDATE licens SET jobnr = "& strjnr &" WHERE id = 1"                      
                    objCmd = New OdbcCommand(strSQLjobnr, objConn)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
                                                        
                                               
                                            
                                                
               
                    
                    End If ' 6 findes i forvejen
                
                
                    t = t + 1
                    
            Next
            
        Next
        
        
      
            
        'Catch When Err.Number <> 0
            
        'Return "<br>ErrNO: " & Err.Number & " ErrDesc: " & Err.Description & " # " & errThis & " # CATI recordID: (" & intCatiId & ")"
        'Exit Try
            
        '''Finally
            
           
            
        'End Try
        
        
        'rtErrTxt(errThisAll)
        'errThisAll = ""
        'Return errThisAll
           
        '*** Mail ***'

    End Function

    
    
    'Public Function rtErrTxt(ByVal erTxt As String) As String
    ' Dim eTxt As String
    '    eTxt = erTxt
        
    '   Return eTxt
    'End Function
   
    
End Class

        




