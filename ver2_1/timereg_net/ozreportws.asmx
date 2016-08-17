<%@ WebService Language="VB" Class="ozreportws" %>


Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Collections.Generic

Imports System
Imports System.Data
Imports System.Data.Odbc

Imports System.IO



' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _

'<WebService(Namespace:="http://tempuri.org/")> _
'<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
'<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _

Public Class ozreportws
    Inherits System.Web.Services.WebService

    <WebMethod()> _
    Public Function oz_report() As List(Of ozreportcls)

        'dim x As integer
        'For x = 0 To 0 Step 1
       
        Dim startDatoAtDSQL As string
        
       
        Dim employeeIDsPgrpNavn As string
        
        Dim lstRet As New List(Of ozreportcls)()
        
        Dim modtEmail As String
        Dim modtName As String
        Dim strConn As String
        Dim lto As String
        Dim rapporttype As String
        Dim aboMtyper As String
        Dim aboPgrp As string
        Dim medid As string
        Dim show_atd As String = "0"
       
        
            '*** Alle der t�ller med i dagligt timeregnskab ***'
        Dim strConn_admin As string
        strConn_admin = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_admin;User=outzource;Password=SKba200473;"
        
        Dim objConn_admin As OdbcConnection
        Dim objCmd_admin As OdbcCommand
            'Dim objDataSet As New DataSet
        Dim objDR_admin As OdbcDataReader
       
            '** �bner Connection ***'
        objConn_admin = New OdbcConnection(strConn_admin)
        objConn_admin.Open()
        
        
        Dim strSQLadmin As String = "SELECT email, navn, lto, rapporttype, medid, medarbejdertyper, projektgrupper FROM rapport_abo WHERE id <> 0 AND (rapporttype = 1 OR rapporttype = 0)"
        objCmd_admin = New OdbcCommand(strSQLadmin, objConn_admin)
        objDR_admin = objCmd_admin.ExecuteReader '(CommandBehavior.closeConnection)
            
        While objDR_admin.Read() = True
             
            
                'Dim errThisTOnoStr As String = errThisTOno.ToString()
            
            
            
                'Do Until objDR_admin.EOF
            
            
            medid = objDR_admin.Item("medid")
            rapporttype = objDR_admin.Item("rapporttype")
            modtEmail = objDR_admin.Item("email") '"ad@dencker.net"
            modtName = objDR_admin.Item("navn")
            aboMtyper = objDR_admin.Item("medarbejdertyper")
            aboPgrp = objDR_admin.Item("projektgrupper")
            
            
            
            lto = objDR_admin.Item("lto")
            
                'Skal modtager kun modtage sin egen data eller skal han modtage alle (Admin)
                'Dim strSQLadmin As String = "SELECT email, navn, lto FROM rapport_abo WHERE id <> 0"
                'objCmd_admin = New OdbcCommand(strSQLadmin, objConn_admin)
                'objDR_admin = objCmd_admin.ExecuteReader '(CommandBehavior.closeConnection)
            
                'If objDR_admin.Read() = True Then
                
            'End If
            
            Select Case lto
                Case "dencker"
                    startDatoAtDSQL = Year(Now) & "/4/21"
                Case Else
             
            startDatoAtDSQL = Year(Now) & "/1/1"
            End Select
                
            
            strConn = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_" & lto & ";User=outzource;Password=SKba200473;"
            'strConn = "timeout_"& lto &"64"
            
            
            Select Case objDR_admin.Item("lto")
                    Case "dencker" '** Dencker
                 
                    show_atd = "1"
                    
                    Case "epi", "epi_no" '** EPINION
                 
                    show_atd = "2"
                    
                    Case "outz", "intranet - local"
                    strConn = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_intranet;User=outzource;Password=SKba200473;"
                    'strConn = "timeout_intranet64"
                        show_atd = "0"
                    Case Else
                        show_atd = "0"
            
                End Select
            
            
       
        
        
        
                Dim objConn As OdbcConnection
                Dim objCmd As OdbcCommand
                'Dim objDataSet As New DataSet
                Dim objDR As OdbcDataReader
                Dim objDR2 As OdbcDataReader
                'Dim dt As DataTable
                'Dim dr As DataRow
        
                '** �bner Connection ***'
                objConn = New OdbcConnection(strConn)
                objConn.Open()
        
        
                Dim dDato As Date = Date.Now 'tirsdag morgen 06.00
            Dim slutDato As Date = dDato.AddDays(-2) '- 2
            Dim startDato As Date = dDato.AddDays(-8) '-8
        
                Dim format As String = "yyyyMd_hhmmss"
                Dim fnEnd As String = slutDato.ToString(format) & "_" & medid

        
                Dim datSQLformat As String = "yyyy-M-d"
                Dim startDatoSQL As String = startDato.ToString(datSQLformat)
                Dim slutDatoSQL As String = slutDato.ToString(datSQLformat)
        
                'slutDatoSQL = "2012-5-30" 'slutDato.ToString '.Day(slutDato) & "-" & Month(slutDato) & "-" & Year(slutDato)
                'startDatoSQL = "2012-4-30" 'startDato.ToString 'Day(slutDato) & "-" & Month(slutDato) & "-" & Year(slutDato)
        
                Dim datoformatTxt As String = "d-M-yyyy"
                Dim startDatoTxt As String = startDato.ToString(datoformatTxt)
                Dim slutDatoTxt As String = slutDato.ToString(datoformatTxt)
        
        
                Dim sygTimer As String = 0
                Dim andreTimer As String = 0
                Dim ferieTimer As String = 0
        
                Dim e1Timer As String = 0
            
                Dim ExpTxtheader As String
                Dim expTxt As String
        
            
                Dim lTim As String = 0
                Dim fakBareReal As String = 0
                Dim effektiv_proc As String = 0
            
                Dim lTim_atd As String = 0
                Dim fakBareReal_atd As String = 0
                Dim effektiv_proc_atd As String = 0
            
                Dim realTimer As String = 0
                Dim normTimer As String = 0
            Dim bal_norm_real As String = 0
            Dim bal_norm_realAtd As String = 0
            Dim bal_norm_lontimer As String = 0
                Dim realTimerAtd As string = 0
            
                Dim ugestatus As String = 0
                Dim ugegodkendt As String = 0
            
                Dim lTimWrt As String = 0
                Dim lTim_addThis As String = 0
            
                Dim lTim_atd_pau As String = 0
                Dim lTim_pau As String = 0
            
                Dim lTim_pau_addThis As String = 0
            
                Dim aboMtyperArr As Array
                Dim aboPgrpArr As Array
            
                Dim ugeNrLastWeek As integer = 0
            
                ExpTxtheader = "Medarbejder;Medarb. Nr;Initialer;Norm. tid;L�ntimer (komme/g�);Realiseret tid;(heraf fakturerbare);"

                If lto = "dencker" Then
                    ExpTxtheader = ExpTxtheader & "E1;"
                End If
            
            ExpTxtheader = ExpTxtheader & "Syg/barn syg;Ferie/Feriefri;Anden frav�r;"
            
            Select Case lto
                Case "cst", "kejd_pb", "outz"
                    ExpTxtheader = ExpTxtheader & "Bal. (Norm./L�ntimer komme/g�);"
                Case Else
                    ExpTxtheader = ExpTxtheader & "Bal. (Norm./Real.);"
            End Select
            

            
            If show_atd = "1" Then
                    ExpTxtheader = ExpTxtheader & "Effektiv % (l�ntimer/fakturerbare timer real.);Effektiv % �TD - "& startDatoAtDSQL &" - (l�ntimer/fakturerbare timer real.);"
            End If
            
            If show_atd = "2" Then
                ExpTxtheader = ExpTxtheader & "Real. �TD;Bal. (Norm./Real.) �TD;"
            End If
            
            
            ExpTxtheader = ExpTxtheader & "Uge afsluttet;Uge godkendt;"
            
            If InStr(lto, "epi") Then
                ExpTxtheader = ExpTxtheader & "Projekgrupper;"
            End If
            
        
                '"& Left(startDato.ToString, 4) &"
                Dim flname As String = "timeout_rapport_" & lto & "_" & fnEnd & ".csv"
        
                Using writer As StreamWriter = New StreamWriter("D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_14\inc\upload\" & lto & "\" & flname, False, Encoding.GetEncoding("iso-8859-1"))
                    
                    writer.WriteLine("Periode: " & startDatoTxt & " - " & slutDatoTxt & ", uge: " & DatePart("ww", slutDato, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays) & ";")
                    writer.WriteLine(ExpTxtheader)
            
            
                    '*** Alle der t�ller med i dagligt timeregnskab ***'
                    Dim akttyper_realhours As String = " AND (tfaktim = 0 "
                    Dim strSQLa As String = "SELECT aty_id FROM akt_typer WHERE aty_on = 1 AND aty_on_realhours = 1 ORDER BY aty_id"
                    objCmd = New OdbcCommand(strSQLa, objConn)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                    While objDR2.Read() = True
                
                        akttyper_realhours = akttyper_realhours & " OR tfaktim = " & objDR2("aty_id")
                
                    End While
                    objDR2.Close()
            
                    akttyper_realhours = akttyper_realhours & ")"
            
            
                    '*** Fakturerbare ***'
                    Dim akttyper_invoiceable As String = " AND (tfaktim = 0 "
                    Dim strSQLf As String = "SELECT aty_id FROM akt_typer WHERE aty_on = 1 AND aty_on_invoiceble = 1 ORDER BY aty_id"
                    objCmd = New OdbcCommand(strSQLf, objConn)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                    While objDR2.Read() = True
                
                        akttyper_invoiceable = akttyper_invoiceable & " OR tfaktim = " & objDR2("aty_id")
                
                    End While
                    objDR2.Close()
            
                    akttyper_invoiceable = akttyper_invoiceable & ")"
                
           
            
            
            
            
            
                    normTimer = 0
            
                    '*** Henter aktive medarbejdere *****'
                    '*** Finder timer p� de valgte typer ***'
                    Dim employeeIDs As String = ""
                    Dim employeeIDsTyp As String = ""
                    Dim employeeIDsPgrp As String = ""
                    'If rapporttype = 0 Then 'Admin = all employess
                
                    '*** Valgte medarbejdertyper **'
                    If aboMtyper = "-2" Then 'sig selv
                    
                
                        employeeIDs = " AND (mid = " & medid & ")"
                    
                    Else '-2
                
                        If aboMtyper <> "-1" Then
                    
                            If aboMtyper <> "0" Then
                        
                                'employeeIDs = "(mid = 0"
                        
                                employeeIDsTyp = ""
                        
                                aboMtyperArr = Split(aboMtyper, ", ")
                                Dim t As Integer = 0
                                For t = 0 To UBound(aboMtyperArr)
                           
                            
                            
                                    Dim strSQLmtyp As String = "SELECT mid FROM medarbejdere WHERE mansat = 1 AND medarbejdertype = " & aboMtyperArr(t)
                        
                                    objCmd = New OdbcCommand(strSQLmtyp, objConn)
                                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                                    While objDR2.Read() = True
                            
                                        employeeIDsTyp = employeeIDsTyp & " OR mid = " & objDR2("mid")
                                
                                    End While
                                    objDR2.Close()
                       
                                Next
                            
                                'employeeIDs =+ ")"
                        
                            Else
                                employeeIDsTyp = " AND mid <> 0"
                            End If
                    
                    
                        Else
                            employeeIDsTyp = " AND mid = 0"
                        End If
                
                
                        '*** Valgte projektgrupper **' 
                        If aboMtyper = "-1" Then 'abonner via projektgrupper
                            If aboPgrp <> "-1" Then
                    
                                If aboPgrp <> "0" Then
                        
                      
                                    employeeIDsPgrp = ""
                                    aboPgrpArr = Split(aboPgrp, ", ")
                                    Dim t As Integer = 0
                                    Dim tjkValue As String
                        
                            
                      
                                
                           
                           
                                    For t = 0 To UBound(aboPgrpArr)
                           
                            
                      
                                        Dim strSQLpgrp As String = "SELECT p.id AS pgrpid, pr.medarbejderId FROM projektgrupper AS p LEFT JOIN progrupperelationer AS pr ON (pr.projektgruppeId = p.id)  WHERE p.id  = " & aboPgrpArr(t) & " GROUP BY pr.medarbejderId"
                        
                            
                                
                                
                                        objCmd = New OdbcCommand(strSQLpgrp, objConn)
                                        objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                                        While objDR2.Read() = True
                                
                                            tjkValue = objDR2("medarbejderId")
                                   
                                 
                                            If InStr(employeeIDsPgrp, "OR mid = " & tjkValue) = 0 Then
                                                employeeIDsPgrp = employeeIDsPgrp & " OR mid = " & objDR2("medarbejderId")
                                            End If
                                
                                        End While
                                        objDR2.Close()
                       
                                    Next
        
                          
                            
                            
                       
                        
                                Else
                                    employeeIDsPgrp = " AND mid <> 0"
                                End If
                    
                    
                            Else
                                employeeIDsPgrp = " AND mid = 0"
                            End If
                    
                        End If
                
                
                
                        If aboMtyper <> "0" And aboMtyper <> "-1" Then
                   
                            If employeeIDsTyp.Length > 0 Then
                                employeeIDs = " AND (mid = 0 " & employeeIDsTyp & ")"
                            Else
                                employeeIDs = " AND (mid = 0) "
                            End If
                    
                        Else
                    
                            If employeeIDsTyp.Length > 0 Then
                                employeeIDs = employeeIDsTyp
                            Else
                                employeeIDs = " AND (mid = 0) "
                            End If
                    
                        End If
                
                
                        If aboMtyper = "-1" Then 'abonner via projektgrupper
                
                            If aboPgrp <> "0" And aboPgrp <> "-1" Then
                                If employeeIDsPgrp.Length > 0 Then
                                    employeeIDs = " AND (mid = 0 " & employeeIDsPgrp & ")"
                                Else
                                    employeeIDs = " AND (mid = 0)"
                                End If
                        
                            Else
                                If employeeIDsPgrp.Length > 0 Then
                                    employeeIDs = employeeIDsPgrp
                                Else
                                    employeeIDs = " AND (mid = 0)"
                                End If
                            End If
                    
                        End If
                
                    
               
                    End If '-2
                    
                
                    If employeeIDs = "" Then
                        employeeIDs = " AND (mid = 0)"
                    Else
                        employeeIDs = employeeIDs
                    End If
                
                    Dim strSQLm As String = "SELECT mid, mnavn, mnr, init , normtimer_man, normtimer_tir, normtimer_ons, normtimer_tor, normtimer_fre, normtimer_lor, normtimer_son "
                    strSQLm = strSQLm & " FROM medarbejdere AS m "
                    strSQLm = strSQLm & " LEFT JOIN medarbejdertyper AS mt ON (mt.id = m.medarbejdertype) WHERE mansat = 1 " & employeeIDs & " GROUP BY mid ORDER BY mnavn"
             
        
        
                    objCmd = New OdbcCommand(strSQLm, objConn)
                    objDR2 = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                    While objDR2.Read() = True
            
                
                        normTimer = Replace(FormatNumber(objDR2("normtimer_man") + objDR2("normtimer_tir") + objDR2("normtimer_ons") + objDR2("normtimer_tor") + objDR2("normtimer_fre") + objDR2("normtimer_lor") + objDR2("normtimer_son"), 2), ".", ",")
                
                        writer.Write(objDR2("mnavn") & ";" & objDR2("mnr") & ";" & objDR2("init") & ";" & normTimer & ";")
                
                
                        '***** L�n timer ***'
                        lTim = 0
                        lTimWrt = 0
                        lTim_addThis = 0
                        '-->Dim strSQLlt As String = "SELECT minutter FROM login_historik "
                        'strSQLlt =+ " LEFT JOIN stempelur AS s ON (s.id = stempelurindstilling) WHERE mid = " & objDR2("mid") & " AND (dato BETWEEN '" & startDatoSQL & "' AND '" & slutDatoSQL & "') GROUP BY mid"
                    
                        Dim strSQLlt As String = "SELECT minutter AS minutter, faktor, minimum FROM login_historik AS lh "
                        strSQLlt = strSQLlt & " LEFT JOIN stempelur AS s ON (s.id = lh.stempelurindstilling) "
                        strSQLlt = strSQLlt & " WHERE mid = " & objDR2("mid") & " AND (lh.dato BETWEEN '" & startDatoSQL & "' AND '" & slutDatoSQL & "') AND stempelurindstilling <> -1"
        
        
                        objCmd = New OdbcCommand(strSQLlt, objConn)
                        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                        While objDR.Read() = True
                            
                            If IsDBNull(objDR.Item("minutter")) Or IsDBNull(objDR.Item("faktor")) Or IsDBNull(objDR.Item("minimum")) Then
                                lTim_addThis = 0
                            Else
                                lTim_addThis = (objDR.Item("minutter") / 1 * objDR.Item("faktor") / 1)
                            
                                If lTim_addThis >= objDR.Item("minimum") Then
                                    lTim_addThis = lTim_addThis
                                Else
                                    lTim_addThis = objDR.Item("minimum")
                                End If
                            End If
                        
                        
                     
                         
                            lTim = lTim + (lTim_addThis / 1)
                    
                       
                            
                            lTimWrt = 1
                        End While
            
                        objDR.Close()
                    
                    
                        '** Pauser **'
                        lTim_pau = 0
                        lTim_pau_addThis = 0
                    
                        Dim strSQLlt_pau As String = "SELECT minutter AS minutter FROM login_historik WHERE mid = " & objDR2("mid") & " AND (dato BETWEEN '" & startDatoSQL & "' AND '" & slutDatoSQL & "') AND stempelurindstilling = -1"
        
        
                        objCmd = New OdbcCommand(strSQLlt_pau, objConn)
                        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                        While objDR.Read() = True
                    
                            If IsDBNull(objDR.Item("minutter")) Then
                                lTim_pau_addThis = 0
                            Else
                                lTim_pau_addThis = objDR.Item("minutter") 'Replace(Replace(FormatNumber(CType(objDR.Item("minutter") / 60, String), 2), ",", ""), ".", ",")
                            End If
                          
                            lTim_pau = (lTim_pau / 1 + lTim_pau_addThis / 1) '& " XX " '(lTim_pau / 1 + (lTim_pau_addThis / 1))
                            lTimWrt = 1
                        
                        End While

            
                        objDR.Close()
                        
                        
                    
                        
                    
                        If lTimWrt = 0 Then
                            writer.Write(";")
                        Else
                        
                            'lTim = Replace(Replace(FormatNumber(CType((lTim / 60) - (lTim_pau / 60), String), 2), ",", ""), ".", ",")
                            lTim = FormatNumber(CType((lTim / 60) - (lTim_pau / 60), String), 2)
                        
                            expTxt = lTim & ";"
                            writer.Write(expTxt)
                        
                            'lTim = Replace(Replace(FormatNumber(CType(lTim, String), 2), ",", ""), ".", ",")
                        End If
                    
                    
                    
                    
                        '** l�n timer �TD ****'
                        If show_atd = "1" Then
                        
                            Dim lTim_atd_Wrt As String = 0
                            lTim_atd = 0
                            lTim_addThis = 0
                            Dim strSQLlt_atd As String = "SELECT minutter AS minutter, faktor, minimum FROM login_historik AS lh "
                            strSQLlt_atd = strSQLlt_atd & " LEFT JOIN stempelur AS s ON (s.id = lh.stempelurindstilling) "
                            strSQLlt_atd = strSQLlt_atd & " WHERE mid = " & objDR2("mid") & " AND (lh.dato BETWEEN '" & Year(startDatoAtDSQL) & "/" & Month(startDatoAtDSQL) & "/" & Day(startDatoAtDSQL) & "' AND '" & slutDatoSQL & "') AND stempelurindstilling <> -1"
        
        
                            objCmd = New OdbcCommand(strSQLlt_atd, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                            While objDR.Read() = True
                            
                            
                                If IsDBNull(objDR.Item("minutter")) Or IsDBNull(objDR.Item("faktor")) Or IsDBNull(objDR.Item("minimum")) Then
                                    lTim_addThis = 0
                                Else
                                    lTim_addThis = (objDR.Item("minutter") / 1 * objDR.Item("faktor") / 1)
                                
                                
                                    If lTim_addThis >= objDR.Item("minimum") Then
                                        lTim_addThis = lTim_addThis
                                    Else
                                        lTim_addThis = objDR.Item("minimum")
                                    End If
                            
                                End If
                           
                            
                           
                         
                                lTim_atd = lTim_atd + (lTim_addThis / 1) 'Replace(Replace(FormatNumber(CType(objDR.Item("minutter") / 60, String), 2), ",", ""), ".", ",")
                            
                                lTim_atd_Wrt = 1
                            End While

                        
                      
            
                            objDR.Close()
                        
                        
                            '** Pauser **'
                            lTim_atd_pau = 0
                            lTim_pau_addThis = 0
                            Dim strSQLlt_atd_pau As String = "SELECT minutter AS minutter FROM login_historik WHERE mid = " & objDR2("mid") & " AND (dato BETWEEN '" & Year(startDatoAtDSQL) & "/" & Month(startDatoAtDSQL) & "/" & Day(startDatoAtDSQL) & "' AND '" & slutDatoSQL & "') AND stempelurindstilling = -1 AND minutter IS NOT NULL"
        
        
                            objCmd = New OdbcCommand(strSQLlt_atd_pau, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                            While objDR.Read() = True
                    
                            
                                If IsDBNull(objDR.Item("minutter")) Then
                                    lTim_pau_addThis = 0
                                Else
                                    lTim_pau_addThis = objDR.Item("minutter") 'Replace(Replace(FormatNumber(CType(objDR.Item("minutter") / 60, String), 2), ",", ""), ".", ",")
                                End If
                                
                            
                        
                                lTim_atd_pau = (lTim_atd_pau / 1 + lTim_pau_addThis / 1) '& " XX " '(lTim_pau / 1 + (lTim_pau_addThis / 1))
                            
                                lTim_atd_Wrt = 1
                            End While

            
                            objDR.Close()
                        
                            If lTim_atd_Wrt <> 0 Then
                                'lTim_atd = 0
                                lTim_atd = Replace(Replace(FormatNumber(CType(lTim_atd / 60 - (lTim_atd_pau / 60), String), 2), ",", ""), ".", ",")
                            Else
                                lTim_atd = 0
                            End If
                        
                        
                        End If 'show ATD
                
                
                 
                

                        '*** Real timer ***'
                        realTimer = 0
                        Dim strSQLext As String = "SELECT sum(timer) AS sumtimer, tmnavn FROM timer WHERE tmnr = " & objDR2("mid") & akttyper_realhours
                        strSQLext = strSQLext & "AND (tdato BETWEEN '" & startDatoSQL & "' AND '" & slutDatoSQL & "') GROUP BY tmnr"
        
                    
                        objCmd = New OdbcCommand(strSQLext, objConn)
                        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                        If objDR.Read() = True Then

                            expTxt = objDR.Item("sumtimer") & ";" 'Replace(Replace(FormatNumber(CType(objDR.Item("sumtimer"), String), 2), ",", ""), ".", ",") & ";"
                            'expTxt = Replace(FormatNumber(CType(objDR.Item("sumtimer"), String), 2), ".", ",") & ";"
                            writer.Write(expTxt)
                        
                            realTimer = objDR.Item("sumtimer")
                        End If
            
                        objDR.Close()
                    
                        If realTimer = 0 Then
                            writer.Write(";")
                        End If
                    
                  
                        'bal_norm_real = ((Replace(realTimer, ",", ".") / 1) - (normTimer / 100))
                        bal_norm_real = (realTimer / 1) - (normTimer / 1)
                        'bal_norm_real = Replace(Replace(FormatNumber(CType(bal_norm_real, String), 2), ",", ""), ".", ",")
                        bal_norm_real = FormatNumber(CType(bal_norm_real, String), 2)
                    
                    
                   
                    
                    
                        'bal_norm_lontimer = ((Replace(lTim, ",", ".") / 1) - (normTimer / 100))
                        bal_norm_lontimer = (lTim / 1) - (normTimer / 1)
                        'bal_norm_lontimer = Replace(Replace(FormatNumber(CType(bal_norm_lontimer, String), 2), ",", ""), ".", ",")
                        bal_norm_lontimer = FormatNumber(CType(bal_norm_lontimer, String), 2)
                    
                    
                
                        '*** Fakturerbare timer ***'
                        fakBareReal = 0
                    
                        Dim strSQLext2 As String = "SELECT sum(timer) AS sumtimerF, tmnavn FROM timer WHERE tmnr = " & objDR2("mid") & akttyper_invoiceable
                        strSQLext2 = strSQLext2 & "AND (tdato BETWEEN '" & startDatoSQL & "' AND '" & slutDatoSQL & "') GROUP BY tmnr"
        
        
                        objCmd = New OdbcCommand(strSQLext2, objConn)
                        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                        If objDR.Read() = True Then

                            'expTxt = Replace(Replace(FormatNumber(CType(objDR.Item("sumtimerF"), String), 2), ",", ""), ".", ",") & ";"
                            expTxt = objDR.Item("sumtimerF") & ";"
                            
                            writer.Write(expTxt)
                        
                            'fakBareReal = Replace(Replace(FormatNumber(CType(objDR.Item("sumtimerF"), String), 2), ",", ""), ".", ",")
                            fakBareReal = objDR.Item("sumtimerF")
                    
                        End If
            
                        objDR.Close()
                    
                        If fakBareReal = 0 Then
                            writer.Write(";")
                        End If
                    
                        effektiv_proc = 0
                        If fakBareReal <> 0 And lTim <> 0 Then
                            effektiv_proc = (fakBareReal / lTim) * 100
                            effektiv_proc = Replace(Replace(FormatNumber(CType(effektiv_proc, String), 0), ",", ""), ".", ",") 'fakBareReal &"/"& lTim & "="& 
                        End If
                        
                        
                        Dim strSQLextE1 As String
                        If lto = "dencker" Then
                     
                            '*** E1 ***'
                            e1Timer = 0
                            strSQLextE1 = "SELECT sum(timer) AS sumtimer, tmnavn FROM timer WHERE tmnr = " & objDR2("mid") & " AND tfaktim = 90 "
                            strSQLextE1 = strSQLextE1 & "AND (tdato BETWEEN '" & startDatoSQL & "' AND '" & slutDatoSQL & "') GROUP BY tmnr"
        
                       
                        
                            objCmd = New OdbcCommand(strSQLextE1, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                            If objDR.Read() = True Then

                                'expTxt = Replace(Replace(FormatNumber(CType(objDR.Item("sumtimer"), String), 2), ",", ""), ".", ",") & ";"
                                expTxt = objDR.Item("sumtimer") & ";"
                                writer.Write(expTxt)
                        
                                e1Timer = objDR.Item("sumtimer")
                        
                            End If
            
                            objDR.Close()
                    
                            If e1Timer = 0 Then
                                writer.Write(";")
                            End If
                        
                        End If
                    
                    
                    
                    
                    
                        If show_atd = "1" Then
                        
                            '*** Fakturerbare timer �TD ***'
                            fakBareReal_atd = 0
                    
                            Dim strSQLext2_atd As String = "SELECT sum(timer) AS sumtimerF, tmnavn FROM timer WHERE tmnr = " & objDR2("mid") & akttyper_invoiceable
                            strSQLext2_atd = strSQLext2_atd & "AND (tdato BETWEEN '" & Year(startDatoAtDSQL) & "/" & Month(startDatoAtDSQL) & "/" & Day(startDatoAtDSQL) & "' AND '" & slutDatoSQL & "') GROUP BY tmnr"
        
        
                            objCmd = New OdbcCommand(strSQLext2_atd, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                            If objDR.Read() = True Then

                            
                                'fakBareReal_atd = Replace(Replace(FormatNumber(CType(objDR.Item("sumtimerF"), String), 2), ",", ""), ".", ",")
                                fakBareReal_atd = objDR.Item("sumtimerF")
                         
                    
                            End If
            
                            objDR.Close()
                    
                            effektiv_proc_atd = 0
                            If fakBareReal_atd <> 0 And lTim_atd <> 0 Then
                                effektiv_proc_atd = (fakBareReal_atd / lTim_atd) * 100
                                effektiv_proc_atd = Replace(Replace(FormatNumber(CType(effektiv_proc_atd, String), 0), ",", ""), ".", ",")
                            End If
                        
                        End If
                    
                    
                        '*** Syg + Barn syg ***'
                        sygTimer = 0
                        Dim strSQLext3 As String = "SELECT sum(timer) AS sumtimerS, tmnavn FROM timer WHERE tmnr = " & objDR2("mid") & " AND (tfaktim = 20 OR tfaktim = 21)"
                        strSQLext3 = strSQLext3 & "AND (tdato BETWEEN '" & startDatoSQL & "' AND '" & slutDatoSQL & "') GROUP BY tmnr"
        
        
                        objCmd = New OdbcCommand(strSQLext3, objConn)
                        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                        If objDR.Read() = True Then

                            'expTxt = Replace(Replace(FormatNumber(CType(objDR.Item("sumtimerS"), String), 2), ",", ""), ".", ",") & ";"
                            expTxt = objDR.Item("sumtimerS") & ";"
                            writer.Write(expTxt)
                        
                            'sygTimer = Replace(Replace(FormatNumber(CType(objDR.Item("sumtimerS"), String), 2), ",", ""), ".", ",")
                            sygTimer = objDR.Item("sumtimerS")
                    
                        End If
            
                        objDR.Close()
                    
                        If sygTimer = 0 Then
                            writer.Write(";")
                        End If
                
                
                        '*** Ferie + Feriefridage afholdt / afholdt u. l�n / 1 maj timer ***'
                        ferieTimer = 0
                        Dim strSQLext4 As String = "SELECT sum(timer) AS sumtimerF, tmnavn FROM timer WHERE tmnr = " & objDR2("mid") & " AND (tfaktim = 13 OR tfaktim = 14 OR tfaktim = 19 OR tfaktim = 25)"
                        strSQLext4 = strSQLext4 & "AND (tdato BETWEEN '" & startDatoSQL & "' AND '" & slutDatoSQL & "') GROUP BY tmnr"
        
        
                        objCmd = New OdbcCommand(strSQLext4, objConn)
                        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                        If objDR.Read() = True Then

                            'expTxt = Replace(Replace(FormatNumber(CType(objDR.Item("sumtimerF"), String), 2), ",", ""), ".", ",") & ";"
                            expTxt = objDR.Item("sumtimerF") & ";"
                            writer.Write(expTxt)
                        
                            ferieTimer = objDR.Item("sumtimerF")
                        
                        End If
            
                        objDR.Close()
                    
                    
                        If ferieTimer = 0 Then
                            writer.Write(";")
                        End If
                
                        '*** Anden frav�r ***'
                        '** Afspadsering, Flex, sundhed, L�ge, Omsorgsdag, Senior 
                        andreTimer = 0
                        Dim strSQLext5 As String = "SELECT sum(timer) AS sumtimerA, tmnavn FROM timer WHERE tmnr = " & objDR2("mid") & " AND (tfaktim = 7 OR tfaktim = 8 OR tfaktim = 22 OR tfaktim = 23 OR tfaktim = 24 OR tfaktim = 31)"
                        strSQLext5 = strSQLext5 & "AND (tdato BETWEEN '" & startDatoSQL & "' AND '" & slutDatoSQL & "') GROUP BY tmnr"
        
        
                        objCmd = New OdbcCommand(strSQLext5, objConn)
                        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                        If objDR.Read() = True Then

                            'expTxt = Replace(Replace(FormatNumber(CType(objDR.Item("sumtimerA"), String), 2), ",", ""), ".", ",") & ";"
                            expTxt = objDR.Item("sumtimerA") & ";"
                            writer.Write(expTxt)
                        
                            'andreTimer = Replace(Replace(FormatNumber(CType(objDR.Item("sumtimerA"), String), 2), ",", ""), ".", ",")
                            andreTimer = objDR.Item("sumtimerA")
                        
                        End If
            
                        objDR.Close()
        
                        If andreTimer = 0 Then
                            writer.Write(";")
                        End If
                    
                        
                        Select Case lto
                            Case "cst", "kejd_pb", "outz"
                                writer.Write(bal_norm_lontimer & ";")
                            Case Else
                                writer.Write(bal_norm_real & ";")
                        End Select
                        
                    
                        If show_atd = "1" Then
                        writer.Write(effektiv_proc & " %;" & effektiv_proc_atd & " %;")
                    End If
                    
                        
                        
                        
                        '** �TD = 2 REAL timer �TD
                        If show_atd = "2" Then
                    
                            '*** Real timer �TD ***'
                            realTimerAtd = 0
                            Dim strSQLrealAtd As String = "SELECT sum(timer) AS sumtimer, tmnavn FROM timer WHERE tmnr = " & objDR2("mid") & akttyper_realhours
                            strSQLrealAtd = strSQLrealAtd & "AND (tdato BETWEEN '" & Year(startDatoAtDSQL) & "/" & Month(startDatoAtDSQL) & "/" & Day(startDatoAtDSQL) & "' AND '" & slutDatoSQL & "') GROUP BY tmnr"
        
                            objCmd = New OdbcCommand(strSQLrealAtd, objConn)
                            objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                            If objDR.Read() = True Then

                                expTxt = objDR.Item("sumtimer") & ";"
                                writer.Write(expTxt)
                        
                                realTimerAtd = objDR.Item("sumtimer")
                            End If
            
                            objDR.Close()
                    
                            If realTimerAtd = 0 Then
                                writer.Write(";")
                            End If
                    
                        
                            ugeNrLastWeek = DatePart("ww", slutDato, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
                        
                            'bal_norm_real = ((Replace(realTimer, ",", ".") / 1) - (normTimer / 100))
                            bal_norm_realAtd = (realTimerAtd / 1) - (normTimer * ugeNrLastWeek / 1)
                            'bal_norm_real = Replace(Replace(FormatNumber(CType(bal_norm_real, String), 2), ",", ""), ".", ",")
                            bal_norm_realAtd = FormatNumber(CType(bal_norm_real, String), 2)
                        
                        
                             writer.Write(bal_norm_realAtd & ";")
                        
                        End If 'show ATD
                        
                        
                        
                       
                  
            
                    
                        ugestatus = 0
                        ugegodkendt = 0
                        '**Uge afsluttet / godkendt '****'
                        Dim strSQLu As String = "SELECT status, ugegodkendt FROM ugestatus WHERE mid = " & objDR2("mid") & " AND uge = '" & slutDatoSQL & "' GROUP BY mid"
                        '** Uge altid = s�ndag i ugen ***'
        
                        objCmd = New OdbcCommand(strSQLu, objConn)
                        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                        If objDR.Read() = True Then
                        
                            ugestatus = 1
                            ugegodkendt = objDR.Item("ugegodkendt")
                        
                            expTxt = "Ja;" & ugegodkendt & ";"
                            writer.Write(expTxt)
                        
                      
                        
                        End If
            
                        objDR.Close()
                    
                    
                        If ugestatus = 0 Then
                            writer.Write(";;")
                        End If
                    
                    
                        '************** projektgrupper ****************
                    
                        If InStr(lto, "epi") Then
                        
                   
                    
                            Dim strSQLpgrpmedlemaf As String = "SELECT p.id AS pgrpid, p.navn as prognavn FROM progrupperelationer AS pr LEFT JOIN projektgrupper AS p ON (p.id = pr.projektgruppeId) " _
                            & " WHERE pr.medarbejderId  = " & objDR2("mid") & " AND p.id <> 10 GROUP BY pr.projektgruppeId ORDER BY p.id DESC"
                        
                            
                        employeeIDsPgrpNavn = ""
                                
                        objCmd = New OdbcCommand(strSQLpgrpmedlemaf, objConn)
                        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                        While objDR.Read() = True
                                
                         
                            employeeIDsPgrpNavn = employeeIDsPgrpNavn & objDR("prognavn") & ";"
                        
                        End While
                        objDR.Close()
                    
                          writer.Write(employeeIDsPgrpNavn)
                        
                        End If
                        
              
                  
                  
                    
                    
                    writer.WriteLine("")
                
                End While
            
                    objDR2.Close()
        
                
            
                End Using
        

            
          


                Dim report2 As New ozreportcls()
                report2.Email = modtEmail
                report2.Name = modtName
                report2.FileName = flname
                report2.Folder = lto
            

                lstRet.Add(report2)
               

                'Dim report1 As New ozreportcls()
                'report1.Email = "support@outzource.dk"
                'report1.Name = modtName
                'report1.FileName = flname
                'report1.Folder = lto '

                'lstRet.Add(report2)

               
                objConn.Close()
                'Next
            End While
            'Loop
        objDR_admin.Close()
        objConn_admin.Close()
        
       
        
            Return lstRet
          
    End Function

End Class


Public Class ozreportcls



Public Folder As String
Public FileName As String
Public Email As String
Public Name As String


Public Sub ozreportcls()

End Sub



End Class