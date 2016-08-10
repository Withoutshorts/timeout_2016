<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
<script language=javascript>

function clwin(){
window.location.reload()
}

function setTimer() {
       setTimeout ("clwin()", 3600000); //millisekunder (er sat til 1 time)
}
</script>
	<title>TimeOut 2.1 Email Notifikations service</title>
</head>

<body onload="setTimer();">
Det er <b>TimeOut 2.1 email notifikations service</b>.<br>
Denne side må ikke lukkes ned. Det afsender mails hver time.<br>
Denne service er startet d. 5/11-2003 kl. 11.15<br><br>

<br>
Afsender notifikationer...

<%
Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject("ADODB.Recordset")

'strConnect = "mySQL_timeOut_intranet_BAK"



'Finder dags dato
tidspunkt = (DatePart("h", Now) + 1) & ":00:00"
tidspunkt30 = (DatePart("h", Now) + 1) & ":59:59"
datenow = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
usedate = Year(Now) & "/" & Month(Now) & "/" & Day(Now)

'session("afsendelser") = session("afsendelser") & "<br>Dato:" & datenow & " kl:" & tidspunkt &" - "& tidspunkt30 

Response.write "<br>Dato:" & datenow
Response.write "&nbsp;&nbsp;Tidsinterval:" & tidspunkt &" - "& tidspunkt30 &"<br>"

'Response.write "<br>Historik:"
'Response.write session("afsendelser")

numberofmails = 0

x = 1
numberoflicens = 49
For x = 1 To numberoflicens 
Select Case x
Case 1
strConnect = "driver={MySQL};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_intranet;" 
'Case 2
'strConnect = "driver={MySQL};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_sthaus;"  
'Case 3
'strConnect = "driver={MySQL};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_buying;"
'Case 4
'strConnect = "driver={MySQL};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_inlead;"
'Case 5
'strConnect = "driver={MySQL};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_netstrategen;"
'Case 6
'strConnect = "driver={MySQL};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_ravnit;"
Case 7
strConnect = "driver={MySQL};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_titoonic;"
'Case 8
'strConnect = "driver={MySQL};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_userneeds;"
Case 9
strConnect = "driver={MySQL};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_demo;"
'case 10
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_mezzo;"
case 11
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_skybrud;"
'case 12
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_gramtech;"
'case 13
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_cybervision;"
'case 14
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_ferro;"
'case 15
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_henton;"
'case 16
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_worldiq;"
case 17
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_lysta;"
case 18
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_execon;"
case 19
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_inclusive;"
case 20
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_kringit;"
case 21
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_kasters;"
'case 22
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_skousen;"
case 23
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_dencker;"
'case 24
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_firmaservice;"
'case 25
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_webmasters;"
case 26
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_proveno;"
case 27
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_oberg;"
'case 28
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bika;"
case 29
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_norma;"
case 30
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_stejle;"
case 31
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_novo_qdc;"
'case 32
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_simi;"
case 33
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_external;"
'case 34
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_netkoncept;"
case 35
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_perspektivait;"
case 36
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_mansoft;"
case 37
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bowe;"
case 38
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bminds"
case 39
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_workpartners"
case 40
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_gp"
case 41
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_optimizer4u"
case 42
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_zonerne"
case 43
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bolsjebutikken"
case 44
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_maintain"

case 45 '**** Start ***
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_start"

case 46
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_fyn-bo"
case 47
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_herbo"
case 48
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_radius"
case 49
strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_mansfield"
		

'Case 
'strConnect = "driver={MySQL};server=192.168.1.33;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_margin;"
'* Margin Media fejler
case else
strConnect = "nogo"
End Select

if strConnect <> "nogo" then


oConn.open strConnect

'Finder CRM aktioner.
strSQL = "SELECT DISTINCT(crmhistorik.id) AS crmaktionid, "_
&"crmhistorik.komm AS besk, crmhistorik.editor, crmhistorik.kundeid, crmdato,"_
&" kontaktemne, crmhistorik.navn, crmemne.navn AS emnenavn, crmstatus.navn AS"_
&" statusnavn, crmkontaktform.navn AS kontaktform, ikon,"_
&" crmhistorik.kontaktpers, crmklokkeslet, crmklokkeslet_slut,"_
&" kunder.kkundenavn, kunder.kid, kunder.telefon, kunder.adresse,"_
&" kunder.postnr, kunder.city, kunder.kkundenr, kunder.land, editorid,"_
&" medarbejdere.email, medarbejdere.mnavn, serialnb FROM crmhistorik,"_
&" aktionsrelationer LEFT JOIN crmemne ON (crmemne.id = kontaktemne) LEFT JOIN"_
&" crmstatus ON (crmstatus.id = status) LEFT JOIN crmkontaktform ON"_
&" (crmkontaktform.id = kontaktform) LEFT JOIN kunder ON (kunder.kid = kundeid)"_
&" LEFT JOIN medarbejdere ON (medarbejdere.mid = editorid) WHERE "_
&" crmhistorik.crmdato = '" & usedate & "' AND crmklokkeslet BETWEEN '"& tidspunkt &"' AND '" & tidspunkt30 & "' "

'Response.write strSQL &"<br><br>"
        

            oRec.open strSQL, oConn, 3
            While Not oRec.EOF

               
                kundeid = oRec.Fields("kundeid").value
				usremail = oRec.Fields("email").value
                usrname = oRec.Fields("mnavn").value
                crmemnenavn = oRec.Fields("emnenavn").value
                crmKpers = oRec.Fields("kontaktpers").value
                crmstarttid = oRec.Fields("crmklokkeslet").value
                crmsluttid = oRec.Fields("crmklokkeslet_slut").value
                statusnavn = oRec.Fields("statusnavn").value
                kontaktform = oRec.Fields("kontaktform").value
                kundenavn = oRec.Fields("kkundenavn").value
                kkundenr = oRec.Fields("kkundenr").value
                kundetlf = oRec.Fields("telefon").value
                kundeadr = oRec.Fields("adresse").value
                kundepostnr = oRec.Fields("postnr").value
               	kundeby = oRec.Fields("city").value
                kundeland = oRec.Fields("land").value
                strkomm = oRec.Fields("besk").value
                aktionsoskrift = oRec.Fields("navn").value


  				If DatePart("n", crmstarttid) <> 0 Then
                    usecrmstmin = DatePart("n", crmstarttid)
                Else
                    usecrmstmin = "00"
                End If

                If DatePart("n", crmsluttid) <> 0 Then
                    usecrmslutmin = DatePart("n", crmsluttid)
                Else
                    usecrmslutmin = "00"
                End If

                            'Sender notifikations mail
                            'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
                            '' Sætter Charsettet til ISO-8859-1
                            'Mailer.CharSet = 2
                            'Mailer.FromName = "Timeout2.1"
                            'Mailer.FromAddress = "support@outzource.dk"
                            'Mailer.RemoteHost = "webmail.abusiness.dk" '"pasmtp.tele.dk"
				            'Mailer.AddRecipient "" & usrname & "", "" & usremail & ""
                


                        Set myMail=CreateObject("CDO.Message")
                        myMail.From="timeout_no_reply@outzource.dk"

                        if len(trim(strEmail)) <> 0 then
                        myMail.To= ""& usrname &"<"& usremail &">"
                        end if


                        ' Selve teksten
                          if cint(kundeid) > 0 then
				            'Mailer.Subject = "Aktions notifikation!"

                            myMail.Subject="Aktions notifikation!"
                        
                            strBody = "Hej " & usrname & vbCrLf & vbCrLf 
                            strBody = strBody & "Aktions oplysninger: " & vbCrLf
                            strBody = strBody &"--------------------------------------------------" & vbCrLf & vbCrLf
                            strBody = strBody & "Aktions dato: " & datenow & " kl." & DatePart("h", crmstarttid) & "." & usecrmstmin & " - " & DatePart("h", crmsluttid) & "." & usecrmslutmin & vbCrLf
                            strBody = strBody & "Overskrift: " & aktionsoskrift & vbCrLf
                            strBody = strBody & "Emne: " & crmemnenavn & vbCrLf &"Kontaktform: " & kontaktform & vbCrLf & "Status: " & statusnavn & vbCrLf & vbCrLf & vbCrLf
                            strBody = strBody & "Kommentar:" & vbCrLf & strkomm & vbCrLf & vbCrLf
                            strBody = strBody &"--------------------------------------------------" & vbCrLf & vbCrLf & vbCrLf
				            strBody = strBody & "Kunde info: " & vbCrLf
                            strBody = strBody &"..................................................." & vbCrLf
				            strBody = strBody & "(" & kkundenr & ") " & kundenavn & vbCrLf
                            strBody = strBody & "Att: " & crmKpers & vbCrLf
                            strBody = strBody & kundeadr & vbCrLf
                            strBody = strBody & kundeby & " " & kundepostnr & vbCrLf
                            strBody = strBody & kundeland & vbCrLf
                            strBody = strBody & "Telefon: " & kundetlf & vbCrLf 
                            strBody = strBody &"..................................................." & vbCrLf & vbCrLf & vbCrLf
                            else
				            'Mailer.Subject = "Personlig note!"
                
                            myMail.Subject="Aktions notifikation!"
                        

                            strBody = "Hej " & usrname & vbCrLf & vbCrLf
				            strBody = strBody &"--------------------------------------------------" & vbCrLf & vbCrLf
				            strBody = strBody & strkomm & vbCrLf & vbCrLf
                            strBody = strBody &"--------------------------------------------------" & vbCrLf & vbCrLf & vbCrLf
				            end if
				
				            strBody = strBody & "Med venlig hilsen" & vbCrLf
                            strBody = strBody & "Timeout2.1 Email Notifikations Service"& vbCrLf
                            strBody = strBody & "http://www.outzource.dk" & vbCrLf & vbCrLf


					    myMail.TextBody = strBody                       
                        
                        myMail.Configuration.Fields.Item _
                        ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
                        'Name or IP of remote SMTP server
                                    
                                    if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                                       smtpServer = "webout.smtp.nu"
                                    else
                                       smtpServer = "formrelay.rackhosting.com" 
                                    end if
                    
                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/smtpserver")= smtpServer

                        'Server port
                        myMail.Configuration.Fields.Item _
                        ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
                        myMail.Configuration.Fields.Update

                        if len(trim(strEmail)) <> 0 then
                        myMail.Send

                            
                        set myMail=nothing




				
				          


                            'Mailer.BodyText = strBody

                            'Mailer.sendmail()
                            'Set Mailer = Nothing
				
				numberofmails = numberofmails + 1
				
                oRec.movenext
            Wend
           oRec.close
		   oConn.close
		
		
		end if 'nogo
        Next
		
		Set oRec = Nothing
		Set oConn = Nothing
%>
<br><br>
Der er afsendt <%=numberofmails%> mails i den sidste forsendelse.<br><br>

</body>
</html>
