<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
<!--#include file="../inc/connection/aktivedb_inc.asp"-->
<script language=javascript>

function clwin(){
window.location.reload()
}

function setTimer() {
       setTimeout ("clwin()", 3600000); //millisekunder (er sat til 1 time)
}
</script>
	<title>TimeOut CRM Reminder Service</title>
</head>

<body onload="setTimer();">
<h4>TimeOut CRM Reminder Service</h4>
Denne side m� ikke lukkes ned. Det afsender mails hver time.<br>
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
numberoflicens = 159
For x = 1 To numberoflicens 

call aktivedb(x)
'Response.Write strConnect_aktiveDB & "<br>"


if strConnect_aktiveDB <> "nogo" then


oConn.open strConnect_aktiveDB

'Finder CRM aktioner.
strSQL = "SELECT DISTINCT(crmhistorik.id) AS crmaktionid, "_
&"crmhistorik.komm AS besk, crmhistorik.editor, crmhistorik.kundeid, crmdato,"_
&" kontaktemne, crmhistorik.navn, crmemne.navn AS emnenavn, crmstatus.navn AS"_
&" statusnavn, crmkontaktform.navn AS kontaktform, ikon,"_
&" crmhistorik.kontaktpers, crmklokkeslet, crmklokkeslet_slut,"_
&" kunder.kkundenavn, kunder.kid, kunder.telefon, kunder.adresse,"_
&" kunder.postnr, kunder.city, kunder.kkundenr, kunder.land, editorid,"_
&" medarbejdere.email, medarbejdere.mnavn, serialnb FROM crmhistorik "_
&" LEFT JOIN medarbejdere ON (medarbejdere.mid = editorid)"_
&" LEFT JOIN aktionsrelationer ON (medarbid = medarbejdere.mid) "_
&" LEFT JOIN crmemne ON (crmemne.id = kontaktemne) "_
&" LEFT JOIN crmstatus ON (crmstatus.id = status) "_
&" LEFT JOIN crmkontaktform ON (crmkontaktform.id = kontaktform) LEFT JOIN kunder ON (kunder.kid = kundeid)"_
&" WHERE "_
&" crmhistorik.crmdato = '" & usedate & "' AND crmklokkeslet BETWEEN '"& tidspunkt &"' AND '" & tidspunkt30 & "' "

'Response.write strSQL &"<br><br>"
'response.Flush
        

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
                            '' S�tter Charsettet til ISO-8859-1
                            'Mailer.CharSet = 2
                            'Mailer.FromName = "Timeout2.1"
                            'Mailer.FromAddress = "support@outzource.dk"
                            'Mailer.RemoteHost = "webmail.abusiness.dk" '"pasmtp.tele.dk"
				            'Mailer.AddRecipient "" & usrname & "", "" & usremail & ""
                


                        Set myMail=CreateObject("CDO.Message")
                        myMail.From="timeout_no_reply@outzource.dk"

                        if len(trim(usremail)) <> 0 then
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

                        if len(trim(usremail)) <> 0 then
                        myMail.Send
                        end if

                            
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
