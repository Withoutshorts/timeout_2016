<%
if len(request("key")) <> "0" then
'Inds�tter cookie
Response.Cookies("licenskey") = request("key")
Response.Cookies("licenskey").Expires = date + 365
end if

'ODBC
'Dim objXMLHTTP, objXMLDOM, i
'Opretter en instans af Microsoft.XMLHTTP, s� det er muligt at f� fat p� dokumentet
'Set objXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")
'Opretter en instans af Microsofts XML-parser, XMLDOM
'Set objXMLDOM = Server.CreateObject("Microsoft.XMLDOM")

if len(request("key")) <> "0" then
strlicenskey = request("key")
else
'Henter Cookie
strLicenskey = Request.Cookies("licenskey")
end if


select case strLicenskey
case "2.052-xxxx-B000" 'demo
	'Call objXMLHTTP.Open("GET", "https://outzource.dk/demo/conf_TOser_2_demo.xml", False)
	strConnThis = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_demo;"
	addnewclient = 1
	lto = "demo"
case "2.152-0416-B001" 'outzource
	'Call objXMLHTTP.Open("GET", "https://outzource.dk/outz/conf_TOser_2_outz.xml", False)
	strConnThis = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_intranet;"
	lto = "outz"
case "2.152-0416-B002" 'sthaus
	'Call objXMLHTTP.Open("GET", "https://outzource.dk/sthaus/conf_TOser_2_sthaus.xml", False)
	strConnThis = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_sthaus;"
	lto = "sthaus"
case "2.052-0423-A005" 'margin
	'Call objXMLHTTP.Open("GET", "https://outzource.dk/margin/conf_TOser_2_margin.xml", False)
	strConnThis = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_margin;"
	lto = "margin"
case "2.182-0416-B004" 'Buying
	'Call objXMLHTTP.Open("GET", "https://outzource.dk/buy/conf_TOser_2_buy.xml", False)
	strConnThis = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_buying;"
	lto = "buying"
case "2.052-2111-B010" 'Netstrategen
	'Call objXMLHTTP.Open("GET", "https://outzource.dk/net/conf_TOser_2_net.xml", False)
	strConnThis = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_netstrategen;"
	lto = "net"
'case "2.052-0705-B008" 'Revisor
	'Call objXMLHTTP.Open("GET", "https://outzource.dk/revisor/conf_TOser_2_rev.xml", False)
	'lto = "rev"
'case "2.152-2711-B012" 'Comone
	'Call objXMLHTTP.Open("GET", "https://outzource.dk/comone/conf_TOser_2_comone.xml", False)
	'lto = "comone" 
case "2.052-2711-B013" 'Userneeds
	'Call objXMLHTTP.Open("GET", "https://outzource.dk/usrneeds/conf_TOser_2_usrneeds.xml", False)
	strConnThis = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_userneeds;"
	lto = "usrneeds"
'case "2.052-0416-B006" 'WPS
	'Call objXMLHTTP.Open("GET", "https://outzource.dk/wps/conf_TOser_2_wps.xml", False)
	'lto = "wps"
'case "2.993-2701-B014" 'Voice
	'Call objXMLHTTP.Open("GET", "https://outzource.dk/voice/conf_TOser_2_voice.xml", False)
	'lto = "voice"
case "xxx" 'ktv udvikling
	if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\login.asp" then
	Call objXMLHTTP.Open("GET", "https://outzource.dk/outz/conf2_ktvudv.xml", False)
	else
	Call objXMLHTTP.Open("GET", "http://localhost/timeout_xp/inc/xml/conf3_ktvudv.xml", False)
	end if
	lto = "intranet"
case "2.013-1803-B015" 'ravnit
	lto = "ravnit"
	strConnThis = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_ravnit;"
	'Call objXMLHTTP.Open("GET", "https://outzource.dk/ravnit/conf_TOser_3_ravnit.xml", False)
case "2.053-0705-B016" 'inlead
	lto = "inlead"
	strConnThis = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_inlead;"
	'Call objXMLHTTP.Open("GET", "https://outzource.dk/inlead/conf_TOser_2_inlead.xml", False)
'case "2.053-2705-B017" 'itsupportpartner
	'lto = "itsuppart"
	'Call objXMLHTTP.Open("GET", "https://outzource.dk/itsuppart/conf_TOser_2_itsuppart.xml", False)
case "2.253-1906-B018" 'titoonic
	lto = "titoonic"
	strConnThis = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_titoonic;"
	'Call objXMLHTTP.Open("GET", "https://outzource.dk/titoonic/conf_TOser_2_titoonic.xml", False)
'case "2.153-1909-B019" 'revitax	
	'lto = "revitax"
	'Call objXMLHTTP.Open("GET", "https://outzource.dk/revitax/conf_TOser_2_revitax.xml", False)
case "2.053-1011-B020" 'Mezzo	
	lto = "mezzo"
	strConnThis = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_mezzo;"
	'Call objXMLHTTP.Open("GET", "https://outzource.dk/mezzo/conf_TOser_2_mezzo.xml", False)
case "2.153-2411-B021" 'Skybrud	
	lto = "skybrud"
	strConnThis = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_skybrud;"
	'Call objXMLHTTP.Open("GET", "https://outzource.dk/skybrud/conf_TOser_2_skybrud.xml", False)	
case "2.154-2301-B022" 'Gramtech
	lto = "gramtech"
	strConnThis = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_gramtech;"
	'Call objXMLHTTP.Open("GET", "https://outzource.dk/gramtech/conf_TOser_2_gramtech.xml", False)	
		
case else 'ktv udvikling (hvis der ikke er en cookie)
	strConnThis = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_demo;"
	lto = "intranet"
end select 

'objXMLHTTP.Send

'L�gger indholdet af dokumentet over i vores XML-parser objekt
'Set objXMLDOM = objXMLHTTP.ResponseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("modul")

'Henter indholdet af alle tags med navnet 'url'
'Set objDBConn = objXMLDOM.getElementsByTagName("dbconn")

strConn = "on"'objDBConn(0).text  'Connection
strA1 = "on"'objModuler(0).text   'Timeregistrering
strA2 = "on"'objModuler(1).text   'Medarbejder log
strA3 = "on"'objModuler(2).text   'Skift medarbejder 
strA4 = "on"'objModuler(3).text   'Se timereg for andre medarb.

strB1 = "on"'objModuler(4).text   'Joboversigt og Aktiviter. Mulighed for at lukke akt. 
strB2 = "on"'objModuler(5).text   'Projektgrupper
strB3 = "on"'objModuler(6).text   'Stamaktiviteter
strB4 = "on"'objModuler(7).text   'Jobplanner
strB5 = "on"'objModuler(8).text   'Km regnskab
strB6 = "on"'objModuler(9).text   'Udgifter paa job
strB7 = "on"'objModuler(10).text   'Ikke planlagt
strB8 = "on"'objModuler(11).text   'Upload af dokumenter
strB9 = "on"'objModuler(12).text   'Incident tool
strB10 = "on"'objModuler(13).text	'Aktivitetsskabeloner

strC1 = "on"'objModuler(14).text  'Kunder + filter og sog paa kunde
strC2 = "on"'objModuler(15).text  'Kunde login
strC3 = "on"'objModuler(16).text	 'Upload logoer paa kunder

strD1 = "on"'objModuler(17).text  'Medarbejdere
strD2 = "on"'objModuler(18).text  'Brugergrupper
strD3 = "on"'objModuler(19).text  'Medarbejdertyper

strE1 = "on"'objModuler(20).text  'Statistik og Fakturering + filter
strE2 = "on"'objModuler(21).text  'Joblog_z_b, joblog, top 5 job
strE3 = "on"'objModuler(22).text	 'Aars omsaetning
strE4 = "on"'objModuler(23).text  'Export
strE5 = "on"'objModuler(24).text	 'Faktura oversigt

strH1 = "on"'objModuler(25).text  'Kalender
strH2 = "on"'objModuler(26).text  'Firmaer
strH3 = "on"'objModuler(27).text	 'Aktions historik + sogefunktion
strH4 = "on"'objModuler(28).text  'Status
strH5 = "on"'objModuler(29).text  'Emner
strH6 = "on"'objModuler(30).text  'Kontaktform
strH7 = "on"'objModuler(31).text  'Outlook integration

'Set objXMLHTTP = Nothing
'Set objXMLDOM = Nothing


if len(strConnThis) <> 0 then
	
	strConnect = strConnThis
	
	Set oConn = Server.CreateObject("ADODB.Connection")
	Set oRec = Server.CreateObject ("ADODB.Recordset")
	Set oRec2 = Server.CreateObject ("ADODB.Recordset")
	Set oRec3 = Server.CreateObject ("ADODB.Recordset")
	Set oCmd = Server.CreateObject ("ADODB.Command")
	
	oConn.open strConnect
	
	sub closeDB
	'now close and clean up
			oRec.Close
			oConn.close
			
			Set oRec = Nothing
			set oConn = Nothing
	
	end sub

end if

%>
