
<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->



<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
%>
<!--#include file="../inc/regular/header_hvd_inc.asp"-->
<%


	'ekspTxt = request("FM_eksportdata")


	'txt1 = request.form("txt1")
	''txt2 = request.form("txt2")
	
	Dim BigTextArea

	For I = 1 To Request.Form("BigTextArea").Count
  		BigTextArea = BigTextArea & Request.Form("BigTextArea")(I)
	Next
	
	txt20 = request.form("txt20")
	txt = txt1 & BigTextArea & txt20
	
	ekspTxt = replace(txt, "xx99123sy#z", vbcrlf)
	
	datointerval = request("datointerval")
	
	call TimeOutVersion()

	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\materiale_stat_eksport.asp" then
					Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\matstat_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\matstat_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				else
					Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\matstat_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\matstat_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				end if
				
				
				
				file = "matstat_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				objF.WriteLine(ekspTxt)
				
	
	Response.redirect "../inc/log/data/"& file &""	
	%>
	<!--
	<div id="sindhold" style="position:absolute; left:20; top:20; height:600; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0">
	<tr>
    <td valign="top"><a href="../inc/log/data/<=file%>" class=vmenu>Data er nu eksporteret til denne fil.&nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a><br>
	<br>
	Gem filen ved at klikke p� filmenuen og herefter "gem som".<br>
	�bn det program du �nsker at hente txt filen ind i og importer.<br><br>
	
	</td>
	</tr></table>
	</div>-->
	
	
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->