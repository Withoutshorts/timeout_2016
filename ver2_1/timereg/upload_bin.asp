<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/header_hvd_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->

<%
    if request("matUpload") <> 0 then
        matUpload = 1
        matId = request("matId")
        'matJobid = request("matJobid")
        thisfile = request("thisfile")

        if thisfile = "expenceweb" then
            '"materialerudlaeg.asp?FM_sog_jobnavn="&request("FM_sog_jobnavn")&"&FM_personlig="&request("FM_personlig")&"&FM_fromdateSearch="&request("FM_fromdateSearch")
            FM_sog_jobnavn = request("FM_sog_jobnavn")
            FM_personlig = request("FM_personlig")
            FM_fromdateSearch = request("FM_fromdateSearch")
            FM_medarbejder = request("FM_medarbejder")
        end if

    else
        matUpload = 0 
    end if

    if request("newsuplaod") <> 0 then
        newsuplaod = 1
        newsid = request("newsid")
    else
        newsuplaod = 0
    end if
%>

<div style="position:absolute; top:20px; left:20px;">
<table>
<tr><td>
<b>F�lgende filer er lagt ud p� TimeOut web-serveren: sgokri <%=sogKri %></b><br><br> 
<%

    call TimeOutVersion() 

'  Variables
'  *********
   Dim Upload
   Dim file
   Dim intCount
   intCount=0
        
'  Object creation
'  ***************
   Set Upload = Server.CreateObject("Persits.Upload")

'  Upload '
'  ******
   Upload.Save
   
'  Maxsize '
'  ***********
   Upload.SetMaxSize 900000, True


   For each file In Upload.Files
  
        if matUpload <> 0 then
        strFileName = "mat_upload_" & matId & file.Ext
        else
        strFileName = file.FileName
        end if
  
      '  Save the files with his original names in a virtual path of the web server
      '  **************************************************************************
       	 'file.SaveAs("C:\www\timeout_xp\wwwroot\ver2_1\inc\upload\"&lto&"\" & file.FileName) 
	 	 file.SaveAs("D:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\upload\"&lto&"\" & strFileName) '*** Denne sti skal �ndres!
		 'file.SaveAs("wwwroot/outsite/upload/" & file.FileName)
		 
      '  Display the properties of the current file
      '  ******************************************
         'Response.Write("Name = " & file.Name & "<BR>")
      	 
         'Response.Write("FileName = " & file.FileName & "<BR>")
		 if file.Ext = ".gif" OR file.Ext = ".GIF" OR file.Ext = ".jpg" OR file.Ext = ".JPG"  OR file.Ext = ".png" OR file.Ext = ".PNG" then
		 Response.Write("<img src='../inc/upload/"&lto&"/" & strFileName & "' width=150 height=100><BR>")
		 Response.write("<br><b>"& strFileName &"</b><br><br>Billede her vist som: 150*100 px.")
		 else
		 Response.write("<br><b>"& strFileName &"</b>")
		 end if
        	
		 Response.Write("<br>Fil type og st�rrelse: <b>"& file.Ext & "</b> / <b> " & file.Size & "</b> bytes.<BR>")
			
		 intCount = intCount + 1
      
   Next


'*****************************'
'*** L�gger fil i database ***'
'*****************************'

	Set oConn = Server.CreateObject("ADODB.Connection")
	strConn2 = strConnect_DBConn 'strConnect 
	oConn.Open strConn2
	%>
<br>
<br>
<%
	
	usedate = year(now)&"/"&month(now)&"/"&day(now)
	x = 1
	For Each File in Upload.Files
		
        if matUpload <> 0 then
        strFileName = "mat_upload_" & matId & file.Ext
        else
        strFileName = file.FileName
        end if

		if x <= intCount then
		
		deletenewfileEntry = 0
			
			
			'*** Tjekker om fil alerede findes, Hvis den g�r skal den ikke l�gges i db igen **
			'*** Filer overskrives altid fysisk ****
			strSQL = "SELECT id FROM filer WHERE filnavn = '"& strFileName &"'"
			oRec.open strSQL, oConn, 3 
			if not oRec.EOF then
			deletenewfileEntry = 1
			end if
			oRec.close 
			
			if deletenewfileEntry = 0 then
				Set oCmd = server.createobject("ADODB.command")
				oCmd.activeconnection = strConnect_DBConn

                if matUpload <> 1 then
                    if newsuplaod = 1 then
                    oCmd.commandText = "UPDATE info_screen SET filnavn = '"& strFileName &"' WHERE id ="& newsid
                    else
				    oCmd.commandText = "UPDATE filer SET filnavn = '" & strFileName & "', editor = '"& session("user") &"', dato = '"& year(date) & "/"& month(date) & "/"& day(date) &"' WHERE oprses = 'x'"
                    end if
				else
                    oCmd.commandText = "UPDATE materiale_forbrug SET filnavn = '"& strFileName &"' WHERE id ="& matId 
                end if

                oCmd.commandType = 1
				oCmd.Execute
				set oCmd.activeconnection = Nothing
				
				Set File = nothing
				
				Set oCmd = server.createobject("ADODB.command")
				oCmd.activeconnection = strConnect_DBConn
				oCmd.commandText = "UPDATE filer SET oprses = '' WHERE oprses = 'x'"
				oCmd.commandType = 1
				oCmd.Execute
				
				set oCmd.activeconnection = Nothing
			end if
			
			
		end if
			
	x = x + 1
	Next
	
	'*** rydder op i db ****
	oConn.execute("DELETE FROM filer WHERE oprses = 'x'")
	
			'*** Finder Type og Id ***
			strSQL = "SELECT id, type, filnavn FROM filer WHERE id <> 0 ORDER BY id DESC"
			oRec.open strSQL, oConn, 3 
			if not oRec.EOF then
			filtype = oRec("type")
			filid = oRec("id")
			filnavn = oRec("filnavn")
			end if
			oRec.close 
%>
<br><br>

<%

if filtype <> 6 then '** 6 = Materiale billeder
	
	'if session("oprsesThis") = "open" then
	'Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
	'>
	'<a href="filer.asp?kundeid=<%=session("this_kid")>&jobid=<=session("this_jobid")>&nomenu=1" class=vmenu>Til filarkivet&nbsp;&nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a><br>
	'
	'else
	
	
	Response.Write("<script language=""JavaScript"">window.opener.document.getElementById(""FM_pic"").value = "& filid &";</script>")
	Response.Write("<script language=""JavaScript"">window.opener.document.getElementById(""FM_pic_navn"").value = '"& filnavn &"';</script>")
	response.write "<a href='#' onClick=""JavaScript:window.opener.location.reload();avascript:window.close();"" class=vmenu>[Luk]</a><br><br>"
	'Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
	
    'end if
	
	
else


'*** Upload af materiale billeder **' 
Response.Write("<script language=""JavaScript"">window.opener.document.getElementById(""FM_pic"").value = "& filid &";</script>")
Response.Write("<script language=""JavaScript"">window.opener.document.getElementById(""FM_pic_navn"").value = '"& filnavn &"';</script>")

end if

if deletenewfileEntry = 1 then%>
<font class=roed><b>!</b>&nbsp;En fil af samme navn som den uploadede fil <b>fandtes allerede</b> i filarkivet.<br>
Den eksisterene fil er blevet <b>overskrevet</b> med den nye fil, men den ligger stadigv�k i den oprindelige mappe.<br>
</font>
<%end if%>


<%
    if matUpload = 1 then        
        select case thisfile
        case "timetag_mobile"        
            response.Redirect "../timetag_web/expence.asp"
        case "expenceweb"
            response.Redirect "../to_2015/materialerudlaeg.asp?FM_sog_jobnavn="&FM_sog_jobnavn&"&FM_personlig="&FM_personlig&"&FM_fromdateSearch="&FM_fromdateSearch&"&FM_medarbejder="&FM_medarbejder
        end select
    end if
%>



<br><br>&nbsp;

</td></tr></table>
</div>