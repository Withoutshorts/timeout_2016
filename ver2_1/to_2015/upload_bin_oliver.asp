<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->


<div class="wrapper">
<div class="content">
<div class="container">
<div class="portlet">
<h3 class="portlet-title"><u>Upload</u></h3>
<div class="portlet-body">


<%if lto = "nt" then

else
%>

<b>Følgende filer er lagt ud på TimeOut web-serveren:</b><br><br>

<%

end if



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

   if lto = "nt" then

    deletenewfileEntry = 0

    For each file in Upload.Files
        
        
        
        'response.Write "test: " & file.FileName & "<br>"

        strSQL = "SELECT id FROM filer WHERE filnavn = '"& file.FileName &"'"
			oRec.open strSQL, oConn, 3 
			if not oRec.EOF then
			deletenewfileEntry = 1
			end if
			oRec.close

            'response.Write "filename: " & deletenewfileEntry
        
       
    Next

        if deletenewfileEntry = 1 then
        %>
            <div class="row" style="text-align:center">
                <div class="col-lg-4">
                    A file of the same name is already uploaded to the server, rename the image or choose another to continue
                </div>
            </div>
            <br />
            <div class="row" style="text-align:center">
                <div class="col-lg-2"><a class="btn btn-default btn-sm" role="button" href="Javascript:history.back()"><b>Back</b></a></div>
            </div>          
        <%
        
        else
       
         For each file In Upload.Files
  
  
      '  Save the files with his original names in a virtual path of the web server
      '  **************************************************************************
         'response.Write "lto: " & lto & file.FileName
       	 'file.SaveAs("C:\www\timeout_xp\wwwroot\ver3_99\inc\upload\"&lto&"\" & file.FileName) 
	 	 file.SaveAs("D:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\upload\"&lto&"\" & file.FileName) '*** Denne sti skal ændres!
		 'file.SaveAs("wwwroot/outsite/upload/" & file.FileName)
		 
      '  Display the properties of the current file
      '  ******************************************
         'Response.Write("Name = " & file.Name & "<BR>")
      	 
         'Response.Write("FileName = " & file.FileName & "<BR>")
		 if file.Ext = ".gif" OR file.Ext = ".GIF" OR file.Ext = ".jpg" OR file.Ext = ".JPG"  OR file.Ext = ".png" OR file.Ext = ".PNG" then
         %>
            <div class="row" style="text-align:center">
                <div class="col-lg-6">
                    <%
                        Response.Write("<img src='../inc/upload/"&lto&"/" & file.FileName & "' width=300 height=350><BR>")
                        Response.write("<br><b>"& file.FileName &"</b><br><br>The picture is shown as: 300*350 px.") 
                    %>
                </div>
            </div>
         <%
		 else
		 Response.write("<br><b>"& file.FileName &"</b>")
		 end if
        
         %>
           <!-- <div class="row" style="text-align:center">
                <div class="col-lg-6">
                    <%
                        Response.Write("<br>File type and size: <b>"& file.Ext & "</b> / <b> " & file.Size & "</b> bytes.<BR>") 
                    %>
                </div>
            </div> -->
         <%
		 
			
		 intCount = intCount + 1
      
         Next


        '*****************************'
        '*** Lægger fil i database ***'
        '*****************************'

	    Set oConn = Server.CreateObject("ADODB.Connection")
	    strConn2 = strConnect_DBConn 'strConnect 
	    oConn.Open strConn2


        x = 1
	    For Each File in Upload.Files
		
		if x <= intCount then
		
			
			if deletenewfileEntry = 0 then
				Set oCmd = server.createobject("ADODB.command")
				oCmd.activeconnection = strConnect_DBConn
				oCmd.commandText = "UPDATE filer SET filnavn = '" & file.FileName & "', editor = '"& session("user") &"', dato = '"& year(date) & "/"& month(date) & "/"& day(date) &"' WHERE oprses = 'x'"
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

            

        end if

        if deletenewfileEntry = 0 then
             %>
                <br />
                <div class="row" style="text-align:center">
                    <div class="col-lg-6">
                        Click <b>close</b> and <b>submit</b> to countinue
                    </div>
                </div>
                <br />
                <div class="row" style="text-align:center">
                    <div class="col-lg-6"><a class="btn btn-secondary btn-sm" href="javascript:window.open('','_self').close();">Close</a></div>
                </div>
             <%             
         end if
    
   else ' lto = nt 

   For each file In Upload.Files
  
  
      '  Save the files with his original names in a virtual path of the web server
      '  **************************************************************************
         'response.Write "lto: " & lto & file.FileName
       	 'file.SaveAs("C:\www\timeout_xp\wwwroot\ver3_99\inc\upload\"&lto&"\" & file.FileName) 
	 	 file.SaveAs("D:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\upload\"&lto&"\" & file.FileName) '*** Denne sti skal ændres!
		 'file.SaveAs("wwwroot/outsite/upload/" & file.FileName)
		 
      '  Display the properties of the current file
      '  ******************************************
         'Response.Write("Name = " & file.Name & "<BR>")
      	 
         'Response.Write("FileName = " & file.FileName & "<BR>")
		 if file.Ext = ".gif" OR file.Ext = ".GIF" OR file.Ext = ".jpg" OR file.Ext = ".JPG"  OR file.Ext = ".png" OR file.Ext = ".PNG" then
		 Response.Write("<img src='../inc/upload/"&lto&"/" & file.FileName & "' width=150 height=100><BR>")
		 Response.write("<br><b>"& file.FileName &"</b><br><br>Billede her vist som: 150*100 px.")
		 else
		 Response.write("<br><b>"& file.FileName &"</b>")
		 end if
        	
		 Response.Write("<br>Fil type og størrelse: <b>"& file.Ext & "</b> / <b> " & file.Size & "</b> bytes.<BR>")
			
		 intCount = intCount + 1
      
   Next


'*****************************'
'*** Lægger fil i database ***'
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
		
		if x <= intCount then
		
		deletenewfileEntry = 0
			
			
			'*** Tjekker om fil alerede findes, Hvis den gør skal den ikke lægges i db igen **
			'*** Filer overskrives altid fysisk ****
			strSQL = "SELECT id FROM filer WHERE filnavn = '"& file.FileName &"'"
			oRec.open strSQL, oConn, 3 
			if not oRec.EOF then
			deletenewfileEntry = 1
			end if
			oRec.close 
			
			if deletenewfileEntry = 0 then
				Set oCmd = server.createobject("ADODB.command")
				oCmd.activeconnection = strConnect_DBConn
				oCmd.commandText = "UPDATE filer SET filnavn = '" & file.FileName & "', editor = '"& session("user") &"', dato = '"& year(date) & "/"& month(date) & "/"& day(date) &"' WHERE oprses = 'x'"
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
Den eksisterene fil er blevet <b>overskrevet</b> med den nye fil, men den ligger stadigvæk i den oprindelige mappe.<br>
</font>
<%end if
    
    
end if 'lto = nt    
%>

<br><br>&nbsp;

</div>
</div>
</div>
</div>
</div>


<!--#include file="../inc/regular/footer_inc.asp"-->