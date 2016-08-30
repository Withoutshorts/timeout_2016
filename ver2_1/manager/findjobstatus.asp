
<br /><br />
<div style="position:relative; left:40px; padding:20px;">
<h4>TimeOut Jobopslag Epi:</h4>


<%
 if session("user") = "" then

  'errortype = 5
  'call showError(errortype)
	
   response.write "Du er ikke logget ind i TimeOut og kan derfor ikke bruge denne service.<br><br>Log ind i TimeOut og refresh derefter denne side."

   response.end
 end if


if len(trim(request("FM_sogjobnr"))) <> 0 then
    jobnr = trim(request("FM_sogjobnr"))
else
    jobnr = "-1"
end if
 %>



<br /><br />
<form>
Søg på jobnr:&nbsp;<input type="text" name="FM_sogjobnr" value="<%=jobnr %>" placeholder="Search" /> <input type="submit" value="Search" />
    <br />


</form>


<hr />
Jobliste: <br />

<%








strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_epi;"
Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")
	
oConn.open strConnThis

x = 0
strSQL = "SELECT id, jobnavn, jobnr, jobstatus FROM job WHERE jobnr LIKE '"& jobnr &"%' ORDER BY jobnr"
oRec.open strSQL, oConn, 3
While not oRec.EOF 
		
		Response.write oRec("jobnr") & " - " & oRec("jobnavn") & "&nbsp; Status: "

        select case oRec("jobstatus")
        case "1"
        jobstatusTxt = "Active"
        case "2"
        jobstatusTxt = "Passive"
        case "3"
        jobstatusTxt = "Quote"
        case "4"
        jobstatusTxt = "Re-view"
        case else
        jobstatusTxt = "Closed" 
        end select

       Response.write "<b>"& jobstatusTxt & "</b><br>"

x = x + 1		
oRec.movenext
wend
oRec.close


		
%>
Der er ialt <b><%=x%></b> job på listen
<br>
    </div>

