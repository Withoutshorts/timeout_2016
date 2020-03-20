

<%
    
'strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_epi;"
 
strConnThis = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")
	
oConn.open strConnThis    
    
func = request("func")

select case func
case "update"

    rapporttype = request("FM_settype")

    strSQLUpd = "UPDATE rapport_abo SET rapporttype = "& rapporttype &" WHERE email = 'SNI@tiatechnology.com'"
    oConn.Execute strSQLUpd

    Response.redirect "reminders_settype.asp"

case else

end select
%>

<br /><br />
<div style="position:relative; left:40px; padding:20px;">
<h4>TimeOut Reminders Set Type:</h4>






<br /><br />


                <br />
                'WEEKLY
                <br />Rapportype:<b>9</b> Liste til LM <b>Sunday 23.59</b> med dem der ikke har afsluttet - Special costcenters 
                <br />Rapportype:<b>5</b> FIRST reminder ALL <b>monday 07.00</b> 
                <br />Rapportype:<b>6</b> SECOND reminder ALL in SnS (cost center 300, 320 & 330) <b>Monday 09.00</b>
                <br />Rapportype:<b>8</b> ALL Timesheet default completed <b>10:00</b>
                
                <br />Rapportype:<b>10</b> REMINDER ALL PM + LM who have not Approved WEEKLY & Project hours. Monday <b>13.00</b> - Special costcenters
                
                <br />Rapportype:<b>13</b> PM REMINDER ALL PM who have Not Approved Project hours. <b>Monday 13.00 </b>- Special costcenters
             
                <br />Rapportype:<b>12</b> Second reminder is sent to SnS mangers Deadline: <b>Monday 3pm</b>
                <br />Rapportype:<b>14</b> Second reminder is sent to SnS mangers Deadline: <b>Monday 3pm</b>
                <br />Rapportype: <b>11</b> Defalut godkend alle timer & uger. <b>monday 16:00</b>
                
             



    <br /><br /><br />
<%






x = 0
rapportype = 0
strSQL = "SELECT rapporttype, email, medid, id, lto FROM rapport_abo WHERE email = 'SNI@tiatechnology.com' ORDER BY id"
oRec.open strSQL, oConn, 3
While not oRec.EOF 
		
    Response.write "Current Rapportype: " & oRec("rapporttype")
	rapportype = oRec("rapporttype")

x = x + 1		
oRec.movenext
wend
oRec.close


		
%>
    <br /><hr />
    <form action="reminders_settype.asp?func=update" method="post">
Sæt Type på næste reminders: <input type="text" name="FM_settype" value="" placeholder="rapportype"  /> <input type="submit" value="Update" />
    <br />


</form>


<br /><br /><hr />
    
    <br /><br />
Der er ialt <b><%=x%></b> job på listen
<br>
    </div>

