
<%
function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function


'*** Opdaterer timer, tfaktim til 0, 1, 2 ***
'strConnect = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_wps;"
strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_epi;"

'strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi_BAK;"
  
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=root;pwd=;database=timeout_intranet;"
'strConnect = "mySQL_timeOut_intranet"
Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")
	
oConn.open strConnect


strSQLjob = "SELECT id, jobnavn, jobnr FROM job WHERE (jobstatus = 2 OR jobstatus = 0 OR jobstatus = 4) AND jobnr > 14400 "'LIMIT 50" 'jobstatus <> 1

oRec2.open strSQLjob, oConn, 3
While not oRec2.EOF 


Response.write "<br><br><h4>"& oRec2("jobnavn") & " ("& oRec2("jobnr") &")</h4>"


strSQL = "SELECT id, navn FROM aktiviteter WHERE "_
&" (navn LIKE 'Tilb%' "_
&" OR navn LIKE 'Projektledelse%' "_
&" OR navn LIKE 'Desk%' "_
&" OR navn LIKE 'Forberedelse%' "_
&" OR navn LIKE 'Opsætning%' "_
&" OR navn LIKE 'Gennemførelse%' "_
&" OR navn LIKE 'E-survey%' "_
&" OR navn LIKE 'Tabelrapportering%' "_
&" OR navn LIKE 'Kvant%' "_
&" OR navn LIKE 'Kval%' "_
&" OR navn LIKE 'Afrapportering%' "_
&" OR navn LIKE 'CATI' "_
&" OR navn LIKE 'Formidling%') "_
&" AND (fakturerbar = 1 OR fakturerbar = 6) AND job = "& oRec2("id") &" ORDER BY sortorder, navn"



'&" OR navn LIKE 'Forberedelse af work%' "_
'&" OR navn LIKE 'Forberedelse af kvan%' "_

'&" OR navn LIKE 'Gennemførelse af wor%' "_

x = 0
nytIdGP0 = 0
nytIdGP1 = 0
nytIdGP2 = 0
nytIdGP3 = 0
nytIdGP4 = 0
nytIdGP5 = 0

oRec.open strSQL, oConn, 3
While not oRec.EOF 




'if (lastNavn) <> oRec("navn") then
Response.write ""& oRec("navn") & "("& oRec("id")  &")<br>"
'end if
        
        SELECT CASE trim(left(oRec("navn"), 8))
        case "Tilbud/s"
        nytnavnGP0 = "Tilbud / Salg"
        nytIdGP0 = oRec("id")
        statusGP0 = 1
                
                usegp = 0
        

        case "Projektl"
        nytnavnGP1 = "Projektledelse / Møder"
        nytIdGP1 = oRec("id")
        statusGP1 = 1
                
                usegp = 1

        case "Deskrese", "Forbered", "Opsætni"
                
                nytnavnGP2 = "Forberedelse / Desk / Spørgeguides"
                
                select case trim(left(oRec("navn"), 8))
                case "Deskrese"
                nytIdGP2 = oRec("id")
                statusGP2 = 1
                case else
                nytIdGP2 = nytId
                statusGP2 = 0
                end select    

                if nytIdGP2 <> 0 then
                nytIdGP2 = nytIdGP2
                else
                nytIdGP2 = oRec("id")
                end if

                usegp = 2

        case "Gennemfø", "Gennemfø", "E-survey", "CATI"

                nytnavnGP3 = "Interview / Grupper / Konference"
                
                select case trim(left(oRec("navn"), 20))
                case "Gennemførelse af kva"
                nytIdGP3 = oRec("id")
                statusGP3 = 1
                case else
                nytIdGP3 = nytIdGP3
                statusGP3 = 0
                end select  
                
                if nytIdGP3 <> 0 then
                nytIdGP3 = nytIdGP3
                else
                nytIdGP3 = oRec("id")
                end if

                usegp = 3  
        
        case "Tabelrap", "Kvant da", "Kval dat", "Afrappor"

                nytnavnGP4 = "Data / Register / Rapport"
                
                select case trim(left(oRec("navn"), 8))
                case "Tabelrap"
                nytIdGP4 = oRec("id")
                statusGP4 = 1
                case else
                nytIdGP4 = nytIdGP4
                statusGP4 = 0
                end select 

                 if nytIdGP4 <> 0 then
                nytIdGP4 = nytIdGP4
                else
                nytIdGP4 = oRec("id")
                end if

                 usegp = 4  
                
        case "Formidli"
        nytnavnGP5 = "Præsentation / Opfølgning"
        nytIdGP5 = oRec("id")                   
        statusGP5 = 1

        usegp = 5

        end select



        select case usegp
        case 0
        nytnavn = nytnavnGP0 
        nytId = nytIdGP0 
        status = statusGP0 
        case 1
        nytnavn = nytnavnGP1 
        nytId = nytIdGP1 
        status = statusGP1 
        case 2
        nytnavn = nytnavnGP2 
        nytId = nytIdGP2 
        status = statusGP2 
        case 3
        nytnavn = nytnavnGP3 
        nytId = nytIdGP3 
        status = statusGP3 
        case 4
        nytnavn = nytnavnGP4 
        nytId = nytIdGP4 
        status = statusGP4 
        case 5
        nytnavn = nytnavnGP5 
        nytId = nytIdGP5 
        status = statusGP5 
        end select




		if status = 1 then
        strSQLanavn = ", navn = '"& nytnavn &"'"
        else
		strSQLanavn = ""
	    end if
        
        strSQLa = "UPDATE aktiviteter SET aktstatus = "& status &" "& strSQLanavn &" , sortorder = "& usegp &" WHERE id = "& oRec("id") 


        'navn = '"& nytnavn &"'
        'select case right(x, 3)
        'case "500", "000"
        'Response.write "strSQLa: "& strSQLa & "<br>"
        'end select

        'oConn.execute(strSQLa)
        Response.flush
		
        strSQLt = "UPDATE timer SET taktivitetnavn = '"& nytnavn &"', taktivitetid = "& nytId &" WHERE taktivitetid = "& oRec("id") 
	    
        'select case right(x, 3)
        'case "500", "000"
        'Response.write "strSQLt: "& strSQLt & "<br><br>"
        'end select

        'oConn.execute(strSQLt)
        Response.flush

		'** LUK akt **'
		
        lastNavn = oRec("navn")

x = x + 1
oRec.movenext
wend
oRec.close


oRec2.movenext
wend
oRec2.close
		
%>
<br>
X: <%=x %>
Opdateringen er gennemført!
