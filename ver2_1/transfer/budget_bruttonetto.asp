<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/budget_func.asp"-->
<!--#include file="inc/helligdage_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	kl = request("kl")
	crmaktion = request("crmaktion")
	thisfile = "budget_bruttonetto"
	
	
	public flt
	function trimFld(flt)
	
	if len(flt) <> 0 then
	lenflt = len(flt)
	leftflt = left(flt, lenflt - 1)
	flt = leftflt
	end if
	
	end function    
	
	
	function top()%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call budgetmainmenu(1)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	if showonejob <> 1 then
		call budgettopmenu()
	end if
	%>
	</div>
	<div style="position:absolute; left:20px; top:132px;">
	<h3>Brutto/Netto Dage & Timer</h3>
	<% end function
	
	
	public thisAar, strNavn
	function vaelgAar()
	%>
	<form method=post action=budget_bruttonetto.asp>
	Vælg år: 
	
	<%
	if len(request("FM_aar")) <> 0 then
	vlgtAar = request("FM_aar")
	else
	vlgtAar = year(now)
	end if 
	%>
	
	
        <select id="FM_aar" name="FM_aar">
        <%
        for x = -5 to 5
        
        if x = -5 then
        thisAar = year(now) - 5
        else
        thisAar = cint(thisAar) + 1
        end if
        
        if cint(vlgtAar) = cint(thisAar) then
        arrCHK = "SELECTED"
        else 
        arrCHK = ""
        end if
        %>
            <option value="<%=thisAar %>" <%=arrCHK %>><%=thisAar %></option>
        
        <%
        next
        %>
        </select>
        &nbsp;<input id="Submit1" type="submit" value="Vis år >> " />
	</form>
	
	<b>Gemte brutto / netto tal:</b><br />
	
	<table cellpadding=2 cellspacing=1 bgcolor="#8cAAe6" width=200>
	<tr>
	<td><b>År</b></td>
	<td><b>Navn</b></td>
	<td>Slet?</td>
	</tr>
	<%
	thisAar = vlgtAar
	strNavn = ""
	
	strSQL = "SELECT id, navn, budget_aar FROM budget_bruttonetto_dt WHERE budget_aar = " & thisAar
	oRec.open strSQL, oConn, 3 
	x = 0
	while not oRec.EOF 
	
	    if cint(id) = oRec("id") then
	    bgThis = "#ffff99"
	    strNavn = oRec("navn")
	    else
	        select case right(x,1)
	        case 0,2,4,6,8
	        bgThis = "#ffffff"
	        case else
	        bgThis = "#D6DFf5"
	        end select
	    end if
	
	%>
	<tr bgcolor="<%=bgThis %>">
	<td><%=oRec("budget_aar") %></td>
	<td><a href="budget_bruttonetto.asp?func=red&id=<%=oRec("id") %>&FM_aar=<%=thisaar %>" class=vmenu><%=oRec("navn") %></a></td>
	<td align=center><a href="budget_bruttonetto.asp?func=slet&id=<%=oRec("id") %>&FM_aar=<%=thisaar %>" class=vmenu>
        <img src="../ill/slet_16.gif" border="0" /></a></td>
	</tr>
	<%
	x = x + 1
	oRec.movenext
	wend
	oRec.close
	
	if x = 0 then%>
	<tr bgcolor="#ffffff"><td colspan=3>(Ingen)</td></tr>
	<%end if %>
	</table>
	<br />
	
	<%if len(strNavn) <> 0 then %>
	<h4><a href="budget_bruttonetto.asp?func=opr&FM_aar=<%=thisaar %>"><img src="../ill/opretny.gif" border=0 alt="Opret ny" /></a>&nbsp;
	<%=strNavn%>
	            
	            <%if request("fromupdate") = "1" then
	            Response.Write " - <font color=darkred>Opdateret!</font>"
	            end if
	             %>
	</h4>
	<%end if 
	end function
	
	
	
	select case func
	case "slet"
	call top()
	
	slttxt = "<b>Slet Brutto/Netto Dage & Timer skema?</b><br />"_
	&"Skema kan ikke gen-skabes."
	slturl = "budget_bruttonetto.asp?func=sletja&id="&id&"&FM_aar="&request("FM_aar")
	
	slttxtalt = ""
    slturlalt = ""
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,210,100)
	
	

	
	case "sletja"
	
	strSQLdel = "DELETE FROM budget_bruttonetto_dt WHERE id = "& id
	oConn.execute(strSQLdel)
	
	Response.Redirect "budget_bruttonetto.asp?FM_aar="&request("FM_aar")
	
	case "dbopr", "reddb"
	
	strNavn = replace(request("FM_navn"), "'", "''")
	if len(trim(strNavn)) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	
	            errortype = 109
				call showError(errortype)
	
	Response.End 
	end if
	
	
	budgetAar = request("FM_aar")
	editor = session("user")
	
	'**** Brutto Dage ****'
	bdFlds = ""
	bruttoDageSQL = ""
	for x = 1 to 12 
	bdFlds = bdFlds & "bd_"&x&"," 
	call erDetInt(request("FM_bd_dage_"&x&"")) 
	
	if len(request("FM_bd_dage_"&x&"")) <> 0 AND isInt = 0 then
	        if func = "reddb" then
	        bruttoDageSQL = bruttoDageSQL & "bd_"&x&" = "& request("FM_bd_dage_"&x&"") &","
	        else
	        bruttoDageSQL = bruttoDageSQL & request("FM_bd_dage_"&x&"") &","
	        end if
	else
	        if func = "reddb" then
	        bruttoDageSQL = bruttoDageSQL & "bd_"&x&" = 0,"
	        else
	        bruttoDageSQL = bruttoDageSQL & "0,"
	        end if
	end if
	
	isInt = 0
	next
	
	call trimFld(bdFlds)
	call trimFld(bruttoDageSQL) 
	
	
	
	'**** Ferie fradrag ***
	ffFlds = ""
	frd_feSQL = ""
	for x = 1 to 12 
	ffFlds = ffFlds & "frd_ferie_"&x&","
	call erDetInt(request("FM_frd_fe_"&x&"")) 
	if len(request("FM_frd_fe_"&x&"")) <> 0 AND isInt = 0 then
	        
	        if func = "reddb" then
	        frd_feSQL = frd_feSQL & "frd_ferie_"&x&" = "& replace(request("FM_frd_fe_"&x&""),",",".") &","
	        else
	        frd_feSQL = frd_feSQL & replace(request("FM_frd_fe_"&x&""),",",".") &","
	        end if
	else
	        if func = "reddb" then
	        frd_feSQL = frd_feSQL & "frd_ferie_"&x&" = 0,"
	        else
	        frd_feSQL = frd_feSQL &"0,"
	        end if
	
	end if
	isInt = 0
	next
	
	call trimFld(ffFlds)
	call trimFld(frd_feSQL) 
	
	'**** Sygdom fradrag ***
	fsFlds = ""
	frd_sySQL = ""
	for x = 1 to 12 
	fsFlds = fsFlds & "frd_syg_"&x&"," 
	call erDetInt(request("FM_frd_sy_"&x&""))
	if len(request("FM_frd_sy_"&x&"")) <> 0 AND isInt = 0 then
	    
	    if func = "reddb" then
	    frd_sySQL = frd_sySQL & "frd_syg_"&x&" = "& replace(request("FM_frd_sy_"&x&""),",",".") &","
	    else
	    frd_sySQL = frd_sySQL & replace(request("FM_frd_sy_"&x&""),",",".") &","
	    end if
	    
	else
	    if func = "reddb" then
	    frd_sySQL = frd_sySQL & "frd_syg_"&x&" = 0,"
	    else
	    frd_sySQL = frd_sySQL &"0,"
	    end if
	    
	end if
	isInt = 0
	next
	
	call trimFld(fsFlds)
	call trimFld(frd_sySQL) 
	
	
	'**** Andet fradrag ***
	faFlds = ""
	frd_anSQL = ""
	for x = 1 to 12 
	faFlds = faFlds & "frd_andet_"&x&"," 
	call erDetInt(request("FM_frd_an_"&x&""))
	    if len(request("FM_frd_an_"&x&"")) <> 0 AND isInt = 0 then
	        
	        if func = "reddb" then
	        frd_anSQL = frd_anSQL & "frd_andet_"&x&" = "& replace(request("FM_frd_an_"&x&""),",",".") &","
	        else
	        frd_anSQL = frd_anSQL & replace(request("FM_frd_an_"&x&""),",",".") &","
	        end if
	        
	    else
	    
	        if func = "reddb" then
	        frd_anSQL = frd_anSQL & "frd_andet_"&x&" = 0,"
	        else
	        frd_anSQL = frd_anSQL &"0,"
	        end if
	    end if
	next
	
	call trimFld(faFlds)
	call trimFld(frd_anSQL) 
	
	if func = "dbopr" then
	strSQL = "INSERT INTO budget_bruttonetto_dt "_
	&" (navn, budget_aar, editor, "& bdFlds &", "& ffFlds &", "& fsFlds &", "& faFlds &") "_
	&" VALUES ('"& strNavn &"', "& budgetAar &", '"& editor &"', "_
	&" "& bruttoDageSQL &", "& frd_feSQL &", "& frd_sySQL &", "& frd_anSQL &")"
	
	'Response.Write strSQL
	'Response.flush
	
	oConn.execute(strSQL)
	
	strSQL = "SELECT id FROM budget_bruttonetto_dt WHERE id <> 0 ORDER BY id DESC"
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	lastid = oRec("id")
	end if
	oRec.close
	
	else
	strSQL = "UPDATE budget_bruttonetto_dt SET "_
	&" navn = '"& strNavn &"', budget_aar = "& budgetAar &", editor = '"& editor &"',"_
	& bruttoDageSQL &", "& frd_feSQL &", "& frd_sySQL &", "& frd_anSQL &""_
	&" WHERE id = "& id
    
    'Response.Write strSQL
	'Response.flush
	
	oConn.execute(strSQL)
   
    lastid = id
    
    end if
	
	
	
	
	
	Response.redirect "budget_bruttonetto.asp?func=red&id="&lastid&"&FM_aar="&budgetAar&"&fromupdate=1"
	
	case "opr", "red"
	
	
	            call top()
	            call vaelgAar()
            	
	            if func = "opr" then
	            dbfunc = "dbopr"
	            else
	            dbfunc = "reddb"
	            end if
	            %>
            	
	            <%
	            
            	
     
    dim bdFlt, ffFlt, fsFlt, faFlt, sumNettoDage
	redim bdFlt(12), ffFlt(12), fsFlt(12), faFlt(12), sumNettoDage(12)
	
    
    
     
    if func = "red" then 	
            	
    bdFlds = ""
	ffFlds = ""
	fsFlds = ""
	faFlds = ""
	for x = 1 to 12 
	bdFlds = bdFlds & "bd_"&x&"," 
	ffFlds = ffFlds & "frd_ferie_"&x&","
	fsFlds = fsFlds & "frd_syg_"&x&"," 
	faFlds = faFlds & "frd_andet_"&x&"," 
	next
	
	call trimFld(bdFlds)
	call trimFld(ffFlds)
	call trimFld(fsFlds)
	call trimFld(faFlds)
	
	
	
	strSQL = "SELECT navn, budget_aar, dato, "_
	&""& bdFlds &","& ffFlds &","& fsFlds &","& faFlds &""_
	&" FROM budget_bruttonetto_dt WHERE id = "& id
	
	'Response.Write strSQL
	'Response.flush
	
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	
	for x = 1 to 12
	bdFlt(x) = oRec("bd_"&x&"")
	ffFlt(x) = oRec("frd_ferie_"&x&"")
	fsFlt(x) = oRec("frd_syg_"&x&"")
	faFlt(x) = oRec("frd_andet_"&x&"")
	next
	
	end if 
	oRec.close
	
	end if
	
	%>
            	
	            <table cellspacing=1 cellpadding=2 border=0 bgcolor="#5582d2" width=90%>
	            <form method=post action=budget_bruttonetto.asp?func=<%=dbfunc%>>
                    <input id="id" name="id" value="<%=id %>" type="hidden" />
                    <input id="FM_aar" name="FM_aar" value="<%=thisaar%>" type="hidden" />
	        <tr bgcolor="#8cAAe6">
                <td width=100 class=alt height=20><b>Type/Måned</b></td>
                <td class=alt><b>Jan</b></td>
                <td class=alt><b>Feb</b></td>
                <td class=alt><b>Mar</b></td>
                <td class=alt><b>Apr</b></td>
                <td class=alt><b>Maj</b></td>
                <td class=alt><b>Jun</b></td>
                <td class=alt><b>Jul</b></td>
                <td class=alt><b>Aug</b></td>
                <td class=alt><b>Sep</b></td>
                <td class=alt><b>Okt</b></td>
                <td class=alt><b>Nov</b></td>
                <td class=alt><b>Dec</b></td>
            </tr>
            <tr bgcolor="#d6dff5">
            <td >Arb. dage i md.:</td>
           
           
            <%
            
            
            for x = 1 to 12 
            countWdays = 0
            
            
            if func = "red" then
            countWdays = bdFlt(x) 
            else
           
           
            select case x
            case 1,3,5,7,8,10,12
            stDato = "1-"&x&"-"& thisAar
            slDato = "31-"&x&"-"& thisAar
            case 4,6,9,11
            stDato = "1-"&x&"-"& thisAar
            slDato = "30-"&x&"-"& thisAar
            case 2
            stDato = "1-"&x&"-"& thisAar


                '** Skudår **'
                select case right(thisAar, 2)
                case "04", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60"
                slDato = "29-"&x&"-"& thisAar '*** Find skudår ***'
                case else
                slDato = "28-"&x&"-"& thisAar '*** Find skudår ***'
                end select
                
            end select
            
            
            antalNettoDage = datediff("d",stDato, slDato, 2,2)
            for d = 0 to antalNettoDage
            datodag = dateadd("d", d, stDato)

                    wDay = weekday(datodag, 2)
                    select case wDay
                    case 0,1,2,3,4  
                        call helligdage(datodag)
                        if erHellig <> 1 then
                        countWdays = countWdays + 1
                        else
                        countWdays = countWdays 
                        end if
                    case 5,6
                    countWdays = countWdays
                    end select 
                    countWdays = countWdays
            next
            
            end if 'red

            sumNettoDage(x) = countWdays
            %>
            <td align=right><input id="FM_bd_dage_<%=x %>" name="FM_bd_dage_<%=x %>" type="text" style="width:35px; font-size:9px;" value=<%=countWdays %> /></td>
            <%next %>
            </tr>

            <tr bgcolor="#8cAAe6"><td colspan=13 class=alt height=20><b>Fradrag</b></td></tr>

            <tr bgcolor="#d6dff5">
            <td>Ferie</td>
            <%for x = 1 to 12 %>
            <td align=right><input id="FM_frd_fe_<%=x %>" name="FM_frd_fe_<%=x %>" type="text" value="<%=ffFlt(x) %>" style="width:35px; font-size:9px;"/></td>
            <%
            sumNettoDage(x) = sumNettoDage(x) - (ffFlt(x))
            next %>
            </tr>
            <tr bgcolor="#ffffff">
            <td>Sygdom</td>
            <%for x = 1 to 12 %>
            <td align=right><input id="FM_frd_sy_<%=x %>" name="FM_frd_sy_<%=x %>" type="text" value="<%=fsFlt(x)%>" style="width:35px; font-size:9px;"/></td>
            <%
            sumNettoDage(x) = sumNettoDage(x) - (fsFlt(x))
            next %>
            </tr>
            <tr bgcolor="#d6dff5">
            <td>Uddannelse/Andet</td>
            <%for x = 1 to 12 %>
            <td align=right><input id="FM_frd_an_<%=x %>" name="FM_frd_an_<%=x %>" type="text" value="<%=faFlt(x)%>"  style="width:35px; font-size:9px;"/></td>
            <%
            sumNettoDage(x) = sumNettoDage(x) - (faFlt(x))
            next %>
            </tr>

            <tr bgcolor="pink">
            <td height=20><b>= Netto dage i md.:</b></td>
            <%for x = 1 to 12 
            %>
            <td align=right style="padding-right:5px;"><%=formatnumber(sumNettoDage(x),2) %></td>
            <%next %>
            </tr>

            <tr bgcolor="#ffffff"><td colspan=13><b>Timer</b></td></tr>

            <tr>
            <td>Brutto timer</td>
            <%for x = 1 to 12 %>
            <td><input id="Text1" type="text" size=5 value=170 /></td>
            <%next %>
            </tr>

            <tr>
            <td>Netto timer</td>
            <%for x = 1 to 12 %>
            <td><input id="Text1" type="text" size=5 value=116 /></td>
            <%next %>
            </tr>

            </table>
            <br />
            <b>Gem brutto / netto tal for det valgte år:</b><br />
             
            * Navn: <input id="FM_navn" name="FM_navn" type="text" value="<%=strNavn %>" size="40" />
            <input id="Submit1" type="submit" value="Gem / Opdater >>" />

            </form>	
            	
	 <%
	case else
		
	 call top()
	 call vaelgAar()
	


	
	end select
	%>
	</div>
	<%
end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
