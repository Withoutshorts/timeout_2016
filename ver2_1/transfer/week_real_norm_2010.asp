<%response.Buffer = true %>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/xml/tsa_xml_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="inc/helligdage_func.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->



<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function
	
	func = request("func")
	id = request("id")
	thisfile = "week_norm_2010.asp"
	
	call erStempelurOn()
	
	'*** Sætter lokal dato/kr format. *****
	Session.LCID = 1030
	
	'sendemail = request("sendemail")
	
	select case func 
	case "-"
	
	case else
	
	
	
	dim intMids
	redim intMids(150)
	
	
	if len(trim(request("FM_progrp"))) <> 0 then
	progrp = request("FM_progrp")
	else
	progrp = 0
	end if
	
	'Response.Write "medid first: "& left(request("FM_medarb"), 1)
	'Response.end
	
	'*** Rettigheder på den der er logget ind **'
	medarbid = session("mid")
	
	media = request("media")

	if len(request("FM_medarb")) <> 0 OR media = "j" OR func = "export" then
	
	    if left(request("FM_medarb"), 1) = 0 then
	    thisMiduse = request("FM_medarb_hidden")
    	intMids = split(thisMiduse, ", ")
	    else
	    thisMiduse = request("FM_medarb")
	    intMids = split(thisMiduse, ", ")
	    end if
	
	else
	    
	    thisMiduse = session("mid") 
	    intMids = split(thisMiduse, ", ")
	   
	end if
	
	
	'Response.Write request("FM_medarb")
	'Response.end
	
	'Response.Write "year(now) = trim(request(yuse)"& year(now) &" = "& trim(request("yuse"))
	
	if len(request("muse")) <> 0 then
    mnow = request("muse")
    else
    mnow = month(now)
    end if
	
	if len(trim(request("yuse"))) = 0 then
    ugp = now
    ddato = datepart("d", ugp) &"/"& datepart("m", ugp) &"/"& datepart("yyyy", ugp)
    ddato = dateadd("d",-1, ddato)
    showigartxt = "("&tsa_txt_319&")"
    else
        if cint(year(now)) = cint(request("yuse")) then
            if cint(mnow) < cint(month(now)) then
                call dageimd(mnow, request("yuse"))
                ds = mthDays
            else
            ds = datepart("d", now)
            end if
         else
            call dageimd(mnow,request("yuse"))
            ds = mthDays
         end if
            
    showigartxt = ""
    ugp = ""& ds &"-"& mnow &"-"&request("yuse")
    
    'Response.Write ugp
    'Response.flush
    ddato = datepart("d", ugp) &"/"& datepart("m", ugp) &"/"& datepart("yyyy", ugp)
    
    ''** Maks opgjort frem til igår ****'
    if cDate(now) > cDate(ddato) then
    ddato = dateadd("d",-1, ddato)
    showigartxt = "("&tsa_txt_319&")"
    end if
    
    end if
	
	
	
	if func <> "export" then
	
	
	if media <> "j" then
	
	leftPos = 20
	topPos = 132
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	<%call tsamainmenu(7)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	if showonejob <> 1 then
		call stattopmenu()
	end if
	%>
	</div>
	
	<%else 
	
	leftPos = 20
	topPos = 20
	
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%end if%>
	
	<%end if%>
	
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=leftPos%>; top:<%=topPos%>; visibility:visible;">
	
	
	<%
	
	oimg = "ikon_medarbaf_48l.png"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = "Medarbejder afstemning"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	
	
	
	if media <> "j" then
	
	call filterheader(0,0,800,pTxt)%>
	<table border=1 cellspacing=2 cellpadding=2 width=100%>
	<form action="week_real_norm_2010.asp" method="post" name="periode" id="periode">
    <input id="FM_usedatokri" name="FM_usedatokri" value="1" type="hidden" />
    <tr>
        <td valign=top><b>Projektgrupper:</b><br />
       
       <%
       
       
       
       call projgrp(progrp,level,medarbid)
       %>
        <select id="Select1" name="FM_progrp" style="width:200px;" onchange="submit()">
       
       <%
       for p = 0 to prgAntal 
       
        if cint(progrp) = prjGoptionsId(p) then
        pgSEL = "SELECTED"
        else
        pgSEl = ""
        end if%>
        <option value="<%=prjGoptionsId(p) %>" <%=pgSEl%>><%=prjGoptionsTxt(p) %> (<%=prjGoptionsAntal(p) %>)</option>
      <%next %>
            </select>
            
            <br /><br /><input id="Submit2" type="submit" value="Vis proj.grp >> " />
        </td>
  
	<td valign=top><b>Medarbejder:</b><br /> 
	
	<select name="FM_medarb" id="FM_medarb" multiple style="width:200px;" size=15>
	<%
	
	mSel = ""
	'if cint(trim(intMids(0))) = 0 then
    'mSel = "SELECTED"
    'end if%>
	
	<option value="0">Alle</option>
	<%
	call medarbiprojgrp(progrp, medarbid)
	
    for m = 0 to antalMedgrp
        
        mSel = ""
        for t = 0 To UBOUND(intMids) 
            
            if mSel = "" then
            if cint(intMids(t)) = cint(medarbgrpId(m)) then
            mSel = "SELECTED"
            else
            mSel = ""
            end if
            
            end if
            
            
        
        next
	%>
	<option value="<%=medarbgrpId(m)%>" <%=mSel%>><%=medarbgrpTxt(m)%></option>
		
	<%
	next
	
	%></select>
	
	<input id="FM_medarb_hidden" name="FM_medarb_hidden" type="hidden" value="0" />
	<%for m = 0 to antalMedgrp %>
	<input id="FM_medarb_hidden" name="FM_medarb_hidden" type="hidden" value="<%=medarbgrpId(m)%>" />
	<%next %>
	
	</td>
	<td valign=top><b>År:</b><br />
        <select id="Select2" name="yuse">
        <%
	ysel = now
	for y = 0 to 5 %>
    <%yShow = dateAdd("yyyy", -y, ysel) 
        
        if year(yShow) = year(ugp) then
        ySele = "SELECTED"
        else
        ySele = ""
        end if
    
    %>
	<option value="<%=datePart("yyyy", yShow, 2,2)%>" <%=ySele %>><%=datePart("yyyy", yShow, 2,2)%></option> 
	<%next %>
            
        </select>
        
        <br />
        <select id="Select3" name="muse">
        
            <%for m = 1 to 12 
                
                if m = month(ugp) then
                mSele = "SELECTED"
                else
                mSele = ""
                end if 
                
            %>       
            <option value="<%=m %>" <%=mSele %>><%=monthname(m) %></option>
            <%next %>
        </select>
	
	</td>
	</tr>
	
	<tr>    
	  
	    <td colspan=3 align=right valign=bottom style="padding-top:16px;">
            <input id="Submit1" type="submit" value=" Søg >> " />
	</td></tr>
	</form></table>
    
    
    
    <!-- filter header sLut -->
	</td></tr></table>
	</div>
    
    
    <%
    
    else
   
    end if 'media
    
    
    
	public anormTimerTot, arealTimerTot, atotalTimerPer100, aafspadUdbBalTot, aAfspadBalTot, arealfTimerTot
    public afradragTimerTot, altimerKorFradTot, aafspTimerTot, aafspTimerBrTot, aafspTimerUdbTot
    		

	
	tTop = 20
	tLeft = 0
	tWdth = 1004
	
	
	
	
	call tableDiv(tTop,tLeft,tWdth)
	%>
	<!--- Table start -->
	
	<table cellpadding=0 cellspacing=0 border=0>
     <tr><td colspan="5" valign=top style="padding:10px;">
	
	<font class=roed>Opgjort pr. <b><%=formatdatetime(ddato, 1) %></b> <%=showigartxt %></font><br />
	<% 
    Response.Write "<b>År -> Dato:</b> "& showigartxt &"<br><br>"
    
    
    %>
    </td></tr></table>
    
    <table cellpadding=0 cellspacing=5 border=0>
    
    <%
	for m = 0 to UBOUND(intMids)
	   
	   if cint(intMids(m)) <> 0 then
	    
	   
	    
	    select case right(m,2)
	    case 0, 3, 6, 9, 12, 15, 18, 21, 24, 27, 30, 33, 36, 39, 42, 45, 48, 51, 54, 57, 60, 63, 66, 69, 72, 75, 78, 81, 84, 87, 90, 93, 96, 99, 102, 105, 108, 111, 114, 117, 120, 123, 126, 129, 132, 135, 138, 141, 144, 147, 150
	     if m <> 0 then%>
	    </td></tr>
	     <%end if %>
	    <tr><td valign=top style="padding:10px; border:1px #cccccc solid;">
	    <%
	    case else
	    %>
	    </td>
	    <td valign=top style="padding:10px; border:1px #cccccc solid;">
	    <%
	    end select
	    
	   
	
	meTxt = ""
	call meStamdata(intMids(m))
	
	
	%>
	
	<br /><br />
	<b><%=meTxt%></b>
    <table cellspacing=0 cellpadding=2 border=0 width="100%"><tr>
    <td style="border-bottom:1px silver dashed;" class=lille><b>Uge</b></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_172%></b></td>
	  <%if lto <> "cst" then %>
	 <td class=lille bgcolor="#cccccc" align=right style="border-bottom:1px silver dashed;">(heraf<br />fak.bare)</td>
	 <%end if %>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_173%></b></td>
	  
	  <%if session("stempelur") <> 0 then %> 
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b>Løn timer</b></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b>Till. / <br />Frad. +/-</b></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b>Sum</b></td>
	 <%end if %>
	 
	  <%if lto <> "cst" then %>
	  <td bgcolor="pink" align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_284%></b><br /> (Real / Norm)</td>
	  <%end if %>
	  
	  <%if session("stempelur") <> 0 then %> 
	 <td bgcolor="#9acd32" align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_284%></b> <br />(Lønt. / Norm)</td>
	 
	  <%if lto <> "cst" then %>
	  <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_166%></b> <br />(Real / Lønt.)</td>
	  <%end if %>  
	    
	 <%end if %>
	 
	   <!-- Afspad / Overarb --->
	       <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 then %>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_283%><br /> <%=tsa_txt_164%></b><br />(enh.)</td>
	          <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><b>Afspads.</b></td>
	           <td align=right style="border-bottom:1px silver dashed;" class=lille><b>Udbetalt</b></td>
	           <td align=right style="border-bottom:1px silver dashed;" class=lille><b>Ønsk. Udbe.</b></td>
	           <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=tsa_txt_283 &" "& tsa_txt_280 %></b></td>
	       
	 
	 <%end if %>
    </tr>
    <%
    
    weekThis = datePart("ww", ddato, 2,2)
    weekStMth = datePart("ww", "1-"& month(ddato)&"-"&year(ddato), 2,2)
    
    weekDiff = dateDiff("ww", "1-"& month(ddato)&"-"&year(ddato), ddato, 2, 2) 
    
   
    
    for wth = 0 to weekDiff
    
    'if wth <> 1 then
    datoB = dateadd("ww", -wth, ddato)
    datoB = day(datoB)&"/"& month(datoB) &"/"& year(datoB)
    'datoB = dateadd("d", -1, datoB)
    
    weekDayThisB = weekday(datoB, 2)
    select case weekDayThisB
    case 1
    datoB = dateadd("d", 6, datoB)
    case 2
    datoB = dateadd("d", 5, datoB)
    case 3
    datoB = dateadd("d", 4, datoB)
    case 4
    datoB = dateadd("d", 3, datoB)
    case 5
    datoB = dateadd("d", 2, datoB)
    case 6
    datoB = dateadd("d", 1, datoB)
    case 7
    datoB = dateadd("d", 0, datoB)
    end select
    
    
    'Response.Write "weekDayThis" & weekDayThisA & "<br>"
    
    
    datoA = dateadd("d", -(6), datoB)
    
    
    
    '** Maks frem til DD ***'
    
    if cDate(ddato) < cDate(datoB) then
    datoB = ddato
    end if
    
    slutDatoLastm_B = datepart("yyyy", datoB) &"/"& datepart("m", datoB) &"/"& datepart("d", datoB)
    slutDatoLastm_A = datepart("yyyy", datoA) &"/"& datepart("m", datoA) &"/"& datepart("d", datoA)
    
    
    'Response.Write slutDatoLastm_A & " - " & slutDatoLastm_B & "<br>"
    
    call medarbafstem(trim(cint(intMids(m))), slutDatoLastm_A, slutDatoLastm_B, 14, akttype_sel)
    'Response.Write "<br>"
	response.flush
	
	next
	%>
	<tr>
	<td style="border-bottom:1px silver dashed;" class=lille><b>Total:</b></td>
	<td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><b><%=formatnumber(arealTimerTot,2)%></b></td>
	 <%if lto <> "cst" then %>
	<td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille>(<%=formatnumber(arealfTimerTot,2)%>)</td>
	<%end if %>
	
	
    <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><b><%=formatnumber(anormTimerTot,2)%></b></td>
	
	<%if session("stempelur") <> 0 then %> 
	 <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille>
	 <%call timerogminutberegning(atotalTimerPer100) %>
    <b><%=thoursTot &":"& left(tminTot, 2) %></b>
	</td>
	 <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><b><%=formatnumber(afradragTimerTot, 2) %></b></td>
	 <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><b><%=formatnumber(altimerKorFradTot, 2) %></b></td>
	 <%
	 end if %>
	 
	 <%if lto <> "cst" then %>
	 <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><b><%=formatnumber((arealTimerTot - anormTimerTot),2)%></b></td>
	  <%
	 end if %>
	 
	   <%if session("stempelur") <> 0 then %> 
	 <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille>
	 <%anormtime_lontimeTot = -((anormTimerTot - (altimerKorFradTot)) * 60) %>
	 <%call timerogminutberegning(anormtime_lontimeTot) %>
		<b><%=thoursTot &":"& left(tminTot, 2) %></b>
	 </td>
	    <%if lto <> "cst" then %>
	   <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><b><%=formatnumber((arealTimerTot - (altimerKorFradTot)),2)%></b></td>
	    <%end if %>
	 <%end if %>
	    
	    
	    
	      <!-- Afspad / Overarb --->
	       <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 then %>
	           
	         <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><b><%=formatnumber(aafspTimerTot, 2)%></b></td>
	         <td align=right class=lille style="border-bottom:1px silver dashed; white-space:nowrap;"><b><%=formatnumber(aafspTimerBrTot, 2)%></b></td>
	         <td align=right class=lille style="border-bottom:1px silver dashed; white-space:nowrap;"><b><%=formatnumber(aafspTimerUdbTot, 2)%></b></td>
	            <td align=right class=lille style="border-bottom:1px silver dashed; white-space:nowrap;"><b><%=formatnumber(aafspadUdbBalTot, 2)%></b></td>
	         
	          <td align=right class=lille style="border-bottom:1px silver dashed; white-space:nowrap;"><b><%=formatnumber(aAfspadBalTot, 2)%></b></td>
        	 
	         
	         
	         
	         <%end if %>
	
	
	
	
	</tr></table>
	
	<%
	
	altimerKorFradTot = 0
	afradragTimerTot = 0
	arealfTimerTot = 0
	anormTimerTot = 0
	atotalTimerPer100 = 0
	anormtime_lontimeTot = 0
	altimerKorFradTot = 0
	arealTimerTot = 0
	aafspTimerTot = 0
	aafspTimerBrTot = 0
	aafspTimerUdbTot = 0
	aafspadUdbBalTot = 0
	aAfspadBalTot = 0
	
	end if 'm <> 0 
	next 'medarb loop%>
	
	
	</td></tr></table>
	
	
	
	
	
	<%
	if media <> "j" AND func <> "export" then

'Response.flush

ptop = 70
pleft = 1014
pwdt = 120

call eksportogprint(ptop,pleft,pwdt)
%>


<tr>
    
  
    <td align=center><a href="bal_real_norm_2007.asp?func=export&FM_medarb=<%=thisMiduse%>" target="_blank"><img src="../ill/export1.png" border="0"></a></td>
    </td><td><a href="bal_real_norm_2007.asp?func=export&FM_medarb=<%=thisMiduse%>" class=rmenu target="_blank">.csv fil eksport</a></td>
 
    </tr>
    <tr>
    <td align=center><a href="bal_real_norm_2007.asp?media=j&FM_medarb=<%=thisMiduse%>" target="_blank"  class='rmenu'>
   &nbsp;<img src="../ill/printer3.png" border=0 alt="" /></a>
</td><td><a href="bal_real_norm_2007.asp?media=j&FM_medarb=<%=thisMiduse%>" target="_blank"  class='rmenu'>Printversion</a></td>
   </tr>
   
   
  
   
   
	
   </table>
</div>
<%else%>

<% 
Response.Write("<script language=""JavaScript"">window.print();</script>")
%>
<%end if%>
	
	
	
	<%
	'** Timer Realiseret ***'
	
	'for m = 0 TO UBOUND(intMids)
	'call medarbafstem(intMids(m), startdato, slutdato, 1, akttype_sel)
	'next
	
	%>
	</table>
    <!-- table div --> 
	</div>
	
	<br /><br />
	
	<%if func = "export" then 

   
	ekspTxt = replace(strEksportTxt, "xx99123sy#z", vbcrlf)
	
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\bal_real_norm_2007.asp" then
					Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\medarbafexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\medarbafexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				else
					Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\data\medarbafexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\data\medarbafexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				end if
				
				
				
				file = "medarbafexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				'**** Eksport fil, kolonne overskrifter ***
				
				
				
				
				
				
				'objF.writeLine("Periode afgrænsning: "& datointerval & vbcrlf)
				'objF.WriteLine(strOskrifter & chr(013))
				objF.WriteLine(ekspTxt)
				
				
				
				
	
	Response.redirect "../inc/log/data/"& file &""	

end if %>

      
      
      <%if sendemail = "j" then %>
	  <% 
	    Response.Write("<script language=""JavaScript"">window.alert(""Der er blevet afsendt en email til de valgte medarbejdere"");</script>")
        Response.Write("<script language=""JavaScript"">window.close();</script>")
        %>
	  <%end if %>
	
	
	  <%if media <> "j" ANd func <> "export" then  
        itop = 56
        ileft = 825
        iwdt = 120
        ihgt = 0
        ibtop = 10 
        ibleft = 150
        ibwdt = 600
        ibhgt = 1150
        iId = "pagehelp"
        ibId = "pagehelp_bread"
        call sideinfoId(itop,ileft,iwdt,ihgt,iId,phDsp,phVzb,ibtop,ibleft,ibwdt,ibhgt,ibId)%>		
       
   
			 <br /><br />
    
			<%
			call akttyper2009(3)
			%>
			</table>
			
		
	    
	    <br /><br />
	    <b>Ferie</b><br />
	    Fra den 1. januar 2001 optjenes 2,08 dages ferie for hver måneds ansættelse i optjeningsåret, som er lig kalenderåret. 
        Dette gælder også medarbejdere på del-tid.<br />
        2,08 * 12 måneder = 25 dage ell. 5 arbejdsuger.
	    <br /><br />
	    <b>Dage</b><br />
	    Ferie of ferie fridage (afspad.) timer angives som timer og omregnes til dage udfra normeret timer pr. uge 
	    (total timer pr. uge / med 5 arb. dage, se medarbejdertyper ell. medarb. afstemning).<br /><br />
	    Hvis man er ansat 37 timer, 
	    angiver man altså 7,4 timer for en hel dags ferie, og 37 timer for en uge. 
	    
	    <br />
	    
                &nbsp;
	   
	    <!-- side info slut -->
        </td></tr></table>
			</div>
			<br /><br />&nbsp;
			
		<%end if %>	
	    
	
	<br /><br /><br /><br /><br /><br /><br /><br />
        &nbsp;
			
			</div>
	
	<% 
	end select
	
	
	
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
