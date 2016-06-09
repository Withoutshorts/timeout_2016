<%
'********************************'
'*** faktura layout opbygning ***'
'********************************'

'***** Fak header ******
public strTypeThis 
public strTypeThis2
public pdfHTML
public filnavn, pdtop4
function fakheader() 

                select case lto
		        case "optionone"
		        pdtop1 = "125px 10px 0px 10px"
		        pdtop2 = "0px 0px 0px 0px"
		        pdtop3 = "5px 0px 0px 0px"
		        pdtop4 = "30px 2px 10px 10px"
		        case "dencker", "tooltest" ', "optionone"
		        pdtop1 = "45px 10px 0px 10px"
		        pdtop2 = "0px 0px 0px 0px"
		        pdtop3 = "5px 0px 0px 0px"
		        pdtop4 = "10px 2px 10px 10px"
		        case "essens"
		        pdtop1 = "10px 10px 0px 10px"
		        pdtop2 = "0px 0px 20px 0px"
		        pdtop3 = "20px 0px 0px 0px"
		        pdtop4 = "10px 2px 10px 10px"
		        case "acc", "intranet - local"
		        pdtop1 = "15px 10px 0px 10px"
		        pdtop2 = "0px 0px 0px 0px"
		        pdtop3 = "55px 0px 0px 0px"
		        pdtop4 = "10px 2px 10px 10px"
                case "epi"
                pdtop1 = "0px 10px 0px 10px"
		        pdtop2 = "0px 0px 20px 0px"
		        pdtop3 = "-10px 0px 0px 0px"
		        pdtop4 = "10px 2px 10px 10px"
		        case else
		        pdtop1 = "10px 10px 0px 10px"
		        pdtop2 = "0px 0px 20px 0px"
		        pdtop3 = "28px 0px 0px 0px"
		        pdtop4 = "10px 2px 10px 10px"
		        end select

%>
<table cellspacing="0" cellpadding="0" border="0" width="620">
		<tr>
		<td valign="top" style="padding:<%=pdtop1%>; width:340px;">
	    
	    <table width=100% cellspacing=0 cellpadding=0 border=0>
		<tr><td valign="top" style="padding:<%=pdtop2 %>;">
		<%select case intFaktype
		case 0
		        
		        select case lto
		        case "rosenloeve"
		        
		        strtopimg = "faktura_top"
		        strTypeThis = "jobopgørelse"
		        strTypeThis2 = "Jobopgørelse"
		        
		        case else
		
		       		        
                strtopimg = "faktura_top"
                strTypeThis = lcase(txt_001) '"faktura"
                strTypeThis2 = txt_001 '"Faktura"
		        
		        
		        end select
		        
		        
		case 1
		strtopimg = "kreditnota_top"
		strTypeThis = lcase(txt_002) '"kreditnota"
		strTypeThis2 = txt_002 '"Kreditnota"   
		'case 2
		'strtopimg = "rykker_top"
		'strTypeThis = "rykker"
		'strTypeThis2 = "Rykker"  
		case else
		strtopimg = "blank"
		end select%>
		
		
		<h4>
        
        <%if (cint(medregnikkeioms) = 1 OR cdate(sysFakdato) = "01-01-2002" OR cint(intGodkendt) = 0) AND lto <> "essens" then%>
        <span style="text-decoration:line-through;"><%=strTypeThis2%> - <%=varFaknr%></span> 
        
        <%if cint(medregnikkeioms) = 1 then %>
        &nbsp;(intern)
        <%end if %>

        <%if cint(intGodkendt) = 0 then %>
        &nbsp;(kladde)
        <%end if %>

        <%if cdate(sysFakdato) = "01-01-2002" then %>
        &nbsp;(nulstillet)
        <%end if %>
        
        
        <%else %>
        <%=strTypeThis2%> - <%=varFaknr%>
        <%end if%>
        
        </h4>
		<!--<img src="../ill/<=strtopimg%>.gif" width="271" height="25" alt="" border="0">-->
		</td>
		</tr>
		<tr>
		<td valign="top" style="padding:<%=pdtop3%>;">
		<%call modtager_layout() %>
		</td>
		</tr>
		</table>
		
		</td>
		<td style="width:280px;" valign=top align=right>
		
		
		
		<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr><td align="right" colspan="4" valign="top" style="height:50px; padding:8px 3px 0px 0px;">
		<%
		
		filnavn = ""
		
		strSQLlogo = "SELECT f.filnavn, k.kid, k.logo FROM kunder k "_
		&" LEFT JOIN filer f ON (f.id = k.logo) WHERE k.useasfak = 1"
		
		'Response.Write strSQLlogo
		'Response.flush
		
		oRec.open strSQLlogo, oConn, 3
		if not oRec.EOF then
		
		filnavn = oRec("filnavn")
		
		%>
		    <!--
            <img src="../inc/upload/<=lto%>/<=oRec("filnavn")%>" /> 
            <img src="../inc/upload/intranet/logo.jpg" />
            -->
		
		<%
		end if
		oRec.close
		
		if filnavn <> "" then %>
		<img src="../inc/upload/<%=lto%>/<%=filnavn%>" /> 
		<%else %>
		<img src="../ill/blank.gif" /> 
		<%end if %>
		
		
		</td>
		</tr>
		<tr>
		<%
		pdtop = "30px 0px 0px 0px;"
		
		select case lto
		case "acc"
		wdt = 140
		case "outz"
		wdt = 120
        case "epi"
		wdt = 80
		case "syncronic"
		wdt= 120
		case "dencker", "tooltest", "optionone", "intranet - local"
		pdtop = "45px 0px 0px 0px;"
		wdt = 80
		case "essens"
		wdt = 100
		pdtop = "35px 0px 0px 0px;"
		case else
		wdt = "100"
		end select
		%>
		<td style="width:<%=wdt %>"><img src="../ill/blank.gif" width="<%=wdt %>" height="1" border="0" /> 
          </td>
		<td align=right valign=top style="height:140px; width:<%=280-wdt %>; padding:<%=pdtop %>">
		<%
		 select case lto
		 case "rosenloeve", "essens"
		 case else
		 call afsender_layout()
		 end select%>
		</td>
		</tr>
		</table>
		
		</td>
		</tr>
		
</table>
<%
end function

'***** Modtager boks ********
function modtager_layout()
%>
<table cellspacing="0" cellpadding="0" border="0">

<tr bgcolor="#ffffff">
	
	<td valign="top" style="padding:0px 0px 0px 0px height:162px;; border:0px silver solid;">
	
	<%
	'*** M65 afsender linie ***'
	select case lto
	case "acc", "syncronic", "pcm", "essens", "execon", "immenso", "epi", "jttek"
	
	case "dencker", "tooltest", "optionone", "intranet - local"
	%>
	<br /><br />
	<font size=1 color="#999999"><%=txt_021 %>:&nbsp;<%=left(yourNavn, 25)%>&nbsp;<%=yourAdr%>&nbsp;<%=yourPostnr%>&nbsp;<%=yourCity%><br></font>
	<%case else%>
	<font size=1 color="#999999"><%=txt_021 %>:&nbsp;<%=left(yourNavn, 25)%>&nbsp;<%=yourAdr%>&nbsp;<%=yourPostnr%>&nbsp;<%=yourCity%><br></font>
	<%end select %>
	
	
	<%if year(fakdato) < "2010" then %>
	
	<%=strKnavn%><br>
	<%=strKadr%><br>
	<%=strKpostnr%>&nbsp;<%=strBy%><br>
	
	<%if intVismodland = 1 then
	Response.write strLand &"<br>"
	end if
	
	if len(trim(ean)) <> 0 then 
	Response.Write "EAN: "& ean & "<br>"
	end if
	
	
	else
	
	Response.Write modtageradr 
	
	end if
	
	
	if intVismodatt = 1 then
	Response.write "<br><b>"& txt_022 &": "& showAtt &"</b>"
	end if 
	%>
	
	<%if fak_ski = 1 then
	Response.Write "<br><i>"& txt_055 &"</i>"
	end if %>
	
	<%
	if lto = "execon" oR lto = "immenso" OR lto = "intranet - local" then
	
	if fak_abo = 1 then
	Response.Write "<br><i>Dette er en Lightpakke</i>"
	end if %>
	
	<%if fak_ubv = 1 then
	Response.Write "<br><i>Er omfattet af Udbudsvagten</i>"
	end if %>
	
	<%
	end if
	
	'if intVismodtlf = 1 then
	'Response.write "<br>"& txt_003 &": "& intTlf
	'end if%>
	
	<%'if intVismodcvr = 1 then
	'Response.write "<br>"& txt_011 &": " & intCvr 
	'end if%>
	
	<br><br><br />
	<%=txt_005 %>:&nbsp;<%=intKnr%>
	<br>
	<%=strTypeThis2 %><%=txt_006 %>: <%=replace(formatdatetime(fakdato,2),"-",".")%>
	
	<%if cint(visperiode) = 1 then  %>
	<br />
	<%=txt_049 %>: <%=replace(formatdatetime(istDato,2),"-",".")%> - <%=replace(formatdatetime(istDato2,2),"-",".")%>
	<%end if %>
	
	<%if cdate(forfaldsdato) >= cdate(fakdato) then  %>
	<br />
	<%=txt_007 %>: <%=replace(formatdatetime(forfaldsdato,2),"-",".")%><br />
	<%end if %>
	<!--<br><%=strTypeThis2%> nr:&nbsp;<b><%=varFaknr%></b>-->
	</td>
	</tr>

</table>		
<%		
end function


'***** maintable **********
sub maintable
%>
<table cellspacing="0" cellpadding="0" border="0" width="620">
	
	<tr>
		<td style="padding-left:0px;" valign="top" colspan="4">
<%
end sub

'**** maintable 2 *********
sub maintable_2
%>
</td>
<td width="20"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
<td style="padding-right:0px;" valign="top" colspan="2">
<%
end sub


'**** maintable 3 *********
sub maintable_3
%>
</td>
</tr>
</table>
<%
end sub


'****************************'
'****** Afsender boks *******'
'****************************'
function afsender_layout()
%>
<table cellspacing="0" cellpadding="0" border="0" width=100%>

<tr>
	<td valign="top" bgcolor="#ffffff" style="padding:<%=pdtop4%>;">
	
	<%select case lto
	case "dencker", "tooltest", "intranet - local", "optionone"
	%>
	<b>Bank og konto nr.:</b><br />
	<%
	case else
	%>
	
	<%=yourNavn%><br>
	<%=yourAdr%><br>
	<%=yourPostnr%>&nbsp;<%=yourCity%><br>
	<%=yourLand & "<br>"%>
	
	
	<%if intVisafstlf = 1 then
	Response.write "<br>"& txt_003 &": "& yourTlf
	end if%>
	
	<%if intVisafsfax = 1 then
	Response.write "<br>"& txt_023 &": "& yourFax
	end if%>
	
	
	<%if intVisafsemail = 1 then
	Response.write "<br>"& txt_024 &": " & yourEmail
	end if%>
	
	
	<%end select%>

    <%if intVisafscvr = 1 then%>
	<br /><%=txt_011 %>: <%=yourCVR%>
	<%end if%>
	
	<%if vorref <> "-" then%>
	<br /><%=txt_056 %>: <%=vorref%><br />
	<%end if%>
	
	<br />
	<%=yourBank%><br>
	<%=txt_008 %>:<br />
	<%=yourRegnr%> - <%=yourKontonr%><br>
	
	<%if intVisafsswift = 1 then%>
	<%=txt_010 %>:&nbsp;<%=yourSwift%><br>
	<%end if%>
	
	<%if intVisafsiban = 1 then%>
	<%=txt_009 %>:&nbsp;<%=yourIban%><br>
	<%end if%>
   
	
	
	
	<%
	'*** FI nummer **'
	select case lto
	case "dencker", "tooltest" ', "intranet - local"
	%>
	FI: <%=finummer %> 
	<%
	case else
	
	end select
	%>

     </font>
	
	</td>
</tr>
</table>
<%
end function




'**********************'
'****** Vedr. *********'
'**********************'

sub vedr
%>
<table cellspacing="0" cellpadding="0" border="0" width="620">
	<%select case lto
	case "syncronic"
	%>
	<tr>
	<td style="padding:5px 15px 15px 5px; border-top:1px #999999 solid; border-bottom:0px #999999 solid;">
    <%
	    
	    if jobid <> 0 then%>
	    <%=txt_012%>: <h4><%=strJobnavn%></h4>
		    <%if len(trim(strJobBesk)) <> 0 then%>
		    <br><%=strJobBesk%>
		    <%end if%>
		
		<br /><%=txt_050 %>: <%=intJobnr%><br />
		
	    <%
	    else
	    %>
	    <%=txt_012%>: <h4><%=strAftNavn%>&nbsp;(<%=intAftnr%>)</h4>
            <%if len(trim(strAftVarenr)) <> 0 AND strAftVarenr <> "0" then %>
            <br>
	        <%=txt_026 %>: <%=strAftVarenr%>
            <%end if %>
	    <%
	    end if
	    
	    %>
	    </td>
	    </tr>
	    <%
	
	case else
	    
	    %>
	    <tr>
	    <td style="padding:25px 5px 2px 5px; border-bottom:1px #999999 solid;">
	    <b><%=txt_014 %></b>
	    </td>
	    </tr>
	    <tr>
	    <td style="padding:5px 15px 15px 5px;">
	    
        <%
	    
	    if jobid <> 0 then%>
	        <%if cint(visikkejobnavn) <> 1 then %>
	        <b><%=strJobnavn%>&nbsp;(<%=intJobnr%>)</b><br />
	        <%end if %>
	    
	        <%if rekvnr <> "0" then %>
	        <%=txt_053 %>: <b><%=rekvnr%></b><br /><br />
	        <%end if %>
	    
		    <%if len(trim(strJobBesk)) <> 0 then%>
		    <%=strJobBesk%><br><br>
		    <%end if%>
	    <%
	    else
	    %>
	    <b><%=strAftNavn%>&nbsp;(<%=intAftnr%>)</b>

	            <%if lto = "outz" OR lto = "intranet" then %>
	            <br /><%=txt_026 %> / Licensnummer:
	            <%else %>

                        <%if len(trim(strAftVarenr)) <> 0 AND strAftVarenr <> "0" then %>
                        <br>
	                    <%=txt_026 %>: <%=strAftVarenr%>
                        <%end if 
	           
	            end if
	    end if  %>
        &nbsp;
	    </td>
	    </tr>
	    <%
	
    end select %>
	
	
	
	

</table>
<br /><br />
<%
end sub



'****** Udspecificering *******
function udspecificering(strFakdet,aktmat,antalAktlinier)

%>

<table cellspacing="0" cellpadding="1" border="0" width="620">

<%if aktmat = 3 then %>
<tr>
<td colspan="9" style="padding:10px 0px 2px 0px;"><b><%=txt_041 %></b></td>
</tr>
<%end if 


'*** Fjern overskrifs linie fra materialer **'
if (lto = "essens") AND (aktmat = 2 OR aktmat = 3) AND antalAktlinier <> 0 AND cint(hideantenh) <> 1 then
%>
                <tr>
                <td align=right style="width:40px;">&nbsp;</td>
                <td style="width:40px;">&nbsp;</td>
	            
	            <td colspan="2" style="width:200px;">
                    &nbsp;</td>
	            <td align=right style="width:80px;">&nbsp;</td>
	            <td align=right style="width:90px;">&nbsp;</td>
	            <td style="width:50px;">&nbsp;</td>
	           
	            <td align=right style="width:120px;">&nbsp;</td>
            	
                </tr>
<%
else
%>


            <tr>
            	
                <%
                '** Skjul antal og enh kolonner **' (aktivitetslinier)
                if (cint(hideantenh) <> 1 AND aktmat = 1) OR aktmat <> 1 then
                %>
                <td align="right" style="width:40px; padding-right:5px; border-bottom:1px #999999 solid;"><b><%=txt_013 %></b></td>
                <%end if %>

            	
	                <%if jobid <> 0 then
            	    
                    '** Skjul antal og enh kolonner **'
                    if (cint(hideantenh) <> 1 AND aktmat = 1) OR aktmat <> 1 then

	                if (enhedsangFak = -1 AND aktmat = 1) then %>
	                <td style="padding-right:5px; width:40px; border-bottom:1px #999999 solid;">
                    &nbsp;</td>
                    <%else %>
	                <td style="padding-right:5px; width:40px; border-bottom:1px #999999 solid;"><b><%=txt_044 %></b></td>
	                <%end if 
                    
                    end if%>
            	
	            <%else %>
                
                        <%
                        '** Skjul antal og enh kolonner **' (aktivitetslinier)
                        if cint(hideantenh) <> 1 then
                        %>
	                    <td style="padding-right:5px; width:40px; border-bottom:1px #999999 solid;">&nbsp;</td>
	                    <%end if %>
                
                <%end if %>
            	
	            <td colspan="2" style="width:200px; padding-left:5px; border-bottom:1px #999999 solid;"><b><%=txt_057 %></b></td>
	            <td align="right" style="width:80px; padding-right:2px; border-bottom:1px #999999 solid;">&nbsp;</td>

                <% 
                '** Skjul antal og enh kolonner **'
                if (cint(hideantenh) <> 1 AND aktmat = 1) OR aktmat <> 1 then%>
                <td align=right style="width:90px; padding-right:2px; border-bottom:1px #999999 solid;"><b><%=txt_015 %></b></td>
                 <%else %>
                <td style="border-bottom:1px #999999 solid;">&nbsp;</td>
                <%end if %>

	            <%if cint(visrabatkol) = 1 then %>
	            <td align=right style="width:50px; padding-right:5px; border-bottom:1px #999999 solid;"><b><%=txt_016 %></b></td>
	            <%else %>
	            <td style="width:50px; border-bottom:1px #999999 solid;">
                    &nbsp;</td>
	            <%end if %>
	           
	            <td align=right style="width:120px; border-bottom:1px #999999 solid;"><b><%=txt_017 %></b></td>
            	
            </tr>

<%end if %>            
            
        
<%=strFakdet%>
</table>
<br /><br />
<!--<tr>-->

<%
end function


'***** Sub-Total aktiviteter***

function subtotakt()
'** Bruges kun til at lave adftans før materiler eller total beløb på faktura. **'
'** Total antal er fjernet efter at KM er kommet med på faktura                **'
%>

<tr>
	<td valign='top' align="right" style="padding:5px 5px 10px 0px;">
	<%select case lto 
	case "-"
	%>
	<b><%=formatnumber(intFaktureretTimer, 2)%></b>
	<%
	case else  %>
	&nbsp;
	<%end select %>
	</td>
	<td colspan=7>
        &nbsp;</td>
</tr>

<!--
<tr>
    <td valign='top' colspan=3 style="padding:0px 0px 0px 0px;"><=txt_027 %></td>
	<td valign='top' colspan=5 align=right style="padding:0px 0px 0px 0px;"><=formatnumber(aktsubtotal, 2) &" "& valutaISO %></td>
</tr>
-->


<%
end function

'***** Sub-Total materialer***'

function subtotmat(m)
%>
<!--
<tr>
    <td valign='top' colspan=3 style="padding:5px 0px 0px 0px;"><%=txt_027 %></td>
    <td valign='top' colspan=5 align=right style="padding:5px 0px 0px 0px;"><%=formatnumber(matsubtotal(m), 2) &" "& valutaISO %></td>
</tr>
-->

<%
end function



'***** Total før moms og Moms ***'
function totogmoms()

'if len(strFakdet) = 0 then
%>
<table cellspacing="0" cellpadding="1" border="0" width="620">
<%
'end if

if cint(showmomsfriTxt) <> 0 then
trpadtop = 5%>
<tr><td valign="top" colspan="8" style="padding:25px 0px 0px 0px;">*) 0&#37; <%=txt_034 %></td></tr>
<%
else
trpadtop = 25
end if %>
<tr>
    <td valign="top" colspan="4" style="padding:<%=trpadtop%>px 0px 0px 0px;"><%=txt_029 %>&nbsp;</td>
   <td colspan="4" valign="top" align="right" style="padding:<%=trpadtop%>px 0px 0px 0px;"><%=formatnumber(intFaktureretBelob) &" "& valutaISO%></td>
</tr>


<tr>
    <td valign="top" colspan="4" style="padding:5px 0px 5px 0px;"><%=momssats &"&#37; "& txt_034 %></td>
    <td colspan="4" valign="top" align="right" style="padding:5px 0px 5px 0px;"><%=formatnumber(intMoms,2) &" "& valutaISO%></td>
	
</tr>
</table>
<%end function 


'***** Total beløb ***'
function totalbelob()%>

<table cellspacing="0" cellpadding="1" border="0" width="620">

<%
select case lto
case "essens"
case else

t = 1000
if len(trim(strFakmat(1))) <> 0 AND t = 1 then %>
<tr>
     <td valign="top" colspan="4" style="padding:25px 0px 0px 0px;"><%=txt_031 %>&nbsp;</td>
	<td valign="top" colspan="4" align="right" style="padding:25px 0px 0px 0px;"><%=formatnumber(intFaktureretBelob, 2) &" "& valutaISO%></td>
</tr>
<%end if 

end select
%>
 <%
	totalinclmoms = (intMoms/1) + (intFaktureretBelob/1)
	%>

<tr>
   <td valign="top" colspan="4" style="padding:5px 5px 5px 0px; border-top:1px #999999 solid; border-bottom:1px #999999 solid;"><b><%=txt_032 %></b>&nbsp;</td>
   <td valign="top" colspan="4" align="right" style="padding:5px 0px 5px 5px; border-top:1px #999999 solid; border-bottom:1px #999999 solid;"><b><%=formatnumber(totalinclmoms, 2) &" "& valutaISO%></b></td>
</tr>
</table>
<%
end function


'*** Komm. og bet. betingelser ***
sub komogbetbetingelser	

%>

<table cellspacing="0" cellpadding="1" border="0" width="620">
<tr>
<td colspan="8" style="border-top:0px silver solid; padding:35px 5px 5px 1px;">
<%select case lto
case "essens" ', "intranet - local"

if h > 16 then
toppx = 1848
else
toppx = 848
end if

%>
<%=strKomm%>
<div>&nbsp;</div>
<div>Ved senere betaling beregnes rente p&#229; 1,5% pr. p&#229;begyndt m&#229;ned. Rykkergebyr kr. 50,-. <br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
&nbsp; </div>
<div>&nbsp;</div>
<div style="width: 620px; position: absolute; top: <%=toppx%>px">
<table cellspacing="0" cellpadding="0" width="100%" border="0">
    <tbody>
        <tr>
            <td class="footer" style="width: 170px" valign="bottom">Bank: Br&#248;rup Sparekasse</td>
            <td class="footer" style="width: 80px" valign="bottom">Reg.: 8146</td>
            <td class="footer" style="width: 160px" valign="bottom">Konto: 0000148105</td>
            <td class="footer" style="width: 30px" valign="bottom"><img height="1" alt="" src="ill/blank.gif" width="1" border="0" /></td>
            <td class="footer" style="width: 170px" valign="top">Essens Kommunikation ApS<br />
            Vojensvej 11<br />
            DK-2610 R&#248;dovre<br />
            <br />
            T.: (+45) 70 22 91 49<br />
            mailbox@essens.info<br />
            www.essens.info<br />
            <br />
            CVR-nr.: 27022995 </td>
        </tr>
    </tbody>
</table>
</div>


<%
case "execon", "imenso"

if h > 16 then
toppx = 1848
else
toppx = 828
end if

%>
<%=strKomm%>

<div style="width: 620px; position: absolute; top: <%=toppx%>px">
<table cellspacing="0" cellpadding="0" width="100%" border="0">
  
        <tr>
            <td valign="bottom">
            <font size="1" face="Helvetica">
            Betaling med frig&#248;rende virkning kan alene ske til : AL Finans A/S, Edithsvej 2 A, Postboks 352, 2600 Glostrup. </font></font><br />
            <strong><font size="1" face="Helvetica-Bold"><font size="1" face="Helvetica-Bold">Bankkonto 5301 0931482<br />
            </strong></font>
            <font size="1" face="Helvetica"><font size="1" face="Helvetica">Vor fordring i henhold til n&#230;rv&#230;rende faktura er overdraget til AL Finans A/S - Levering har fundet sted.Ved betaling bedes leverand&#248;r samt fakturanummer anf&#248;rt. Eventuelle uoverensstemmelser bedes meddelt til AL Finans A/S, Telefon 43 43 49 90</font>
            </td>
            </tr>
            </table>

</div>

<%
case else %>
<%if len(trim(strKomm)) <> 0 then %>
<b><%=txt_033 %>:</b><br><%=strKomm%>
<%end if %>
<%end select %>

<br>&nbsp;</td>
</tr>
</table>

<%
end sub



'**** Brevhoved *************

'**** Brevfod ***************







function rykkergebyrer()

               
               strSQLryk = "SELECT rykkertxt, rykkerantal, rykkerbelob, rykkerdato, "_
               &" fr.valuta, fr.kurs, v.valutakode "_
               &" FROM faktura_rykker fr "_
               &" LEFT JOIN valutaer v ON (v.id = fr.valuta) WHERE fakid = "&  id
               oRec.open strSQLryk, oConn, 3
               r = 0
               while not oRec.EOF 
               
               if r = -10 then%>
                <tr>
               <td align=right style="padding-right:2px; padding-top:5px; border-bottom:1px #999999 solid;"><b><%=txt_013 %></b></td>
               
               <td colspan=5 style="padding-left:5px; padding-top:5px; border-bottom:1px #999999 solid;"><b><%=txt_014 %></b></td>
               <td style="border-bottom:1px #999999 solid; padding-top:5px;"><b><%=txt_42 %></b></td>
               <td align=right style="border-bottom:1px #999999 solid; padding-top:5px;"><b><%=txt_017 %></b></td>
               
               </tr>
               <%end if
              
               
               %><tr>
               <td valign=top align=right style="padding-right:7px;"><%=formatnumber(oRec("rykkerantal"), 2) %></td>
               <td>
                   &nbsp;</td>
              
               <td colspan=4 style="padding-left:5px;"><%=formatdatetime(oRec("rykkerdato"), 1) %> - <%=oRec("rykkertxt") %></td>
               <td>
                   &nbsp;</td>
               <td valign=top align=right><%=formatnumber(oRec("rykkerbelob"), 2) & " "& oRec("valutakode") %></td>
               </tr>
               <%
             
               intFaktureretBelob = intFaktureretBelob + oRec("rykkerbelob")
               
               r = r + 1 
               oRec.movenext
               wend
               oRec.close
               
              
               

end function

%>
