<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/xml/tsa_xml_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/helligdage_func.asp"-->
<%
if session("user") = "" then
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<%
	errortype = 5
	call showError(errortype)
	else
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->


   
   <script>
       $(window).load(function() {
           // run code
           $("#loadbar").hide(1000);
       }); 
   </script>
   
   
  
   <body onLoad="self.focus()">
    <div name="medarbafstem_aar" id="medarbafstem_aar" style="position:relative; left:20px; top:20px; display:; visibility:visible; width:340px; background-color:#FFFFFF; padding:5px; border:1px #8caae6 solid;">
    
    <h4>Afstemning status:</h4>
    
   
    
    
    <%
	thisfile = "afstem_tot.asp"
	if len(trim(request("usemrn"))) <> 0 then
	usemrn = request("usemrn")
	else
	usemrn = 0
	end if
	
	intMid = usemrn
	
	%>
	 <a href="afstem_tot.asp?usemrn=<%=intMid %>&show=5" class=vmenu><%=global_txt_163 %></a>&nbsp;&nbsp;|&nbsp;&nbsp;
	<a href="afstem_tot.asp?usemrn=<%=intMid %>&show=12" class=vmenu>År -> Dato</a>&nbsp;&nbsp;|&nbsp;&nbsp;
	<a href="afstem_tot.asp?usemrn=<%=intMid %>&show=4" class=vmenu>Ferie <%=global_txt_163 %></a>
	<br />
	<br />

	
	<div id="loadbar" style="position:absolute; display:; visibility:visible; top:80px; left:20px; width:300px; background-color:#ffffff; border:1px #cccccc solid; padding:2px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	</td></tr></table>

	</div>
	
	
	<%
	show = request("show")
	
	
	'call licensStartDato()
	
	strSQL = "SELECT ansatdato, mnavn, mnr, init FROM medarbejdere WHERE mid = "& usemrn
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	ansatdato = oRec("ansatdato")
	mnavn = oRec("mnavn") & " ("& oRec("mnr") &")"
	init = oRec("init")
	end if
	oRec.close
	
	call licensStartDato()
	ltoStDato = startDatoDag &"/"& startDatoMd &"/"& startDatoAar
	
	if cDate(ltoStDato) > cDate(ansatdato) then
	startDato = datepart("yyyy", ltoStDato) &"/"& datepart("m", ltoStDato) &"/"& datepart("d", ltoStDato)
    else
    startDato = datepart("yyyy", ansatdato) &"/"& datepart("m", ansatdato) &"/"& datepart("d", ansatdato)
    
    end if
    
    ddato = datepart("d", now) &"/"& datepart("m", now) &"/"& datepart("yyyy", now)
    slutDato = datepart("yyyy", ddato) &"/"& datepart("m", ddato) &"/"& datepart("d", ddato)
   
	'dateDiff("ww", slutDato, slutDato, 2, 2) 
	
	'Response.Write startDato  & "ltoStDato:" & ltoStDato & " Weeks:" & dateDiff("ww", startDato, slutDato, 2, 2)  & "<br>"
	'Response.Flush
	Response.Write "<b>"& mnavn & "</b>"
	if len(trim(init)) <> 0 then
	Response.write " - "& init 
	end if
	
	
	
	call akttyper2009(6)
	akttype_sel = aktiveTyper
	
	select case show 
	case 5
	
	Response.flush
	
	Response.Write "<br><b>Total.</b> ("& dateDiff("ww", startDato, slutDato, 2, 2)  & " uger)<br>"
	call medarbafstem(usemrn, startDato, slutDato, 5, akttype_sel)
    response.flush
    
    case 12
    
      
    Response.Write "<br><b>År -> Dato:</b><br>"
    
    monthThis = month(now)
    
    for mth = -1 to monthThis-2
    
    if mth <> -1 then
    datoB = dateadd("m", -mth, ddato)
    datoB = "1/"& month(datoB) &"/"& year(datoB)
    datoB = dateadd("d", -1, datoB)
    datoA = dateadd("m", -(mth+1), ddato)
    else
    datoB = day(ddato)&"/"& month(ddato) &"/"& year(ddato)
    datoA = ddato
    end if
    
    slutDatoLastm_B = datepart("yyyy", datoB) &"/"& datepart("m", datoB) &"/"& datepart("d", datoB)
    slutDatoLastm_A = datepart("yyyy", datoA) &"/"& datepart("m", datoA) &"/1"
    
    call medarbafstem(usemrn, slutDatoLastm_A, slutDatoLastm_B, 5, akttype_sel)
    'Response.Write "<br>"
	response.flush
	
	next
	
	'call medarbafstem(usemrn, slutDatoLastm_2, slutDatoLastm_1, 5, akttype_sel)
	'response.flush
	'Response.Write "<br>: " & formatdatetime(ddato_lm3, 2)  & "<br>"
	'call medarbafstem(usemrn, startDato, slutDatoLastm_3, 5, akttype_sel)
	'response.flush 
    
    Response.Write "<br>"
    case 4,5
    
    '** Ferie / Ferie år + feriefridag **'
    
    strAar = year(now)
    
    '*** Ferie år ***
	if cdate(now) >= cdate("1-5-"& strAar) AND cdate(now) <= cdate("31-12-"& strAar) then
	'Response.Write "OK her"
	ferieaarST = strAar &"/5/1"
	ferieaarSL = strAar+1 &"/4/30"
	else
	ferieaarST = strAar-1 &"/5/1"
	ferieaarSL = strAar &"/4/30"
	end if
	
    
    
    akttype_sel = "#11#, #12#, #13#, #14#, #15#, #16#, #17#, #18#, #19#, "
    call medarbafstem(usemrn, ferieaarST, ferieaarSL, 4, akttype_sel)
    
    end select%>
    
    
    </div>
    
    <br /><br />
    <%
    
    itop = 30
    ileft = 20
    iwdt = 430
    call sideinfo(itop,ileft,iwdt)
    
   
			call akttyper2009(3)
			%>
			</table>
   
   
   </td>
   </tr>
   </table>
   </div>
    
    
    
    
    
    <br /><br />
    
    
    
     <br /><br /><br /><br /><br /><br /><br /><br />
    <br /><br /><br /><br /><br /> <br /><br /><br /><br /><br /> <br /><br /><br /><br /><br /><br /><br /><br />
    <br /><br /><br /><br /><br /><br /><br /><br />
   <br /><br /><br /><br /><br /><br /><br /><br /> <br /><br /><br /><br /><br /><br /><br /><br />
    <br /><br /><br /><br /><br /> <br /><br /><br /><br /><br /> <br /><br /><br /><br /><br /><br /><br /><br />
    <br /><br /><br /><br /><br /><br /><br /><br />
   <br /><br /><br /><br /><br /><br /><br /><br /> <br /><br /><br /><br /><br /><br /><br /><br />
    <br /><br /><br /><br /><br /> <br /><br /><br /><br /><br /> <br /><br /><br /><br /><br /><br /><br /><br />
    <br /><br /><br /><br /><br /><br /><br /><br />
   <br /><br /><br /><br /><br /><br /><br /><br />
    <br /><br /><br /><br /><br /><br /><br /><br />
    <br /><br /><br /><br /><br /> <br /><br /><br /><br /><br /> <br /><br /><br /><br /><br /><br /><br /><br />
    <br /><br /><br /><br /><br /><br /><br /><br />
   <br /><br /><br /><br /><br /><br /><br /><br />&nbsp;
 

   
  
   
<%end if 'validering %>
<!--#include file="../inc/regular/footer_inc.asp"-->
