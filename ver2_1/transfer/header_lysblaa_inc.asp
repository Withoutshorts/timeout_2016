<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!-- "http://www.w3.org/TR/html4/loose.dtd" -->
<html>
<head>
	<title>TimeOut - Tid & Overblik</title>
	<!--<LINK REL="SHORTCUT ICON" HREF="https://outzource.dk/favicon.ico">-->
	
	<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print.css" />
    <link href="../inc/jquery/jquery-ui-1.7.1.custom.css" rel="stylesheet" type="text/css" />
	<script src="../inc/jquery/jquery-1.3.2.min.js" type="text/javascript"></script>
	<script src="../inc/jquery/jquery-ui-1.7.1.custom.min.js" type="text/javascript"></script>
	<script src="../inc/jquery/timeout.jquery.js" type="text/javascript"></script>
	<script src="../inc/jquery/jquery.coookie.js" type="text/javascript"></script>
	
	
	
	<script>
function showabout(){
document.all["about"].style.visibility = "visible";
}
function hideabout(){
document.all["about"].style.visibility = "hidden";
}
function NewWin(url)    {
window.open(url, 'Help', 'width=700,height=580,scrollbars=yes,toolbar=no');    
}
function NewWin_cal(url)    {
window.open(url, 'Calender', 'width=300,height=380,scrollbars=yes,toolbar=no');    
}
function NewWin_popupaktion(url)    {
window.open(url, 'Calender', 'width=410, height=650,scrollbars=yes,toolbar=no');    
}
function NewWin_help(url)    {
window.open(url, 'Help', 'width=650,height=580,scrollbars=yes,toolbar=no');    
}
function NewWin_large(url)    {
window.open(url, 'Help', 'width=980,height=600,scrollbars=yes,toolbar=no');    
}
function NewWin_todo(url)    {
window.open(url, 'Opgavelisten', 'width=650,height=650,scrollbars=yes,toolbar=no');    
}
//Alle ovenst�ende popup's �ndres til nedenst�ende//
//G�r ogs� left og top til variable//
function popUp(URL,width,height,left,top) {
		window.open(URL, 'navn', 'left='+left+',top='+top+',toolbar=0,scrollbars=1,location=0,statusbar=1,menubar=0,resizable=1,width=' + width + ',height=' + height + '');
	}

function NewWin_popupcal(url)    {
	window.open(url, 'Calpick', 'width=250,height=250,scrollbars=no,toolbar=no');    
}


function bgcolthisMON(divid){
document.getElementById(""+ divid +"").style.backgroundColor = "#ffffff";
}

function bgcolthisON(divid){
document.getElementById(""+ divid +"").style.backgroundColor = "#5C75AA";
}


function bgcolthisONcol(divid, bgcol) {
    document.getElementById("" + divid + "").style.backgroundColor = bgcol;
}

function bgcolthisOFF(divid, bgcol){
document.getElementById(""+ divid +"").style.backgroundColor = bgcol;
}

function screenw() {
//alert(screen.width)
    if (document.getElementById("screenw") == null) {
    } else {
    wdt = screen.width - 20
    document.getElementById("screenw").style.width = wdt;
    }
}
</script>
	
</head>

<%
if len(trim(browstype_client)) <> 0 then
call browsertype()
else
browstype_client = "ie"
end if

if browstype_client = "mz" then
%>
<body topmargin="0" leftmargin="0" class="regular" onload="screenw()";>
<%else %>
<body topmargin="0" leftmargin="0" class="regular">
<%end if %>
<!--=======================================================
TimeOut er t�nkt og produceret af OutZourCE 2002 - 2009
www.OutZourCE.dk 
Tlf: +45 2684 2000
timeout@outzource.dk
========================================================-->

