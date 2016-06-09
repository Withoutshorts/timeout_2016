<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>timeOut 2.1</title>
<script>

// (C) 2001 www.CodeLifter.com
// http://www.codelifter.com
// Free for all users, but leave in this  header

// set the page to go to...
url = "login.asp";

// set how fast to expand horizontally
// lower is slower
var speedX = 37;

// set how fast to expand vertically
// lower is slower
var speedY = 25; 

// set background color of "Loading..." screen
var bgColor = "#fff8dc";

// set text color of "Loading..." screen
var txtColor = "#000000";


// do not edit below this line
// ---------------------------

if (document.all){
var wide = window.screen.availWidth;
var high = window.screen.availHeight;
}

function andBoom(){
  if (document.all){
    var Boomer = window.open("","BoomWindow","fullscreen");
    Boomer.document.write('<HTML><BODY BGCOLOR='+bgColor+' SCROLL=NO><FONT FACE=ARIAL COLOR='+txtColor+'></FONT></BODY></HTML>');
    Boomer.focus();
    for (H=1; H<high; H+= speedY){
         Boomer.resizeTo(1,H);
    }
    for (W=1; W<wide; W+= speedX){
         Boomer.resizeTo(W+25,H-10);
    }
    Boomer.location = url;
  }  
  else {
    window.open(url,"BoomWindow","");
  }
}

</script>
</head>
<body onLoad="andBoom();">
<!--#include file="inc/regular/footer_inc.asp"-->


