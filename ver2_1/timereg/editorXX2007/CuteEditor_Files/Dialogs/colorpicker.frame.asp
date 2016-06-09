<!-- #include file = "Include_GetString.asp" -->
<%
dim GetDialogQueryString
Theme="Custom"

GetDialogQueryString = "Theme=Custom"
if Request.QueryString("Dialog") = "Standard" then
    GetDialogQueryString=GetDialogQueryString & "&Dialog=Standard"
End If
if Request.QueryString("setting") <> "" then
    GetDialogQueryString=GetDialogQueryString & "&setting=" & Request.QueryString("setting")
end if 
%>
<html xmlns="http://www.w3.org/1999/xhtml">
	<head ID="Head1">
		<title><%= GetString("WebPalette") %></title>
		<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.1)" />
		<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.1)" />
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>
		<script type="text/javascript" src="../Scripts/Dialog/Dialog_ColorPicker.js"></script>
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<style type="text/css">
			.colorcell
			{
				width:22px;
				height:11px;
				cursor:hand;
			}
			.colordiv
			{
				border:solid 1px #808080;
				width:22px;
				height:11px;
				font-size:1px;
			}
		</style>
		<script>
var OxO9263=["0","#","length","\x3Ctr\x3E","\x3Ctd class=\x27colorcell\x27\x3E\x3Cdiv class=\x27colordiv\x27 style=\x27background-color:","\x27 cvalue=\x27","\x27 title=\x27","\x27\x3E\x26nbsp;\x3C/div\x3E\x3C/td\x3E","\x3C/tr\x3E"]; function DoubleHex(Ox5c){if(Ox5c<0x10){return OxO9263[0x0]+Ox5c.toString(0x10);} ;return Ox5c.toString(0x10);}  ; function ToHexString(Ox5e,Ox5f,Ox60){return (OxO9263[0x1]+DoubleHex(Ox5e*0x33)+DoubleHex(Ox5f*0x33)+DoubleHex(Ox60*0x33)).toUpperCase();}  ; function MakeHex(z,x,y){var Ox65=z%0x2;var Ox17=(z-Ox65)/0x2; z=Ox65*0x3+Ox17 ;if(z<0x3){ x=0x5-x ;} ;if(z==0x1||z==0x4){ y=0x5-y ;} ;return ToHexString(0x5-y,0x5-x,0x5-z);}  ;var colors= new Array(0xd8);for(var z=0x0;z<0x6;z++){for(var x=0x0;x<0x6;x++){for(var y=0x0;y<0x6;y++){var hex=MakeHex(z,x,y);var xx=(z%0x2)*0x6+x;var yy=Math.floor(z/0x2)*0x6+y; colors[yy*0xc+xx]=hex ;} ;} ;} ;var arr=[];for(var i=0x0;i<colors[OxO9263[0x2]];i++){if(i%0xc==0x0){ arr.push(OxO9263[0x3]) ;} ; arr.push(OxO9263[0x4]) ; arr.push(colors[i]) ; arr.push(OxO9263[0x5]) ; arr.push(colors[i]) ; arr.push(OxO9263[0x6]) ; arr.push(colors[i]) ; arr.push(OxO9263[0x7]) ;if(i%0xc==0xb){ arr.push(OxO9263[0x8]) ;} ;} ;
		</script>
	</head>
	<body>
	<div id="container">
			<div class="tab-pane-control tab-pane" id="tabPane1">
				<div class="tab-row">
					<h2 class="tab selected">
						<a tabindex="-1" href='colorpicker.asp?<%=GetDialogQueryString%>'>
							<span style="white-space:nowrap;">
								<%= GetString("WebPalette") %>
							</span>
						</a>
					</h2>
					<h2 class="tab">
							<a tabindex="-1" href='colorpicker_basic.asp?<%=GetDialogQueryString%>'>
								<span style="white-space:nowrap;">
									<%= GetString("NamedColors") %>
								</span>
							</a>
					</h2>
					<h2 class="tab">
							<a tabindex="-1" href='colorpicker_more.asp?<%=GetDialogQueryString%>'>
								<span style="white-space:nowrap;">
									<%= GetString("CustomColor") %>
								</span>
							</a>
					</h2>
				</div>
				<div class="tab-page">
					<table cellSpacing='2' cellPadding="1" align="center">
						<script>
							var OxO181c=[""]; document.write(arr.join(OxO181c[0x0])) ;
						</script>
						<tr>
							<td colspan="12" height="12"><p align="left"></p>
							</td>
						</tr>
						<tr>
							<td colspan="12" valign="middle" height="24">								
					<span style="height:24px;width:50px;vertical-align:middle;"><%= GetString("Color") %>: </span>&nbsp;
					<input type="text" id="divpreview" size="7" maxlength="7" style="width:180px;height:24px;border:#a0a0a0 1px solid; Padding:4;"/>
							</td>
						</tr>
					</table>
				</div>
			</div>
		<div id="container-bottom">
			<input type="button" id="buttonok" value="<%= GetString("OK") %>" class="formbutton" style="width:70px"	onclick="do_insert();" /> 
			&nbsp;&nbsp;&nbsp;&nbsp; 
			<input type="button" id="buttoncancel" value="<%= GetString("Cancel") %>" class="formbutton" style="width:70px"	onclick="do_Close();" />
		</div>
	</div>
	</body>
</html>