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
	<head runat="server">
		<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.1)" />
		<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.1)" />
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>
		<script type="text/javascript" src="../Scripts/Dialog/Dialog_ColorPicker.js"></script>
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<style type="text/css">
			.colorcell
			{
				width:16px;
				height:17px;
				cursor:hand;
			}
			.colordiv,.customdiv
			{
				border:solid 1px #808080;
				width:16px;
				height:17px;
				font-size:1px;
			}
		</style>
		<title><%= GetString("NamedColors") %></title>
		<script>
								
		var OxOb42b=["Green","#008000","Lime","#00FF00","Teal","#008080","Aqua","#00FFFF","Navy","#000080","Blue","#0000FF","Purple","#800080","Fuchsia","#FF00FF","Maroon","#800000","Red ","#FF0000","Olive","#808000","Yellow","#FFFF00","White","#FFFFFF","Silver","#C0C0C0","Gray","#808080","Black","#000000","DarkOliveGreen","#556B2F","DarkGreen","#006400","DarkSlateGray","#2F4F4F","SlateGray","#708090","DarkBlue","#00008B","MidnightBlue","#191970","Indigo","#4B0082","DarkMagenta","#8B008B","Brown","#A52A2A","DarkRed","#8B0000","Sienna","#A0522D","SaddleBrown","#8B4513","DarkGoldenrod","#B8860B","Beige","#F5F5DC","HoneyDew","#F0FFF0","DimGray","#696969","OliveDrab","#6B8E23","ForestGreen","#228B22","DarkCyan","#008B8B","LightSlateGray","#778899","MediumBlue","#0000CD","DarkSlateBlue","#483D8B","DarkViolet","#9400D3","MediumVioletRed","#C71585","IndianRed","#CD5C5C","Firebrick","#B22222","Chocolate","#D2691E","Peru","#CD853F","Goldenrod","#DAA520","LightGoldenrodYellow","#FAFAD2","MintCream","#F5FFFA","DarkGray","#A9A9A9","YellowGreen","#9ACD32","SeaGreen","#2E8B57","CadetBlue","#5F9EA0","SteelBlue","#4682B4","RoyalBlue","#4169E1","BlueViolet","#8A2BE2","DarkOrchid","#9932CC","DeepPink","#FF1493","RosyBrown","#BC8F8F","Crimson","#DC143C","DarkOrange","#FF8C00","BurlyWood","#DEB887","DarkKhaki","#BDB76B","LightYellow","#FFFFE0","Azure","#F0FFFF","LightGray","#D3D3D3","LawnGreen","#7CFC00","MediumSeaGreen","#3CB371","LightSeaGreen","#20B2AA","DeepSkyBlue","#00BFFF","DodgerBlue","#1E90FF","SlateBlue","#6A5ACD","MediumOrchid","#BA55D3","PaleVioletRed","#DB7093","Salmon","#FA8072","OrangeRed","#FF4500","SandyBrown","#F4A460","Tan","#D2B48C","Gold","#FFD700","Ivory","#FFFFF0","GhostWhite","#F8F8FF","Gainsboro","#DCDCDC","Chartreuse","#7FFF00","LimeGreen","#32CD32","MediumAquamarine","#66CDAA","DarkTurquoise","#00CED1","CornflowerBlue","#6495ED","MediumSlateBlue","#7B68EE","Orchid","#DA70D6","HotPink","#FF69B4","LightCoral","#F08080","Tomato","#FF6347","Orange","#FFA500","Bisque","#FFE4C4","Khaki","#F0E68C","Cornsilk","#FFF8DC","Linen","#FAF0E6","WhiteSmoke","#F5F5F5","GreenYellow","#ADFF2F","DarkSeaGreen","#8FBC8B","Turquoise","#40E0D0","MediumTurquoise","#48D1CC","SkyBlue","#87CEEB","MediumPurple","#9370DB","Violet","#EE82EE","LightPink","#FFB6C1","DarkSalmon","#E9967A","Coral","#FF7F50","NavajoWhite","#FFDEAD","BlanchedAlmond","#FFEBCD","PaleGoldenrod","#EEE8AA","Oldlace","#FDF5E6","Seashell","#FFF5EE","PaleGreen","#98FB98","SpringGreen","#00FF7F","Aquamarine","#7FFFD4","PowderBlue","#B0E0E6","LightSkyBlue","#87CEFA","LightSteelBlue","#B0C4DE","Plum","#DDA0DD","Pink","#FFC0CB","LightSalmon","#FFA07A","Wheat","#F5DEB3","Moccasin","#FFE4B5","AntiqueWhite","#FAEBD7","LemonChiffon","#FFFACD","FloralWhite","#FFFAF0","Snow","#FFFAFA","AliceBlue","#F0F8FF","LightGreen","#90EE90","MediumSpringGreen","#00FA9A","PaleTurquoise","#AFEEEE","LightCyan","#E0FFFF","LightBlue","#ADD8E6","Lavender","#E6E6FA","Thistle","#D8BFD8","MistyRose","#FFE4E1","Peachpuff","#FFDAB9","PapayaWhip","#FFEFD5"];var colorlist=[{n:OxOb42b[0x0],h:OxOb42b[0x1]},{n:OxOb42b[0x2],h:OxOb42b[0x3]},{n:OxOb42b[0x4],h:OxOb42b[0x5]},{n:OxOb42b[0x6],h:OxOb42b[0x7]},{n:OxOb42b[0x8],h:OxOb42b[0x9]},{n:OxOb42b[0xa],h:OxOb42b[0xb]},{n:OxOb42b[0xc],h:OxOb42b[0xd]},{n:OxOb42b[0xe],h:OxOb42b[0xf]},{n:OxOb42b[0x10],h:OxOb42b[0x11]},{n:OxOb42b[0x12],h:OxOb42b[0x13]},{n:OxOb42b[0x14],h:OxOb42b[0x15]},{n:OxOb42b[0x16],h:OxOb42b[0x17]},{n:OxOb42b[0x18],h:OxOb42b[0x19]},{n:OxOb42b[0x1a],h:OxOb42b[0x1b]},{n:OxOb42b[0x1c],h:OxOb42b[0x1d]},{n:OxOb42b[0x1e],h:OxOb42b[0x1f]}];var colormore=[{n:OxOb42b[0x20],h:OxOb42b[0x21]},{n:OxOb42b[0x22],h:OxOb42b[0x23]},{n:OxOb42b[0x24],h:OxOb42b[0x25]},{n:OxOb42b[0x26],h:OxOb42b[0x27]},{n:OxOb42b[0x28],h:OxOb42b[0x29]},{n:OxOb42b[0x2a],h:OxOb42b[0x2b]},{n:OxOb42b[0x2c],h:OxOb42b[0x2d]},{n:OxOb42b[0x2e],h:OxOb42b[0x2f]},{n:OxOb42b[0x30],h:OxOb42b[0x31]},{n:OxOb42b[0x32],h:OxOb42b[0x33]},{n:OxOb42b[0x34],h:OxOb42b[0x35]},{n:OxOb42b[0x36],h:OxOb42b[0x37]},{n:OxOb42b[0x38],h:OxOb42b[0x39]},{n:OxOb42b[0x3a],h:OxOb42b[0x3b]},{n:OxOb42b[0x3c],h:OxOb42b[0x3d]},{n:OxOb42b[0x3e],h:OxOb42b[0x3f]},{n:OxOb42b[0x40],h:OxOb42b[0x41]},{n:OxOb42b[0x42],h:OxOb42b[0x43]},{n:OxOb42b[0x44],h:OxOb42b[0x45]},{n:OxOb42b[0x46],h:OxOb42b[0x47]},{n:OxOb42b[0x48],h:OxOb42b[0x49]},{n:OxOb42b[0x4a],h:OxOb42b[0x4b]},{n:OxOb42b[0x4c],h:OxOb42b[0x4d]},{n:OxOb42b[0x4e],h:OxOb42b[0x4f]},{n:OxOb42b[0x50],h:OxOb42b[0x51]},{n:OxOb42b[0x52],h:OxOb42b[0x53]},{n:OxOb42b[0x54],h:OxOb42b[0x55]},{n:OxOb42b[0x56],h:OxOb42b[0x57]},{n:OxOb42b[0x58],h:OxOb42b[0x59]},{n:OxOb42b[0x5a],h:OxOb42b[0x5b]},{n:OxOb42b[0x5c],h:OxOb42b[0x5d]},{n:OxOb42b[0x5e],h:OxOb42b[0x5f]},{n:OxOb42b[0x60],h:OxOb42b[0x61]},{n:OxOb42b[0x62],h:OxOb42b[0x63]},{n:OxOb42b[0x64],h:OxOb42b[0x65]},{n:OxOb42b[0x66],h:OxOb42b[0x67]},{n:OxOb42b[0x68],h:OxOb42b[0x69]},{n:OxOb42b[0x6a],h:OxOb42b[0x6b]},{n:OxOb42b[0x6c],h:OxOb42b[0x6d]},{n:OxOb42b[0x6e],h:OxOb42b[0x6f]},{n:OxOb42b[0x70],h:OxOb42b[0x71]},{n:OxOb42b[0x72],h:OxOb42b[0x73]},{n:OxOb42b[0x74],h:OxOb42b[0x75]},{n:OxOb42b[0x76],h:OxOb42b[0x77]},{n:OxOb42b[0x78],h:OxOb42b[0x79]},{n:OxOb42b[0x7a],h:OxOb42b[0x7b]},{n:OxOb42b[0x7c],h:OxOb42b[0x7d]},{n:OxOb42b[0x7e],h:OxOb42b[0x7f]},{n:OxOb42b[0x80],h:OxOb42b[0x81]},{n:OxOb42b[0x82],h:OxOb42b[0x83]},{n:OxOb42b[0x84],h:OxOb42b[0x85]},{n:OxOb42b[0x86],h:OxOb42b[0x87]},{n:OxOb42b[0x88],h:OxOb42b[0x89]},{n:OxOb42b[0x8a],h:OxOb42b[0x8b]},{n:OxOb42b[0x8c],h:OxOb42b[0x8d]},{n:OxOb42b[0x8e],h:OxOb42b[0x8f]},{n:OxOb42b[0x90],h:OxOb42b[0x91]},{n:OxOb42b[0x92],h:OxOb42b[0x93]},{n:OxOb42b[0x94],h:OxOb42b[0x95]},{n:OxOb42b[0x96],h:OxOb42b[0x97]},{n:OxOb42b[0x98],h:OxOb42b[0x99]},{n:OxOb42b[0x9a],h:OxOb42b[0x9b]},{n:OxOb42b[0x9c],h:OxOb42b[0x9d]},{n:OxOb42b[0x9e],h:OxOb42b[0x9f]},{n:OxOb42b[0xa0],h:OxOb42b[0xa1]},{n:OxOb42b[0xa2],h:OxOb42b[0xa3]},{n:OxOb42b[0xa4],h:OxOb42b[0xa5]},{n:OxOb42b[0xa6],h:OxOb42b[0xa7]},{n:OxOb42b[0xa8],h:OxOb42b[0xa9]},{n:OxOb42b[0xaa],h:OxOb42b[0xab]},{n:OxOb42b[0xac],h:OxOb42b[0xad]},{n:OxOb42b[0xae],h:OxOb42b[0xaf]},{n:OxOb42b[0xb0],h:OxOb42b[0xb1]},{n:OxOb42b[0xb2],h:OxOb42b[0xb3]},{n:OxOb42b[0xb4],h:OxOb42b[0xb5]},{n:OxOb42b[0xb6],h:OxOb42b[0xb7]},{n:OxOb42b[0xb8],h:OxOb42b[0xb9]},{n:OxOb42b[0xba],h:OxOb42b[0xbb]},{n:OxOb42b[0xbc],h:OxOb42b[0xbd]},{n:OxOb42b[0xbe],h:OxOb42b[0xbf]},{n:OxOb42b[0xc0],h:OxOb42b[0xc1]},{n:OxOb42b[0xc2],h:OxOb42b[0xc3]},{n:OxOb42b[0xc4],h:OxOb42b[0xc5]},{n:OxOb42b[0xc6],h:OxOb42b[0xc7]},{n:OxOb42b[0xc8],h:OxOb42b[0xc9]},{n:OxOb42b[0xca],h:OxOb42b[0xcb]},{n:OxOb42b[0xcc],h:OxOb42b[0xcd]},{n:OxOb42b[0xce],h:OxOb42b[0xcf]},{n:OxOb42b[0xd0],h:OxOb42b[0xd1]},{n:OxOb42b[0xd2],h:OxOb42b[0xd3]},{n:OxOb42b[0xd4],h:OxOb42b[0xd5]},{n:OxOb42b[0xd6],h:OxOb42b[0xd7]},{n:OxOb42b[0xd8],h:OxOb42b[0xd9]},{n:OxOb42b[0xda],h:OxOb42b[0xdb]},{n:OxOb42b[0xdc],h:OxOb42b[0xdd]},{n:OxOb42b[0x9c],h:OxOb42b[0x9d]},{n:OxOb42b[0xde],h:OxOb42b[0xdf]},{n:OxOb42b[0xe0],h:OxOb42b[0xe1]},{n:OxOb42b[0xe2],h:OxOb42b[0xe3]},{n:OxOb42b[0xe4],h:OxOb42b[0xe5]},{n:OxOb42b[0xe6],h:OxOb42b[0xe7]},{n:OxOb42b[0xe8],h:OxOb42b[0xe9]},{n:OxOb42b[0xea],h:OxOb42b[0xeb]},{n:OxOb42b[0xec],h:OxOb42b[0xed]},{n:OxOb42b[0xee],h:OxOb42b[0xef]},{n:OxOb42b[0xf0],h:OxOb42b[0xf1]},{n:OxOb42b[0xf2],h:OxOb42b[0xf3]},{n:OxOb42b[0xf4],h:OxOb42b[0xf5]},{n:OxOb42b[0xf6],h:OxOb42b[0xf7]},{n:OxOb42b[0xf8],h:OxOb42b[0xf9]},{n:OxOb42b[0xfa],h:OxOb42b[0xfb]},{n:OxOb42b[0xfc],h:OxOb42b[0xfd]},{n:OxOb42b[0xfe],h:OxOb42b[0xff]},{n:OxOb42b[0x100],h:OxOb42b[0x101]},{n:OxOb42b[0x102],h:OxOb42b[0x103]},{n:OxOb42b[0x104],h:OxOb42b[0x105]},{n:OxOb42b[0x106],h:OxOb42b[0x107]},{n:OxOb42b[0x108],h:OxOb42b[0x109]},{n:OxOb42b[0x10a],h:OxOb42b[0x10b]},{n:OxOb42b[0x10c],h:OxOb42b[0x10d]},{n:OxOb42b[0x10e],h:OxOb42b[0x10f]},{n:OxOb42b[0x110],h:OxOb42b[0x111]}];
		
		</script>
	</head>
	<body>
		<div id="container">
			<div class="tab-pane-control tab-pane" id="tabPane1">
				<div class="tab-row">
					<h2 class="tab">
						<a tabindex="-1" href='colorpicker.asp?<%=GetDialogQueryString%>'>
							<span style="white-space:nowrap;">
								<%= GetString("WebPalette") %>
							</span>
						</a>
					</h2>
					<h2 class="tab selected">
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
					<table class="colortable" align="center">
						<tr>
							<td colspan="16" height="16"><p align="left">Basic:
								</p>
							</td>
						</tr>
						<tr>
							<script>
								var OxO8da8=["length","\x3Ctd class=\x27colorcell\x27\x3E\x3Cdiv class=\x27colordiv\x27 style=\x27background-color:","\x27 title=\x27"," ","\x27 cname=\x27","\x27 cvalue=\x27","\x27\x3E\x3C/div\x3E\x3C/td\x3E",""];var arr=[];for(var i=0x0;i<colorlist[OxO8da8[0x0]];i++){ arr.push(OxO8da8[0x1]) ; arr.push(colorlist[i].n) ; arr.push(OxO8da8[0x2]) ; arr.push(colorlist[i].n) ; arr.push(OxO8da8[0x3]) ; arr.push(colorlist[i].h) ; arr.push(OxO8da8[0x4]) ; arr.push(colorlist[i].n) ; arr.push(OxO8da8[0x5]) ; arr.push(colorlist[i].h) ; arr.push(OxO8da8[0x6]) ;} ; document.write(arr.join(OxO8da8[0x7])) ;
							</script>
						</tr>
						<tr>
							<td colspan="16" height="12"><p align="left"></p>
							</td>
						</tr>
						<tr>
							<td colspan="16"><p align="left">Additional:
								</p>
							</td>
						</tr>
						<script>
							var OxO45db=["length","\x3Ctr\x3E","\x3Ctd class=\x27colorcell\x27\x3E\x3Cdiv class=\x27colordiv\x27 style=\x27background-color:","\x27 title=\x27"," ","\x27 cname=\x27","\x27 cvalue=\x27","\x27\x3E\x3C/div\x3E\x3C/td\x3E","\x3C/tr\x3E",""];var arr=[];for(var i=0x0;i<colormore[OxO45db[0x0]];i++){if(i%0x10==0x0){ arr.push(OxO45db[0x1]) ;} ; arr.push(OxO45db[0x2]) ; arr.push(colormore[i].n) ; arr.push(OxO45db[0x3]) ; arr.push(colormore[i].n) ; arr.push(OxO45db[0x4]) ; arr.push(colormore[i].h) ; arr.push(OxO45db[0x5]) ; arr.push(colormore[i].n) ; arr.push(OxO45db[0x6]) ; arr.push(colormore[i].h) ; arr.push(OxO45db[0x7]) ;if(i%0x10==0xf){ arr.push(OxO45db[0x8]) ;} ;} ;if(colormore%0x10>0x0){ arr.push(OxO45db[0x8]) ;} ; document.write(arr.join(OxO45db[0x9])) ;
						</script>
						<tr>
							<td colspan="16" height="8">
							</td>
						</tr>
						<tr>
							<td colspan="16" height="12">
								<input checked id="CheckboxColorNames" style="width: 16px; height: 20px" type="checkbox">
								<span style="width: 118px;">Use color names</span>
							</td>
						</tr>
						<tr>
							<td colspan="16" height="12">
							</td>
						</tr>
						<tr>
							<td colspan="16" valign="middle" height="24">
							<span style="height:24px;width:50px;vertical-align:middle;">Color : </span>&nbsp;
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

