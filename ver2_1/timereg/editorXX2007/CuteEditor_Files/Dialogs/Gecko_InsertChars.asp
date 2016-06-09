<!-- #include file = "Include_GetString.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title><%= GetString("InsertChars") %> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
			&nbsp; </title>
		
		<meta name="content-type" content="text/html ;charset=Unicode" />
		<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.1)" />
		<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.1)" />
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<!--[if IE]>
			<link href="../style/IE.css" type="text/css" rel="stylesheet" />
		<![endif]-->
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>		
		<script type="text/javascript">
		
			var OxO6946=["","\x3Ctr\x3E","\x3Ctd style=\x27height: 20; font-size: 12px; \x27 bgcolor=white width=\x2718\x27 onClick=\x27getchar(this)\x27 onmouseover=\x27spcOver(this)\x27 onmouseout=\x27spcOut(this)\x27 title=\x27","\x27 \x3E","\x26#",";","\x3C/td\x3E","\x3C/tr\x3E","background","style","#0A246A","color","white","black","Verdana","innerHTML","Unicode","\x3CFONT CLASS=\x27UNICODE\x27\x3E","\x3Cspan style=\x27font-family:","\x27\x3E","\x3C/span\x3E"];var editor=Window_GetDialogArguments(window); function cancel(){ Window_CloseDialog(window) ;}  ;var tds=0x16; function writeChars(){var Ox7d=OxO6946[0x0];for(var i=0x21;i<0x100;){ document.write(OxO6946[0x1]) ;for(var j=0x0;j<=tds;j++){ document.write(OxO6946[0x2]+i+OxO6946[0x3]) ; document.write(OxO6946[0x4]+i+OxO6946[0x5]) ; document.write(OxO6946[0x6]) ; i++ ;} ; document.write(OxO6946[0x7]) ;} ;}  ; function spcOver(Ox80){ Ox80[OxO6946[0x9]][OxO6946[0x8]]=OxO6946[0xa] ; Ox80[OxO6946[0x9]][OxO6946[0xb]]=OxO6946[0xc] ;}  ; function spcOut(Ox80){ Ox80[OxO6946[0x9]][OxO6946[0x8]]=OxO6946[0xc] ; Ox80[OxO6946[0x9]][OxO6946[0xb]]=OxO6946[0xd] ;}  ; function getchar(obj){var Ox83;var Ox84=getFontValue()||OxO6946[0xe];if(!obj[OxO6946[0xf]]){return ;} ; Ox83=obj[OxO6946[0xf]] ;if(Ox84==OxO6946[0x10]){ Ox83=OxO6946[0x11]+obj[OxO6946[0xf]]+OxO6946[0x0] ;} else {if(Ox84!=OxO6946[0xe]){ Ox83=OxO6946[0x12]+Ox84+OxO6946[0x13]+obj[OxO6946[0xf]]+OxO6946[0x14] ;} ;} ; editor.PasteHTML(Ox83) ; Window_CloseDialog(window) ;}  ;
		</script>
	</head>
	<body>
		<div id="container">
			<table border="0" cellspacing="2" cellpadding="2" width="99%" id="Table1">
				<tr style="display:none">
					<td class="normal">
						<%= GetString("FontName") %>: 
						<input type="radio" onpropertychange="sel_font_change()" id="selfont1" name="selfont" value="" checked="checked" />
						<label for="selfont1"><%= GetString("Default") %></label> 
						<input type="radio" onpropertychange="sel_font_change()" id="selfont2" name="selfont" value="webdings" />
						<label for="selfont2">Webdings</label>
						<input type="radio" onpropertychange="sel_font_change()" id="selfont3" name="selfont" value="wingdings" />
						<label for="selfont3">Wingdings</label>
						<input type="radio" onpropertychange="sel_font_change()" id="selfont4" name="selfont" value="symbol" />
						<label for="selfont4">Symbol</label>
						<input type="radio" onpropertychange="sel_font_change()" id="selfont5" name="selfont" value="Unicode" />
						<label for="selfont5">Unicode</label>
						<script type="text/javascript">
						var OxObbe4=["selfont","length","checked","value","Verdana","display","style","charstable1","Unicode","block","none","charstable2","fontFamily"]; function getFontValue(){var Ox87=document.getElementsByName(OxObbe4[0x0]);for(var i=0x0;i<Ox87[OxObbe4[0x1]];i++){if(Ox87[i][OxObbe4[0x2]]){return Ox87[i][OxObbe4[0x3]];} ;} ;}  ; function sel_font_change(){var Ox89=getFontValue()||OxObbe4[0x4]; document.getElementById(OxObbe4[0x7])[OxObbe4[0x6]][OxObbe4[0x5]]=(Ox89!=OxObbe4[0x8]?OxObbe4[0x9]:OxObbe4[0xa]) ; document.getElementById(OxObbe4[0xb])[OxObbe4[0x6]][OxObbe4[0x5]]=(Ox89==OxObbe4[0x8]?OxObbe4[0x9]:OxObbe4[0xa]) ; document.getElementById(OxObbe4[0x7])[OxObbe4[0x6]][OxObbe4[0xc]]=Ox89 ;if(Ox89==OxObbe4[0x8]){} ;}  ;
						</script>
					</td>
				</tr>
				<tr>
					<td align="center">
						<fieldset>
							<legend>
								<%= GetString("InsertChars") %>
							</legend>
							<br />
							<table id="charstable1" width="95%" cellspacing="1" cellpadding="1" border="0"
								 style="FONT-FAMILY: Verdana; background-color:#696969; border-color:#696969; height:222">
								<script type="text/javascript">
								var OxO9580=[]; writeChars() ;
								</script>
							</table>
							<table id="charstable2" width="95%" cellspacing="1" cellpadding="1" border="0"
								style="FONT-FAMILY: Verdana; background-color:#696969; border-color:#696969;display:none; height:222">
								<script type="text/javascript">
								var OxOd196=["\x26#402;","\x26#913;","\x26#914;","\x26#915;","\x26#916;","\x26#917;","\x26#918;","\x26#919;","\x26#920;","\x26#921;","\x26#922;","\x26#923;","\x26#924;","\x26#925;","\x26#926;","\x26#927;","\x26#928;","\x26#929;","\x26#931;","\x26#932;","\x26#933;","\x26#934;","\x26#935;","\x26#936;","\x26#937;","\x26#945;","\x26#946;","\x26#947;","\x26#948;","\x26#949;","\x26#950;","\x26#951;","\x26#952;","\x26#953;","\x26#954;","\x26#955;","\x26#956;","\x26#957;","\x26#958;","\x26#959;","\x26#960;","\x26#961;","\x26#962;","\x26#963;","\x26#964;","\x26#965;","\x26#966;","\x26#967;","\x26#968;","\x26#969;","\x26#977;","\x26#978;","\x26#982;","\x26#8226;","\x26#8230;","\x26#8242;","\x26#8243;","\x26#8254;","\x26#8260;","\x26#8472;","\x26#8465;","\x26#8476;","\x26#8482;","\x26#8501;","\x26#8592;","\x26#8593;","\x26#8594;","\x26#8595;","\x26#8596;","\x26#8629;","\x26#8656;","\x26#8657;","\x26#8658;","\x26#8659;","\x26#8660;","\x26#8704;","\x26#8706;","\x26#8707;","\x26#8709;","\x26#8711;","\x26#8712;","\x26#8713;","\x26#8715;","\x26#8719;","\x26#8722;","\x26#8727;","\x26#8730;","\x26#8733;","\x26#8734;","\x26#8736;","\x26#8869;","\x26#8870;","\x26#8745;","\x26#8746;","\x26#8747;","\x26#8756;","\x26#8764;","\x26#8773;","\x26#8800;","\x26#8801;","\x26#8804;","\x26#8805;","\x26#8834;","\x26#8835;","\x26#8836;","\x26#8838;","\x26#8839;","\x26#8853;","\x26#8855;","\x26#8901;","\x26#8968;","\x26#8969;","\x26#8970;","\x26#8971;","\x26#9001;","\x26#9002;","\x26#9674;","\x26#9824;","\x26#9827;","\x26#9829;","\x26#9830;","length","\x3Ctr\x3E","\x26","\x26amp;","\x3Ctd style=\x27height: 20; font-size: 12px; \x27 bgcolor=white width=\x2718\x27 onClick=\x27getchar(this)\x27 onmouseover=\x27spcOver(this)\x27 onmouseout=\x27spcOut(this)\x27 title=\x27"," - ","\x27 \x3E","\x3C/td\x3E","\x3C/tr\x3E"];var arr=[OxOd196[0x0],OxOd196[0x1],OxOd196[0x2],OxOd196[0x3],OxOd196[0x4],OxOd196[0x5],OxOd196[0x6],OxOd196[0x7],OxOd196[0x8],OxOd196[0x9],OxOd196[0xa],OxOd196[0xb],OxOd196[0xc],OxOd196[0xd],OxOd196[0xe],OxOd196[0xf],OxOd196[0x10],OxOd196[0x11],OxOd196[0x12],OxOd196[0x13],OxOd196[0x14],OxOd196[0x15],OxOd196[0x16],OxOd196[0x17],OxOd196[0x18],OxOd196[0x19],OxOd196[0x1a],OxOd196[0x1b],OxOd196[0x1c],OxOd196[0x1d],OxOd196[0x1e],OxOd196[0x1f],OxOd196[0x20],OxOd196[0x21],OxOd196[0x22],OxOd196[0x23],OxOd196[0x24],OxOd196[0x25],OxOd196[0x26],OxOd196[0x27],OxOd196[0x28],OxOd196[0x29],OxOd196[0x2a],OxOd196[0x2b],OxOd196[0x2c],OxOd196[0x2d],OxOd196[0x2e],OxOd196[0x2f],OxOd196[0x30],OxOd196[0x31],OxOd196[0x32],OxOd196[0x33],OxOd196[0x34],OxOd196[0x35],OxOd196[0x36],OxOd196[0x37],OxOd196[0x38],OxOd196[0x39],OxOd196[0x3a],OxOd196[0x3b],OxOd196[0x3c],OxOd196[0x3d],OxOd196[0x3e],OxOd196[0x3f],OxOd196[0x40],OxOd196[0x41],OxOd196[0x42],OxOd196[0x43],OxOd196[0x44],OxOd196[0x45],OxOd196[0x46],OxOd196[0x47],OxOd196[0x48],OxOd196[0x49],OxOd196[0x4a],OxOd196[0x4b],OxOd196[0x4c],OxOd196[0x4d],OxOd196[0x4e],OxOd196[0x4f],OxOd196[0x50],OxOd196[0x51],OxOd196[0x52],OxOd196[0x53],OxOd196[0x54],OxOd196[0x54],OxOd196[0x55],OxOd196[0x56],OxOd196[0x57],OxOd196[0x58],OxOd196[0x59],OxOd196[0x5a],OxOd196[0x5b],OxOd196[0x5c],OxOd196[0x5d],OxOd196[0x5e],OxOd196[0x5f],OxOd196[0x60],OxOd196[0x61],OxOd196[0x61],OxOd196[0x62],OxOd196[0x63],OxOd196[0x64],OxOd196[0x65],OxOd196[0x66],OxOd196[0x67],OxOd196[0x68],OxOd196[0x69],OxOd196[0x6a],OxOd196[0x6b],OxOd196[0x6c],OxOd196[0x5a],OxOd196[0x6d],OxOd196[0x6e],OxOd196[0x6f],OxOd196[0x70],OxOd196[0x71],OxOd196[0x72],OxOd196[0x73],OxOd196[0x74],OxOd196[0x75],OxOd196[0x76],OxOd196[0x77],OxOd196[0x78]];for(var i=0x0;i<arr[OxOd196[0x79]];i+=tds){ document.write(OxOd196[0x7a]) ;for(var j=i;j<i+tds&&j<arr[OxOd196[0x79]];j++){var n=arr[j]; document.write(OxOd196[0x7d]+n+OxOd196[0x7e]+n.replace(OxOd196[0x7b],OxOd196[0x7c])+OxOd196[0x7f]) ; document.write(n) ; document.write(OxOd196[0x80]) ;} ; document.write(OxOd196[0x81]) ;} ;
								</script>
							</table>
							<br />	
						</fieldset>
					</td>
				</tr>
				<tr>
					<td align="right">
					    <input type="button" value="<%= GetString("Cancel") %>" onclick="cancel()" />
					</td>
				</tr>
			</table>
	</body>
</html>
