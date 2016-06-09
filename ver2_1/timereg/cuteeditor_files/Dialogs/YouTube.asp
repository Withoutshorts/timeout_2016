<!-- #include file = "Include_GetString.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>		
	    <title><%= GetString("YouTube") %></title>
		<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.1)" />
		<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.1)" />
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<!--[if IE]>
			<link href="../Style/IE.css" type="text/css" rel="stylesheet" />
		<![endif]-->
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>	
		<style type="text/css">
		#idSource{width: 480px; height: 120px;background-color:#ffffff;padding:0;border:1px solid #a0a0a0;}
		html, body,#ajaxdiv,#Form1 {height: 100%;}
		</style>
	</head>
	<body>
		<div id="ajaxdiv">
			<table border="0" cellspacing="2" cellpadding="2" width="100%">
				<tr>
					<td style="white-space:nowrap;">
						<p><%= GetString("PasteYouTube") %>:</p>
						<textarea name="idSource" id="idSource" rows="9" cols="50" onpaste="setTimeout('do_preview()',100)" onchange="do_preview()"></textarea>
						<input type="hidden" id="TargetUrl" size="50" name="TargetUrl" />
					</td>
				</tr>
			</table>
			<div style="padding-top:10px;text-align:center">
<input class="formbutton" type="button" value="<%= GetString("Insert") %>" onclick="do_insert()" />&nbsp;&nbsp;&nbsp;
<input class="formbutton" type="button" value="<%= GetString("Cancel") %>" onclick="do_Close()" />
			</div>
		</div>
	</body>
	<script type="text/javascript" src="../Scripts/Dialog/DialogFoot.js"></script>
	<script type="text/javascript" src="../Scripts/Dialog/Dialog_YouTube.js"></script>
</html>