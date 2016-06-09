<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<html>	
    <head>
		<title></title>
		<link rel="stylesheet" href="asp.css"  type="text/css" />
	</head>
	<body bottommargin="0" topmargin="0" marginheight="0" marginwidth="0">
		<form name="theForm" action="Get_HtmlContent.asp" method="post" ID="Form2">
		<table border="0" cellpadding="0" cellspacing="0" width=100% ID="Table2">
			<tr>			
				<td width=10 nowrap></td>
				<td valign="top" nowrap id="leftcolumn">
					 <!-- #include file = "leftmenu.asp" -->			
				</td>
				<td width=10 nowrap></td>
				<td valign="top" id="content">
				
					<h1>Simple Configuration</h1>
					<hr>
					<div class="infobox">This example show you the Simple Configuration of editor.</div>
					<br>
					<input name="TitleField" type="text" maxlength="100" size="40" id="TitleField" class="NormalTextBox" style="width:600px;"><br><br>
					<br>
					<%
						dim content
								
						content ="<P><FONT face=""Comic Sans MS"" color=#800000 size=2><STRONG>simple</STRONG></FONT></P>"
						content = content &"<P><STRONG>&nbsp;</STRONG></P>"
						content = content &"<P><STRONG><FONT color=#2f4f4f size=3>simple</FONT></STRONG></P>"
						content = content &"<P><STRONG>&nbsp;</STRONG></P>"
						content = content &"<P><STRONG><FONT face=Verdana color=#ffff00 size=4>simple</FONT></STRONG></P>"
						content = content &"<P><STRONG>&nbsp;</STRONG></P>"
						content = content &"<P><STRONG><FONT face=Tahoma color=#0000ff size=5>simple</FONT></STRONG></P>"
						content = content &"<P><STRONG>&nbsp;</STRONG></P>"
						content = content &"<P><STRONG><FONT face=""Times New Roman"" color=#2e8b57 size=6>simple</FONT></STRONG></P>"
						content = content &"<P><STRONG>&nbsp;</STRONG></P>"
						content = content &"<P><STRONG><FONT face=Tahoma color=#808080 size=7>simple</FONT></STRONG></P>"
						
						Dim editor
						Set editor = New CuteEditor
					
						editor.ID = "Editor1"
						editor.Text = content
						editor.FilesPath = "CuteEditor_Files"
						editor.ImageGalleryPath = "/Uploads"
						editor.MaxImageSize = 50
						editor.AutoConfigure = "Simple"
						'editor.Template= "Bold,Italic,Underline"
						
						'editor.Width = 600
						'editor.Height = 240
						editor.Draw()
			
						' Request.Form(ID) access from other page
		
						%>
						<br>
						<input type="submit" value="Submit"  ID="Submit1" NAME="Submit1">
				</td>
				<td width=80></td>
			</tr>
			</table>
		</form>
	</body>
</html>