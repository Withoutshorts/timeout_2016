<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<html>	
    <head>
		<link rel="stylesheet" href="asp.css"  type="text/css" />
	</head>
    <body>
		
		<form name="theForm" action="Get_HtmlContent.asp" method="post" ID="Form1">
					
		<h1>Customized Toolbar Example</h1>	
		<hr>						
		<div class="infobox">Developers can personalize the toolbar by playing with Template property (see the example code on side). </div>
		<br />
		<%
			dim content
			content = "<H1><FONT color=#2f4f4f>Customized Toolbar</FONT></H1>"
		
			Dim editor
			Set editor = New CuteEditor
			
			editor.ID = "Editor1"
			editor.Text = content
			editor.FilesPath = "CuteEditor_Files"
			editor.ShowBottomBar = false
			editor.TemplateItemList="G_start,Bold,Italic,Underline,Separator,JustifyLeft,JustifyCenter,JustifyRight,G_end"
			
			editor.Width = 600
			editor.Draw()
			
			' Request.Form(ID) access from other page
		
			%>
			
			<br>
			
			<input type="submit" value="Submit" ID="Submit2" NAME="Submit1"><br>
						
		</form>
	</body>
</html>