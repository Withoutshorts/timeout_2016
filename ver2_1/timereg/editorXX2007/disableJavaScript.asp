<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<html>	
    <head>
		<link rel="stylesheet" href="asp.css"  type="text/css" />
	</head>
    <body>
		
		<form name="theForm" action="Get_HtmlContent.asp" method="post" ID="Form2">
					
		<h1>Disable Javascript Example</h1>
		<hr>
		CuteEditor for asp  supports uplevel and downlevel client browsers. Program interface available to disable rich edit as you wish.
					
		<br />
		<%
			dim content
			content = "Type here"	
			
			Dim editor
			Set editor = New CuteEditor
			editor.ID = "Editor1"
			editor.Text = content
			editor.FilesPath = "CuteEditor_Files"
			editor.EnableClientScript = "false"
			
			editor.EditorBodyStyle = "font:normal 12px arial;"
			editor.EditorWysiwygModeCss = "asp.css"
			editor.Draw()
												
			' Request.Form(ID) access from other page
		%>
						
		</form>
	</body>
</html>