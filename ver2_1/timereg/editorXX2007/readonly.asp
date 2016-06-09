<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<html>	
    <head>
		<link rel="stylesheet" href="asp.css"  type="text/css" />
	</head>
    <body>
		
		<form name="theForm" action="Get_HtmlContent.asp" method="post" ID="Form1">
					
		<h1>Read Only Support</h1>
		<hr>
		<p>This example show you you can set ReadOnly property to true if you would like to display read-only content in the Cute Editor in some situations. </p>
		<br />
		<%
			dim content
			content = "Type here"		
					
			Dim editor
			Set editor = New CuteEditor
			editor.ID = "Editor1"
			editor.Text = content
			editor.ReadOnly=true
			editor.FilesPath = "CuteEditor_Files"	
			editor.EditorBodyStyle = "font:normal 12px arial;"
			editor.EditorWysiwygModeCss = "asp.css"
			editor.Draw()
												
			' Request.Form(ID) access from other page
		%>
						
		</form>
	</body>
</html>