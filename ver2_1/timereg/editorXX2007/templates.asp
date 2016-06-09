<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<html>	
    <head>
		<link rel="stylesheet" href="asp.css"  type="text/css" />
	</head>
    <body>
		
		<form name="theForm" action="Get_HtmlContent.asp" method="post" ID="Form1">
					
		<h1>Content management with templates </h1>
		<hr>
		<div class="infoBox">The basic idea behind a Content Management System (<B>CMS</B>) is to separate 
the management of content from design. Cute Editor allows the site designer to 
easily create and establish <B>templates</B> to give the site a uniform look. 
Templates may be modified when desired. <BR></div>
		<br />
		<%
			dim content
			content = "Type here"				
				
			Dim editor
			Set editor = New CuteEditor
			editor.ID = "Editor1"
			editor.FilesPath = "CuteEditor_Files"
			editor.LoadHTML("Templates/template3.htm")
			editor.EditorBodyStyle = "font:normal 12px arial;"
			editor.EditorWysiwygModeCss = "asp.css"
			editor.Draw()
												
			' Request.Form(ID) access from other page
		%>
						
		</form>
	</body>
</html>