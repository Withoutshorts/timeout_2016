<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<html>	
    <head>
		<link rel="stylesheet" href="asp.css"  type="text/css" />
		<%
			dim result
			result = ""
			
			If request.QueryString("postback") = "true" then
				result = Request.Form("Editor1")
			End if
		%>
	</head>
    <body>
		
		<form name="theForm" action="xmloutput.asp?postback=true" method="post" ID="Form1">
					
		<h1>Support output well-formed HTML and XHTML</h1>
		<hr>
		<div class="infoBox">This example shows you CuteEditor supports output well-formed XHTML. Your choice of XHTML 1.0 or HTML 4.01 output. </div>
		<br />
		<%
			dim content
			content = "Type here"				
				
			Dim editor
			Set editor = New CuteEditor
			editor.ID = "Editor1"
			editor.Text = content
			editor.XHTMLOutput=true
			editor.UseRelativeLinks=false 
			editor.FilesPath = "CuteEditor_Files"
			editor.EditorBodyStyle = "font:normal 12px arial;"
			editor.EditorWysiwygModeCss = "asp.css"
			editor.Draw()
												
			' Request.Form(ID) access from other page
		%>
		
		<br/><br/>
				
		<textarea name="textbox1" id="Textarea1" style="font-family:Arial;height:250px;width:730px;"><%= Server.HTMLEncode(result) %></textarea>		
				
						
		</form>
	</body>
</html>