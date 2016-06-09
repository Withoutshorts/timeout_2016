<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<html>	
    <head>
		<link rel="stylesheet" href="asp.css"  type="text/css" />
	</head>
    <body>
		
		<form name="theForm" action="Edithtml.asp?postback=true"  method="post" ID="Form1">
					
		<h1>Edit Static Html</h1>
		<hr>
		<div class="infoBox">This example demonstrates you can use Cute Editor to edit static html page. 
<BR><SPAN class=Normal>Below is an example page that displays a document held in 
an HTML file on the hard drive. When you submit the form, the document is saved 
back to the drive. <A href="document.html"><B>Check the document.html</B></A> <BR></div>
		<br />
		<%
						
			Dim editor
			Set editor = New CuteEditor
			editor.ID = "Editor1"
			editor.EditCompleteDocument = true
			editor.LoadHTML("document.html")
			editor.FilesPath = "CuteEditor_Files"
			editor.EditorBodyStyle = "font:normal 12px arial;"
			editor.UseRelativeLinks=true
			
			If request.QueryString("postback") <> "true" then
				editor.Draw()
			else
				editor.SaveFile("document.html")
				editor.Draw()
			End if 
												
			' Request.Form(ID) access from other page
		%>
						
		</form>
	</body>
</html>
