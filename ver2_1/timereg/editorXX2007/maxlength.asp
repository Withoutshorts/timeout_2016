<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<html>	
    <head>
		<link rel="stylesheet" href="asp.css"  type="text/css" />
	</head>
    <body>
		
		<form name="theForm" action="Get_HtmlContent.asp" method="post" ID="Form1">
					
		<h1>Use MaxHTMLLength to Protect Your Database</h1>
		<hr>
		<div class="infoBox">We often use a database backend such as <A 
href="http://www.microsoft.com/sql/">SQL Server</A> to store data. In such 
databases, you have to specify the length of any textual field when a table is 
designed. If you tried to insert a record whose text length is greater than 
allowed by your table, an error will be reported. <BR><BR>To prevent this type 
of error from occurring, developers can use <STRONG>MaxHTMLLength</STRONG> in the Cute Editor to limit the length of the 
user’s input. <BR><BR>Try click the <B>submit</B> button. <BR></div>
<br>
		<%
			dim content
			content = "Type here"				
			Dim editor
			Set editor = New CuteEditor
			editor.ID = "Editor1"
			editor.Text = content
			editor.MaxHTMLLength = "5"
			editor.FilesPath = "CuteEditor_Files"
			editor.EditorBodyStyle = "font:normal 12px arial;"
			editor.EditorWysiwygModeCss = "asp.css"
			editor.Draw()
												
			' Request.Form(ID) access from other page
		%>
						
		</form>
	</body>
</html>
