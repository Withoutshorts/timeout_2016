<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<html>	
    <head>
		<link rel="stylesheet" href="asp.css"  type="text/css" />
	</head>
    <body>
		
		<form name="theForm" action="Get_HtmlContent.asp" method="post" ID="Form2">
					
		<h1>Integration with spell checker </h1>
		<hr>
		<P>The NetSpell project is a <B>free</B> spell checking engine written entirely 
in managed C# .net code. Package includes a medium sized English dictionaries. 
Suggestions for misspelled words are generated using phonetic (sounds like) 
matching and ranked by a typographical score (looks like). Also supports "ignore 
all" and "replace all" misspelled-word handling. <BR><BR>For detail information, 
please check <A 
href="http://www.loresoft.com/Default.aspx?tabid=14">LoreSoft.com</A> Site. </P><br />
		<p>Please note that: You need to have the .NET framework installed on your web server to use this feature. </p>
		<br />
		<%
			Dim editor
			Set editor = New CuteEditor
			editor.ID = "Editor1"
			editor.Text = "Type here"
			editor.FilesPath = "CuteEditor_Files"
			editor.EditorBodyStyle = "font:normal 12px arial;"
			editor.EditorWysiwygModeCss = "asp.css"
			editor.Draw()
												
			' Request.Form(ID) access from other page
		%>
			
						
		</form>
	</body>
</html>