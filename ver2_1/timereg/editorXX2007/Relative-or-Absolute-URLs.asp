<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<html>	
    <head>
		<title></title>
		<link rel="stylesheet" href="asp.css"  type="text/css" />
		
		<script language="JavaScript">

		function jumpTo(strMenu) {
			if (strMenu.value != "")
				top.location.href = strMenu.value;

			return true;
		}
		</script>
		
		<% 
			Dim UrlsList
			UrlsList = "false"
			
			if Request.QueryString("UrlsList") <> "" then
				UrlsList = Request.QueryString("UrlsList")
			end if
				
			Function Checked(value, Checking_String)
				If Instr(1,Checking_String,value) <> 0 then
					Checked = "checked"
				Else
					Checked = ""
				End If 			
			End Function
	    %>
	</head>
    <body>
	<form name="theForm" action="Get_HtmlContent.asp" method="post">
					
		<h1>Relative or Absolute URLs, which do you prefer?</h1>
		<hr />
		<p>With Cute Editor, you have the choice of using either a <span style="color:#cc0000"><b>relative or absolute URL</b></span>.</p>
		<br />
		<table border="0">
			<tr>
				<td><input type="radio" name="RadioList" <%= Checked("Default", UrlsList) %> value="Relative-or-Absolute-URLs.asp?UrlsList=Default" onclick="jumpTo(this)" language="javascript" ID="Radio3"/>Default</td>
				<td><input type="radio" name="RadioList" <%= Checked("SiteRelative", URLType) %> value="Relative-or-Absolute-URLs.asp?URLType=SiteRelative" onclick="jumpTo(this)" language="javascript" ID="Radio4"/>SiteRelative Urls</td>
				<td><input type="radio" name="RadioList" <%= Checked("Absolute", URLType) %> value="Relative-or-Absolute-URLs.asp?URLType=Absolute" onclick="jumpTo(this)" language="javascript" ID="Radio1"/>Absolute Urls</td>
			</tr>
		</table>
		<br />
		<%
			Dim editor
			Set editor = New CuteEditor
			editor.ID = "Editor1"
			editor.Text = "<div><a href=""some.htm"">This is a relative path</a><br><br><a href=""/some.htm"">This is a Site Root Relative path</a> <br><br><a href=""Http://somesite/some.htm"">This is a absolute path.</a><br><br><a href=""#tips"">This is a link to anchor.</a><br><br></div>"
			editor.FilesPath = "CuteEditor_Files"
			editor.EditorBodyStyle = "font:normal 12px arial;"
			editor.EditorWysiwygModeCss = "asp.css"
			
			If request.QueryString("URLType") <> ""  Then
				editor.URLType = request.QueryString("URLType")
			Else								
				editor.URLType = request.QueryString("UrlsList")
			End If 
			
			editor.Draw()
												
			' Request.Form(ID) access from other page
		%>
						
		</form>
	</body>
</html>