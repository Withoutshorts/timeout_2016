<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<html>	
    <head>
		<title></title>
		<link rel="stylesheet" href="asp.css"  type="text/css" />
	</head>
	<body bottommargin="0" topmargin="0" marginheight="0" marginwidth="0">
		<form name="theForm" action="Get_HtmlContent.asp" method="post">
		<table border="0" cellpadding="0" cellspacing="0" width=100% ID="Table1">
			<tr>			
				<td width=10 nowrap></td>
				<td valign="top" nowrap id="leftcolumn">
					 <!-- #include file = "leftmenu.asp" -->			
				</td>
				<td width=10 nowrap></td>
				<td valign="top">
					<h1>Integration with ieSpell </h1>
					<hr>
					<div class="infoBox">ieSpell is a <b>free</b> Internet Explorer browser extension that spell checks text input boxes on a webpage. It is <I>much</I> faster than a server-side solution. Plus you get to 
store and use your personal word list across all your applications, instead of 
maintaining separate ones on each application.</div>
						<br/>
						For detail information, please check <A href="http://www.ieSpell.com">ieSpell.com</A> Site. 
						<br/><br/>
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
				</td>
				<td width=10></td>
			</tr>
			</table>
		</form>
	</body>
</html>