<html>	
    <head>
		<title></title>
		<link rel="stylesheet" href="asp.css"  type="text/css" />
	</head>
	<body>
    	<h1>Get the Html Result</h1>
		<hr>
		This example shows how to extract the html content created by the Cute Editor for ASP . 
		<br><br>									
		
		<p>
			<b>Title: </b><br>
			<%=	Request.Form("TitleField")%>
		</p>
		
		<hr width="95%">
		
		<p>
			<b>First Editor content:</b> <br>
			<%=	Request.Form("Editor1")%>
		</p>
		
		<hr width="95%">
		
		<%
			if Request.Form("Editor2") <> "" then
		%>
		<p>
			<b>Second Editor content:</b> <br>
			<%=	Request.Form("Editor2")%>
		</p>
		<% end if %>
	</body>
</html>