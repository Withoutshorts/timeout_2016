<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>XML</title>
</head>

<body>
<%
set objXML = Server.CreateObject("msxml.domdocument")
objXML.load("conf.xml")


Response.write objXML.nodeName


%>


</body>
</html>
