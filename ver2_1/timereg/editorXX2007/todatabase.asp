<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<html>	
    <head>
		<title></title>
		<link rel="stylesheet" href="asp.css"  type="text/css" />
	</head>
	<body>
		<form name="theForm" action="todatabase.asp?action=save" method="post" ID="Form2">
		<h1>Add the editor content to the database</h1>
		<hr>
		<div class="infobox">This example shows you the how to feed the editor with content from database and add the editor content to the database.  </div>
		<br>
		<%
			dim conn
			dim connstr
			Set conn = Server.CreateObject("ADODB.Connection")
			connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("test.mdb")
			conn.Open connstr
			
			dim content
			
			if request("action")="save" then
				TitleField		=	Request.Form("TitleField")
				MessageField	=	Request.Form("Editor1")     	
            
				sql="select TitleField,MessageField from test"
				set rs=server.createobject("adodb.recordset")
				rs.open sql,conn,3,3
				
				rs.addnew
				rs("TitleField")	=	TitleField
				rs("MessageField")	=	MessageField
				rs.update
				
				response.write "<font size=5 color=darkred>Message added</font>"
				response.write "<br>"
				rs.close
				set rs = nothing
			else
				MySQL="select top 1 * from test order by id DESC"
				set rs2=conn.execute(MySQL)
				if not rs2.eof then
					TitleField	 = rs2("TitleField")
					MessageField  = rs2("MessageField")
					rs2.close
					set rs2=nothing
				end if
			end if	
			%>
			
			<input name="TitleField" type="text" maxlength="100" size="40" id="TitleField" value="<%= TitleField %>"class="NormalTextBox" style="width:600px;"> <br><br>
						
			<%						
						
				Dim editor
				Set editor = New CuteEditor
				editor.ID = "Editor1"
				editor.Text = MessageField
				editor.FilesPath = "CuteEditor_Files"
				editor.Draw()
			
				' Request.Form(ID) access from other page
		
			%>	
		</form>
	</body>
</html>
