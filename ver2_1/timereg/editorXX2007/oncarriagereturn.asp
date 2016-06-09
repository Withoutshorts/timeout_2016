<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<html>	
    <head>
		<link rel="stylesheet" href="asp.css"  type="text/css" />
		<script language="JavaScript">

		function jumpTo(strMenu) {
			if (strMenu.value != "")
				top.location.href = strMenu.value;

			return true;
		}
		</script>
		
		<% 
			Dim BreakElement
			BreakElement = "div"
			
			if Request.QueryString("BreakElement") <> "" then
				BreakElement = Request.QueryString("BreakElement")
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
		
		<form name="theForm" action="Get_HtmlContent.asp" method="post" ID="Form1">
					
		<h1>&lt;div&gt;,&lt;p&gt; or &lt;br&gt;, which do you prefer?</h1>
		<hr>
		<p>With Cute Editor, you can defines what happens when the "enter" key is pressed in the editor in the <b>BreakElement</b> property.</p>
		<table id="RadioList" border="0">
			<tr>
				<td>
					<input type="radio" name="RadioList" <%= Checked("p", BreakElement) %> value="oncarriagereturn.asp?BreakElement=p" onclick="jumpTo(this)" language="javascript" ID="Radio1"/>P </td>
				<td>
					<input type="radio" name="RadioList" <%= Checked("br", BreakElement) %> value="oncarriagereturn.asp?BreakElement=br" onclick="jumpTo(this)" language="javascript" ID="Radio2"/>BR	</td>
				<td>
					<input type="radio" name="RadioList" <% if BreakElement="div" then response.Write "Checked" End If %> value="oncarriagereturn.asp" onclick="jumpTo(this)" language="javascript" ID="Radio3"/>DIV
				</td>
			</tr>
		</table>
		<br />
		<%			
			Dim editor
			Set editor = New CuteEditor
			editor.ID = "Editor1"
			editor.Text = "Type here"
			editor.FilesPath = "CuteEditor_Files"
			editor.BreakElement = BreakElement	
			editor.EditorBodyStyle = "font:normal 12px arial;"
			editor.EditorWysiwygModeCss = "asp.css"
			editor.Draw()
												
			' Request.Form(ID) access from other page
		%>
						
		</form>
	</body>
</html>