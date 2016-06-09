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
			Dim EditorOnPaste
			EditorOnPaste = "Disabled"
			
			if Request.QueryString("EditorOnPaste") <> "" then
				EditorOnPaste = Request.QueryString("EditorOnPaste")
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
		
		<form name="theForm" action="xmloutput.asp?postback=true" method="post" ID="Form1">
					
		<h1>Handle pasted text</h1>
		<hr>
		<div class="infoBox">With Cute Editor, you can specify if formatting is stripped when text is pasted into the editor. </div>
		
		<table id="RadioList" border="0">
			<tr>
				<td>
					<input type="radio" name="RadioList" <%= Checked("Disabled", EditorOnPaste) %> value="EditorOnPaste.asp?EditorOnPaste=Disabled" onclick="jumpTo(this)" language="javascript" ID="Radio1"/>Disabled </td>
				<td>
					<input type="radio" name="RadioList" <%= Checked("PasteText", EditorOnPaste) %> value="EditorOnPaste.asp?EditorOnPaste=PasteText" onclick="jumpTo(this)" language="javascript" ID="Radio2"/>PasteText</td>
				<td>
					<input type="radio" name="RadioList" <%= Checked("PastePureText", EditorOnPaste) %> value="EditorOnPaste.asp?EditorOnPaste=PastePureText" onclick="jumpTo(this)" language="javascript" ID="Radio5"/>PastePureText</td>
				<td>
					<input type="radio" name="RadioList" <%= Checked("PasteWord", EditorOnPaste) %>value="EditorOnPaste.asp?EditorOnPaste=PasteWord" onclick="jumpTo(this)" language="javascript" ID="Radio3"/>PasteWord</td>
				<td>
					<input type="radio" name="RadioList" <%= Checked("ConfirmWord", EditorOnPaste) %>value="EditorOnPaste.asp?EditorOnPaste=ConfirmWord" onclick="jumpTo(this)" language="javascript" ID="Radio4"/>ConfirmWord</td>
			</tr>
		</table>
		<br />
		<%
			dim content
			content = "Type here"				
				
			Dim editor
			Set editor = New CuteEditor
			editor.ID = "Editor1"
			editor.Text = content
			editor.FilesPath = "CuteEditor_Files"
			editor.EditorBodyStyle = "font:normal 12px arial;"
			editor.EditorWysiwygModeCss = "asp.css"
			editor.EditorOnPaste = EditorOnPaste
			editor.Draw()
												
			' Request.Form(ID) access from other page
		%>
		
		<br/><br/>
				
						
		</form>
	</body>
</html>
