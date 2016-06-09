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
				Dim RadioList_Value
				
				RadioList_Value = "en-en"
				
				If Request.QueryString("language") <> "" then
				
					RadioList_Value = Request.QueryString("language")
				
				End If
			%>
	</head>
    <body>
		
		<form name="theForm" action="Get_HtmlContent.asp" method="post" ID="Form2">
					
		<h1>Built-in support for Dutch, Spanish, German ... </h1>
		<hr>
		<div class="infobox">
			This example shows you CuteEditor for ASP has built in support for multiple languages.  
		</div>

		<table id="RadioList" border="0">
			<tr>
				<td nowrap>
					<input type="radio" name="RadioList" <%= Checked("en-en", RadioList_Value) %> value="localization.asp?language=en-en" onclick="jumpTo(this)" language="javascript" ID="Radio1"/><SPAN style="WIDTH: 80px">English 
						(en)</SPAN></SPAN>
				</td>
				<td nowrap>
					<input type="radio" name="RadioList" <%= Checked("fr-fr", RadioList_Value) %> value="localization.asp?language=fr-fr" onclick="jumpTo(this)" language="javascript" ID="Radio2"/><SPAN style="WIDTH: 80px">French 
					(fr)</SPAN></SPAN>
				</td>
				<td nowrap>
					<input type="radio" name="RadioList" <%= Checked("de-de", RadioList_Value) %> value="localization.asp?language=de-de" onclick="jumpTo(this)" language="javascript" ID="Radio3"/><SPAN style="WIDTH: 80px">German 
					(de)</SPAN></SPAN>
				</td>
				<td nowrap>
					<input type="radio" name="RadioList" <%= Checked("es-ES", RadioList_Value) %> value="localization.asp?language=es-ES" onclick="jumpTo(this)" language="javascript" ID="Radio4"/><SPAN style="WIDTH: 80px">Spanish 
					(es)</SPAN></SPAN>
				</td>
				<td nowrap>
					<input type="radio" name="RadioList" <%= Checked("it-IT", RadioList_Value) %> value="localization.asp?language=it-IT" onclick="jumpTo(this)" language="javascript" ID="Radio5"/><SPAN style="WIDTH: 80px">Italian 
					(it)</SPAN></SPAN>
				</td>
			</tr>
			<tr>
				<td nowrap>
					<input type="radio" name="RadioList" <%= Checked("ja-JP", RadioList_Value) %> value="localization.asp?language=ja-JP" onclick="jumpTo(this)" language="javascript" ID="Radio6"/><SPAN style="WIDTH: 80px">Japanese 
					(jp)</SPAN></SPAN>
				</td>
				<td nowrap>
					<input type="radio" name="RadioList" <%= Checked("nl-NL", RadioList_Value) %> value="localization.asp?language=nl-NL" onclick="jumpTo(this)" language="javascript" ID="Radio7"/><SPAN style="WIDTH: 80px">Dutch 
					(nl)</SPAN></SPAN>
				</td>
				<td nowrap>
					<input type="radio" name="RadioList" <%= Checked("nb-NO", RadioList_Value) %> value="localization.asp?language=nb-NO" onclick="jumpTo(this)" language="javascript" ID="Radio8"/><SPAN style="WIDTH: 80px">Norwegian 
					(no)</SPAN></SPAN>
				</td>
				<td nowrap>
					<input type="radio" name="RadioList" <%= Checked("ru-RU", RadioList_Value) %> value="localization.asp?language=ru-RU" onclick="jumpTo(this)" language="javascript" ID="Radio9"/><SPAN style="WIDTH: 80px">Russian 
					(ru)</SPAN></SPAN>
				</td>
				<td nowrap>
					<input type="radio" name="RadioList" <%= Checked("zh-cn", RadioList_Value) %> value="localization.asp?language=zh-cn" onclick="jumpTo(this)" language="javascript" ID="Radio10"/><SPAN style="WIDTH: 80px">Chinese 
					(zh-cn)</SPAN>
				</td>
			</tr>
			<tr>
				<td nowrap>
					<input type="radio" name="RadioList" <%= Checked("sv-SE", RadioList_Value) %> value="localization.asp?language=sv-SE" onclick="jumpTo(this)" language="javascript" ID="Radio11"/><SPAN style="WIDTH: 80px">Swedish 
					(sv)</SPAN></SPAN>
				</td>
				<td nowrap>
					<input type="radio" name="RadioList" <%= Checked("pt-BR", RadioList_Value) %> value="localization.asp?language=pt-BR" onclick="jumpTo(this)" language="javascript" ID="Radio12"/><SPAN style="WIDTH: 80px">Portuguese 
					(pt)</SPAN></SPAN>
				</td>
				<td nowrap>
					<input type="radio" name="RadioList" <%= Checked("da-DK", RadioList_Value) %> value="localization.asp?language=da-DK" onclick="jumpTo(this)" language="javascript" ID="Radio13"/><SPAN style="WIDTH: 80px">Danish 
					(dk)</SPAN></SPAN>
				</td>
				<td colspan =2>
				</td>
			</tr>
		</table>
		<br />
		<%
			dim content
			content = "Type here"				
				
			dim contactstring,temp_styledropdownMenuNames,temp_styledropdownMenuList,temp_snippetDropDownMenuNames,temp_snippetDropDownMenuList,temp_imagesDropDownMenuNames,temp_imagesDropDownMenuList
			
					
			Dim editor
			Set editor = New CuteEditor
			editor.ID = "Editor1"
			editor.Text = content
			editor.FilesPath = "CuteEditor_Files"
			editor.EditorBodyStyle = "font:normal 12px arial;"
			
			editor.CustomCulture = RadioList_Value
			editor.EditorWysiwygModeCss = "asp.css"
			editor.Draw()
												
			' Request.Form(ID) access from other page
		%>
						
		</form>
	</body>
</html>



<% 
	Function Checked(value, Checking_String)
		If Instr(1,Checking_String,value) <> 0 then
			Checked = "checked"
		Else
			Checked = ""
		End If 			
	End Function
%>