<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<html>	
    <head>
		<title></title>
		<link rel="stylesheet" href="asp.css"  type="text/css" />
	</head>
    <body>		
		<form name="theForm" action="Get_HtmlContent.asp" method="post" ID="Form1">
		<h1>Multiple Editors Example</h1>
		<hr>
		<div class="infoBox">
			Need to handle mutliple rich-text fields on one page? It's easy with CuteEditor! Check out the sample below to see how 2 instances placed on a single page.
		</div>
		<br>
		<%
			dim content
			content =  "<P><FONT color=#cd0000><b>This paragraph will be edited by the first instance...</b> </FONT></P>"				
		
			Dim editor
			Set editor = New CuteEditor  'Create a new editor class object					
					
			'Set the ID of this editor class
			editor.ID = "Editor1" 
					
			'Set the initial HTML value of editor control
			editor.Text = content 
			editor.FilesPath = "CuteEditor_Files"
			
			'The toolbar items needed to be disabled going to this string. Example DisableItemList="Bold, New, Delete" 
			editor.DisableItemList = "Save"
			
			editor.ThemeType="Office2003"
			editor.EditorWysiwygModeCss = "asp.css"
			editor.Draw()
			
			Response.Write "<br><br>"
			Dim editor2
			Set editor2 = New CuteEditor  'Create a new editor class object					
			
			editor2.ID = "Editor2" 'Set the ID of this editor class
			editor2.Text = "<P><FONT color=#0000cd><b>This paragraph will be edited by the second instance...</b> </FONT></P>" 'Set the initial HTML value of editor control
			editor2.FilesPath = "CuteEditor_Files"
			editor2.EditorWysiwygModeCss = "asp.css"
			editor2.Height = 150
			editor2.AutoConfigure = "Simple"
			editor2.Draw()			
	
			' Retrieve the data from editor: Request.Form("Editor1") 
			%>
			<br>
			<input type="submit" value="Update"  class="button65" NAME="Submit1">
		</form>
	</body>
</html>