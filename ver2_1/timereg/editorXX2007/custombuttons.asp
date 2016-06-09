<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<html>	
    <head>
		<title></title>
		<link rel="stylesheet" href="asp.css"  type="text/css" />
	</head>
	<body>
		<form name="theForm" action="Get_HtmlContent.asp" method="post" ID="Form2">
		<h1>Add custom buttons</h1>
		<hr>
		<div class="infobox">This example shows you how easy it can be to add your own functions to the CuteEditor with the help of CuteEditor extended functionalities. </div>
		<br>			
			<%
				dim content
				content = "<H1><FONT color=#2f4f4f>Customized Toolbar</FONT></H1>"
		
				Dim editor
				Set editor = New CuteEditor
			
				editor.ID = "Editor1"
				editor.Text = content
				editor.FilesPath = "CuteEditor_Files"
				editor.ShowBottomBar = false
				editor.TemplateItemList="G_start,Bold,Italic,Underline,Separator,JustifyLeft,JustifyCenter,JustifyRight,Holder,G_end"
				editor.CustomAddons = "<img title=""Using oncommand"" class=""CuteEditorButton"" onmouseover=""CuteEditor_ButtonCommandOver(this)"" onmouseout=""CuteEditor_ButtonCommandOut(this)"" onmousedown=""CuteEditor_ButtonCommandDown(this)"" onmouseup=""CuteEditor_ButtonCommandUp(this)"" ondragstart=""CuteEditor_CancelEvent()"" Command=""MyCmd"" src=""CuteEditor_Files/images/contact.gif"" />"
				
				editor.Draw()
			
				' Request.Form(ID) access from other page
		
		%>
					
					<script language="JavaScript" type="text/javascript" >
		                var editor1=document.getElementById("<%= editor.ClientID %>");
                		
		                function CuteEditor_OnCommand(editor,command,ui,value)
		                {
			                //handle the command by yourself
			                if(command=="MyCmd")
			                {
				             //   editor.ExecCommand("InsertTable");
				                editor1.PasteHTML("Hello World");
				                return true;
			                }
		                }
                    </script>
		<br>
		<input type="submit" value="Submit"  ID="Submit1" NAME="Submit1"><br>
		</form>
	</body>
</html>