<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<html>	
    <head>
		<title></title>
		<link rel="stylesheet" href="asp.css"  type="text/css" />
		<script language="JavaScript" type="text/javascript" >
	    function ShowMyDialog(button)
	    {
			//use CuteEditor_GetEditor(elementinsidetheEditor) to get the cute editor instance
			var editor=CuteEditor_GetEditor(button);
			//show the dialog page , and pass the editor as newwin.dialogArguments
			//(handler,url,args,feature)
		    var newwin=showModalDialog("My_Custom_Text.html?_rand="+new Date().getTime()
				,editor,"dialogWidth:400px;dialogHeight:240px");
	    }
</script>
	</head>
	<body>
		<form name="theForm" action="Get_HtmlContent.asp" method="post" ID="Form2">
		<h1>Add custom dialogs</h1>
		<hr>
		<div class="infobox">This example shows you how easy it can be to add your own dialogs to the CuteEditor with the help of CuteEditor extended functionalities. </div>
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
				editor.CustomAddons = "<img title=""Using oncommand"" class=""CuteEditorButton"" onmouseover=""CuteEditor_ButtonCommandOver(this)"" onmouseout=""CuteEditor_ButtonCommandOut(this)"" onmousedown=""CuteEditor_ButtonCommandDown(this)"" onmouseup=""CuteEditor_ButtonCommandUp(this)"" ondragstart=""CuteEditor_CancelEvent()"" onclick=""ShowMyDialog(this)"" src=""CuteEditor_Files/images/contact.gif"" />"
				
				editor.Draw()
			
				' Request.Form(ID) access from other page
		
		%>
		<br>
		<input type="submit" value="Submit"  ID="Submit1" NAME="Submit1"><br>
		</form>
	</body>
</html>