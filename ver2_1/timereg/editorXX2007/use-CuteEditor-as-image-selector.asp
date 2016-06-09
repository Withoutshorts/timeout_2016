<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<html>	
    <head>
	</head>
	<body>
		    <h1>Use CuteEditor as an image selector</h1>
		    <hr />
		    <div class="infoBox">This example demonstrates how to use CuteEditor as an image selector.</div>
		    <br />
            <input name="imageFld" type="text" id="imageFld" style="width:300px;" /> 
            <input name="Change" id="Change" type="button" value="Change Image" onclick="callInsertImage()" />   
			    <br />
			    <%
				    Dim editor
				    Set editor = New CuteEditor
    				
				    editor.ID = "Editor1"
				    editor.Text = ""
                    editor.Width="400"
                    editor.Height="300" 
                    editor.AutoConfigure="None"
                    editor.ThemeType="OfficeXp"
                    editor.ShowBottomBar=false
				    editor.Draw()
    																					
				    ' Request.Form(ID) access from other page
			    %>
    				
			    <br/><br/>					

               <Script Language="javascript"> 
                function callInsertImage()  
                {  
                        var editor1 = document.getElementById('<%=editor.ClientID%>');
                        editor1.FocusDocument();  
                        var editdoc = editor1.GetDocument();  
                        editor1.ExecCommand('new');
                        editor1.ExecCommand('ImageGalleryByBrowsing');
                        InputURL();
                        document.getElementById("imageFld").focus(); 
                }    
                
                function InputURL()
                { 
                    var editor1 = document.getElementById('<%=editor.ClientID%>');
                    var editdoc = editor1.GetDocument();  
                    var imgs = editdoc.getElementsByTagName("img");
                    if(imgs.length>0)  
                    {	document.getElementById("imageFld").value = imgs[imgs.length-1].src;
                    }  
                    else
                    {
                        setTimeout(InputURL,500); 
                    }  
                }       
            </script> 			
				 
	</body>
</html>