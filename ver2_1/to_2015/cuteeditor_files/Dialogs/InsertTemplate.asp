<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>		
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<!--[if IE]>
			<link href="../Style/IE.css" type="text/css" rel="stylesheet" />
		<![endif]-->
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>
		<style type="text/css">
		#framepreview {
			width: 100%;
			height: 100%;
			overflow:auto;
			text-align: left;
			vertical-align: top;
			padding: 0;
			margin: 0;
			background-color: white;
		}
		
	    #upload_image {VISIBILITY: inherit; Z-INDEX: 2;width:500px;height:65px;margin:10px 0 0 10px}
	    #browse_Frame{border:1px solid #a0a0a0;width:270px;height:246px}
	    #divouter{border:1px solid #a0a0a0;padding:0;vertical-align: top; overflow: auto; width:357px;height:250px; background-color:white;}
		#divpreview{background-color: white; height: 100%; width: 100%}
			html, body,#ajaxdiv,#Form1 {height: 100%;}
		</style>
		
	<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.1)" />
	<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.1)" />
	<title><%= GetString("InsertTemplate") %></title>
	<%
		Response.Expires = -1
		
		Dim Current_TemplateGalleryPath
		Current_TemplateGalleryPath=TemplateGalleryPath
		if Request.QueryString("MP") <> "" then
			Current_TemplateGalleryPath = Request.QueryString("MP")
		End if
		
	%>	
	</head>
	<body>
		<div id="ajaxdiv">
		    <% 
                dim fs1,fullfolderpath
                fullfolderpath=Server.MapPath(Current_TemplateGalleryPath)
                set fs1=Server.CreateObject("Scripting.FileSystemObject")
                if fs1.FolderExists(fullfolderpath)<>true then
                    Response.write("<div style='margin:5px'>Folder "&Current_TemplateGalleryPath&" does not exist! Please use TemplateGalleryPath property to specify the Template gallery path!</div>")
                    Response.End
                end if
                set fs1=nothing
            %>	
			<table border="0" cellspacing="0" cellpadding="2" width="100%">
	            <tr>
		            <td style="white-space:nowrap; width:250px">
		            </td>
		            <td valign="bottom" style="width:200px">
<img src="../Images/newfolder.gif" id="btn_CreateDir" onclick="CreateDir_click();" title="<%= GetString("Createdirectory") %>" <% if CBool(AllowCreateFolder) then %> class="cursor dialogButton" <% else %> class="CuteEditorButtonDisabled" <% end if%> onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
<img src="../Images/zoom_in.gif" id="btn_zoom_in" onclick="Zoom_In();" title="<%= GetString("ZoomIn") %>" class="cursor dialogButton" onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
<img src="../Images/zoom_out.gif" id="btn_zoom_out" onclick="Zoom_Out();" title="<%= GetString("ZoomOut") %>" class="cursor dialogButton" onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
<img src="../Images/Actualsize.gif" id="btn_Actualsize" onclick="Actualsize();" title="<%= GetString("ActualSize") %>" class="cursor dialogButton" onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
		            </td>
	            </tr>
				<tr>
					<td style="width:260px;height:240px;">
<iframe src="browse_template.asp?<%=setting %>&Theme=<%=Theme%>&MP=<%=Current_TemplateGalleryPath%>" id="browse_Frame" frameborder="0" scrolling="auto"></iframe>		
					</td>
					<td>
						<div id="divouter">
<iframe id="framepreview" src="../template.asp" scrolling="auto" frameborder="0"></iframe>
						</div>
					</td>
				</tr>
			</table>
			<table border="0" cellspacing="0" cellpadding="0" width="100%" style="margin:2px 0 2px 0">
				<tr>
					<td>
					</td>
					<td style="width:10px">
					</td>
					<td>
					    <input type="hidden" id="hiddenHTML" name="hiddenDirectory" /> 
						<input type="hidden" id="TargetUrl" name="TargetUrl" />						
						<%
						    Dim Style_Display_None
							if Not CBool(AllowUpload) then
							    Style_Display_None="Style='Display:none'"
						    end if
						%>
						<fieldset id="fieldsetUpload" <%= Style_Display_None %> >
							<legend><%= GetString("Upload") %> (<%= GetString("MaxFileSizeAllowed") %> <%=MaxTemplateSize%>K)</legend>
<iframe src="upload.asp?<%=setting %>&Theme=<%=Theme%>&FP=<%=Current_TemplateGalleryPath%>&Type=template" id="upload_image" frameborder="0" scrolling="auto"></iframe>
						</fieldset>
						<div style="padding-top:10px; text-align:center">
<input class="formbutton" type="button" value="   <%= GetString("Insert") %>   " onclick="do_insert()" id="Button1" /> 
&nbsp;&nbsp;&nbsp; 
<input class="formbutton" type="button" value="   <%= GetString("Cancel") %>  " onclick="do_Close()" id="Button2" />
						</div>
					</td>
				</tr>
			</table>
		</div>
		<script type="text/javascript" src="../Scripts/Dialog/DialogFoot.js"></script>
		<script type="text/javascript" src="../Scripts/Dialog/Dialog_InsertTemplate.js"></script>
	    <script type="text/javascript">	
	        var OK = "<%= GetString("OK")%>"
	        var Cancel = "<%= GetString("Cancel")%>"
	        var InputRequired = "<%= GetString("InputRequired")%>"
	        var ValidID = "<%= GetString("ValidID")%>"
	        var ValidColor = "<%= GetString("ValidColor")%>"
	        var SelectImagetoInsert = "<%= GetString("SelectImagetoInsert")%>";
	        
	        function UploadSaved(sFileName,path){
		        ResetFields();
		        try{
		            browse_Frame.location.reload();
		            }
		            catch(x)
		            {}
		        TargetUrl.value = sFileName;
		        setTimeout(function(){do_preview(sFileName);}, 100); 
		        row_click(sFileName);
	        }
        	
	        function Refresh(path)
	        {
		        browse_Frame.location="browse_template.asp?<%=setting %>&Theme=<%=Theme%>&MP=<%=TemplateGalleryPath%>&loc="+path+"";
	        }
	        function CreateDir_click()
	        {
		        <%
			        if not CBool(AllowCreateFolder) then
		        %>
			        alert("<%= GetString("Disabled")%>");
			        return;
		        <%
			        End If
		        %>		    
	            if(Browser_IsIE7())
	            {
		            IEprompt(promptCallback,'<%= GetString("SpecifyNewFolderName")%>', "");		
		            function promptCallback(dir)
		            {
			            var tempPath = browse_Frame.location;	
			            tempPath = tempPath + "&action=createfolder&foldername="+dir;
			            browse_Frame.location = tempPath;		
		            }
	            }
	            else
	            {
		            var dir=prompt("<%= GetString("SpecifyNewFolderName")%>","")
		            if(dir)
		            {
			            var tempPath = browse_Frame.location;	
			            tempPath = tempPath + "&action=createfolder&foldername="+dir;
			            browse_Frame.location = tempPath;			
		            }
	            }
	        }
	        function row_click(path)
	        {
		        framepreview.location=path;
		        setTimeout(do_preview,500);
	        }	    
    		
	        function SetUpload_FolderPath(path)
	        {	if(path.substring(path.length-1, path.length)=='/')
		        {
			        path=path.substring(0, path.length-1);
		        }
		        upload_image.location="upload.asp?<%=setting %>&FP="+path+"&Theme=<%=Theme%>&Type=template";
	        }	
	        function ResetImage()
	        {	
	        }
	    </script>	
	</body>
</html>