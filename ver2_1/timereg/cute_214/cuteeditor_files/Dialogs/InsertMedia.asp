<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>		
		<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.1)" />
		<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.1)" />
		<meta http-equiv="Cache-Control" content="no-cache" />
		<meta http-equiv="Pragma" content="no-cache" />		
	    <title><%= GetString("InsertMedia") %></title>
		<meta http-equiv="EXPIRES" content="0" />
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<!--[if IE]>
			<link href="../Style/IE.css" type="text/css" rel="stylesheet" />
		<![endif]-->
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>
        <style type="text/css">
	    #upload_image {VISIBILITY: inherit; Z-INDEX: 2;width:350px;height:65px;margin:10px 0 0 10px}
	    #browse_Frame{border:1px solid #a0a0a0;width:270px;height:250px}
	    #divouter{border:1px solid #a0a0a0; vertical-align: top; overflow: auto; width: 290px;height: 250px; background-color: white; text-align: center}
	    #divpreview{background-color: white; height: 100%; width: 100%}
		html, body,#ajaxdiv,#Form1 {height: 100%;}
		</style>
	<%
		Response.Expires = -1
		
		Dim Current_MediaGalleryPath
		Current_MediaGalleryPath=MediaGalleryPath
		if Request.QueryString("MP") <> "" then
			Current_MediaGalleryPath = Request.QueryString("MP")
		End if		
	%>	
	</head>
	<body>
		<div id="ajaxdiv">
		    <% 
                dim fs1,fullfolderpath
                fullfolderpath=Server.MapPath(Current_MediaGalleryPath)
                set fs1=Server.CreateObject("Scripting.FileSystemObject")
                if fs1.FolderExists(fullfolderpath)<>true then
                    Response.write("<div style='margin:5px'>Folder "&Current_MediaGalleryPath&" does not exist! Please use MediaGalleryPath property to specify the Media gallery path!</div>")
                    Response.End
                end if
                set fs1=nothing
            %>	
			<table border="0" cellspacing="2" cellpadding="0" width="100%">
	            <tr>
		            <td style="white-space:nowrap; width:250px">
		            </td>
		            <td style="width:200px">
<img src="../Images/newfolder.gif" id="btn_CreateDir" onclick="CreateDir_click();" title="<%= GetString("Createdirectory") %>" <% if CBool(AllowCreateFolder) then %> class="cursor dialogButton" <% else %> class="CuteEditorButtonDisabled" <% end if%> onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
<img src="../Images/zoom_in.gif" id="btn_zoom_in" onclick="Zoom_In();" title="<%= GetString("ZoomIn") %>" class="cursor dialogButton" onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
<img src="../Images/zoom_out.gif" id="btn_zoom_out" onclick="Zoom_Out();" title="<%= GetString("ZoomOut") %>" class="cursor dialogButton" onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
<img src="../Images/Actualsize.gif" id="btn_Actualsize" onclick="Actualsize();" title="<%= GetString("ActualSize") %>" class="cursor dialogButton" onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
		            </td>
	            </tr>
				<tr>
					<td style="width:260px;">
<iframe src="browse_media.asp?<%=setting %>&Theme=<%=Theme%>&MP=<%=Current_MediaGalleryPath%>" id="browse_Frame" frameborder="0" scrolling="auto"></iframe>		
					</td>
					<td style="width:290px">
                        <div id="divouter">
                            <div id="divpreview">
                                <img id="img_demo" alt="" src="../Images/1x1.gif" />
                            </div>
                        </div>
					</td>
				</tr>
			</table>
			<table border="0" cellspacing="0" cellpadding="0" style="margin-top:3px;width:100%;">
				<tr>
					<td valign="top">
						<fieldset>
							<legend><%= GetString("Properties") %></legend>
							<table border="0" cellpadding="4" cellspacing="0" class="normal">
								<tr>
									<td>
										<table border="0" cellpadding="1" cellspacing="0" class="normal">
											<tr>
												<td style="width:100px;">
													<%= GetString("Width") %>:</td>
												<td>
<input type="text" name="Width" id="Width" style="width:80px;" onchange="do_preview()" onpropertychange="do_preview()" value="200" /></td>
											</tr>
											<tr>
												<td>
													<%= GetString("Height") %>:</td>
												<td>
<input type="text" name="Height" id="Height" style="width:80px;" onchange="do_preview()" onpropertychange="do_preview()" value="200" /></td>
											</tr>
											<tr>
												<td>
													<%= GetString("AutoStart") %>:</td>
												<td>
<input type="checkbox" name="Transparency" onchange="do_preview()" onpropertychange="do_preview()" id="AutoStart" checked="checked" value="AutoStart" /></td>
											</tr>
											<tr>
												<td style="height: 24px">
													<%= GetString("ShowControls") %>:</td>
												<td style="height: 24px">
<input type="checkbox" name="Transparency" onchange="do_preview()" onpropertychange="do_preview()" id="ShowControls" checked="checked" value="ShowControls" /></td>
											</tr>
											<tr>
												<td>
													<%= GetString("ShowStatusBar") %>:</td>
												<td>
<input type="checkbox" name="Transparency" onchange="do_preview()" onpropertychange="do_preview()" id="ShowStatusBar" value="ShowStatusBar" /></td>
											</tr>
											<tr>
												<td>WindowlessVideo:</td>
												<td>
<input type="checkbox" name="WindowlessVideo" onchange="do_preview()" onpropertychange="do_preview()" id="WindowlessVideo" value="windowlessVideo" /></td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
						</fieldset>
					</td>
					<td style="width:2px">
					</td>
					<td valign="top">
						<fieldset style="margin-bottom:5px">
							<legend><%= GetString("Insert") %></legend>
							
							<div style="padding:3px;">
										<table border="0" cellpadding="1" cellspacing="0" class="normal" width="100%">
											<tr>
												<td valign="middle">
													<%= GetString("Url") %>:</td>
												<td>
													<input type="text" id="TargetUrl" size="60" name="TargetUrl" />
												</td>
											</tr>
										</table>
									</div>
						</fieldset>
						
						<%
						    Dim Style_Display_None
							if Not CBool(AllowUpload) then
							    Style_Display_None="Style='Display:none'"
						    end if
						%>
						<fieldset id="fieldsetUpload" <%= Style_Display_None %> >
							<legend><%= GetString("Upload") %> (<%= GetString("MaxFileSizeAllowed") %> <%=MaxMediaSize%>K)</legend>
<iframe src="upload.asp?<%=setting %>&FP=<%=Current_MediaGalleryPath%>&Theme=<%=Theme%>&Type=media" id="upload_image" frameborder="0" scrolling="auto"></iframe>
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
	        <script type="text/javascript" src="../Scripts/Dialog/Dialog_InsertMedia.js"></script>
			<script type="text/javascript">	
	            var OK = "<%= GetString("OK")%>";
	            var Cancel = "<%= GetString("Cancel")%>";
	            var InputRequired = "<%= GetString("InputRequired")%>";
	            var ValidID = "<%= GetString("ValidID")%>";
	            var ValidColor = "<%= GetString("ValidColor")%>";
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
		           row_click(sFileName,"");
	            }
    	
	            function Refresh(path)
	            {
		            browse_Frame.location="browse_media.asp?<%=setting %>&Theme=<%=Theme%>&MP=<%=MediaGalleryPath%>&loc="+path+"";
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
		            ResetFields();
		            TargetUrl.value=path;
		            do_preview();
	            }
				function SetUpload_FolderPath(path)
				{	if(path.substring(path.length-1, path.length)=='/')
					{
						path=path.substring(0, path.length-1);
					}
					upload_image.location="upload.asp?<%=setting %>&FP="+path+"&Theme=<%=Theme%>&Type=media";
				}
	            function ResetImage()
	            {	
		            document.getElementById("img_demo").src="../Images/1x1.gif";
	            }
	        </script>
	</body>
</html>
