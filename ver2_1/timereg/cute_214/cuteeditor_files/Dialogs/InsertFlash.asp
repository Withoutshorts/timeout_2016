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
	    <title><%= GetString("InsertFlash") %></title>
		<meta http-equiv="EXPIRES" content="0" />
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<!--[if IE]>
			<link href="../Style/IE.css" type="text/css" rel="stylesheet" />
		<![endif]-->
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>

        <style type="text/css">
	        #upload_image {VISIBILITY: inherit; Z-INDEX: 2;width:300px;height:65px;margin:10px 0 0 10px}
			#browse_Frame{border:1px solid #a0a0a0;width:270px;height:250px}
	        #divouter{border:1px solid #a0a0a0; vertical-align: top; overflow: auto; width: 290px;height: 250px; background-color: white; text-align: center}            #divpreview{background-color:white;height:100%;width:100%;}
		    html, body,#ajaxdiv,#Form1 {height: 100%;}
		</style>		
	    <%
		    Response.Expires = -1
    				
		    Dim Current_FlashGalleryPath
		    Current_FlashGalleryPath=FlashGalleryPath
		    if Request.QueryString("MP") <> "" then
			    Current_FlashGalleryPath = Request.QueryString("MP")
		    End if
	    %>	
	</head>
	<body>
		<div id="ajaxdiv">
		    <% 
                dim fs1,fullfolderpath
                fullfolderpath=Server.MapPath(Current_FlashGalleryPath)
                set fs1=Server.CreateObject("Scripting.FileSystemObject")
                if fs1.FolderExists(fullfolderpath)<>true then
                    Response.write("<div style='margin:5px'>Folder "&Current_FlashGalleryPath&" does not exist! Please use FlashGalleryPath property to specify the Flash gallery path!</div>")
                    Response.End
                end if
                set fs1=nothing
            %>	
			<table border="0" cellspacing="2" cellpadding="0" width="100%">
	            <tr>
		            <td style="white-space:nowrap; width:270px">
		            </td>
		            <td style="width:200px">
<img src="../Images/newfolder.gif" id="btn_CreateDir" onclick="CreateDir_click();" title="<%= GetString("Createdirectory") %>" <% if CBool(AllowCreateFolder) then %> class="cursor dialogButton" <% else %> class="CuteEditorButtonDisabled" <% end if%> onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
<img src="../Images/zoom_in.gif" id="btn_zoom_in" onclick="Zoom_In();" title="<%= GetString("ZoomIn") %>" class="cursor dialogButton" onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
<img src="../Images/zoom_out.gif" id="btn_zoom_out" onclick="Zoom_Out();" title="<%= GetString("ZoomOut") %>" class="cursor dialogButton" onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
<img src="../Images/Actualsize.gif" id="btn_Actualsize" onclick="Actualsize();" title="<%= GetString("ActualSize") %>" class="cursor dialogButton" onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
		            </td>
	            </tr>
				<tr>
					<td style="width:270px;">
<iframe src="browse_flash.asp?<%=setting %>&MP=<%=Current_FlashGalleryPath%>&Theme=<%=Theme%>" id="browse_Frame" frameborder="0" scrolling="auto"></iframe>		
					</td>
					<td style="width:290px;padding:0">
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
					<td style="vertical-align:top">
						<fieldset><legend><%= GetString("Properties") %></legend>
						<div style="padding:3px;">
										<table border="0" cellpadding="1" cellspacing="0" id="Table3">
											<tr>
												<td style="width:110px;"><%= GetString("Width") %>:</td>
												<td style="width:120px;">
<input type="text" name="Width" id="Width" style="width:80px;" onchange="do_preview()"	onkeypress="return CancelEventIfNotDigit()" value="200" />
												</td>
											</tr>
											<tr>
												<td><%= GetString("Height") %>:</td>
												<td>
 <input type="text" name="Height" id="Height" style="width:80px;" onchange="do_preview()" onkeypress="return CancelEventIfNotDigit()" value="200" />
											    </td>
											</tr>
								            <tr>
									            <td><%= GetString("Backgroundcolor") %>:</td>
									            <td style="text-align:left">
<input type="text" id="bgColortext" name="bgColortext" size="7" style="width:57px;vertical-align:top;" />
<img title="" src="../Images/colorpicker.gif" id="bgColortext_Preview" class="cursor" style="margin-bottom:-2px" />
									            </td>
								            </tr>
											<tr>
												<td><%= GetString("Quality") %>:
												</td>
												<td>
<select name="Quality" id="Quality" style="width:80px;" onchange="do_preview()" onpropertychange="do_preview()" >
    <option selected="selected" value="high">High</option>
    <option value="medium">Medium</option>
    <option value="low">Low</option>
</select>
												</td>
											</tr>
											<tr>
												<td style="width:100px"><%= GetString("Scale") %>:</td>
												<td>
<select name="Scale" style="width:80px;" id="Scale" onchange="do_preview()" onpropertychange="do_preview()">
    <option selected="selected" value=""><%= GetString("Default") %></option>
    <option value="Noborder"><%= GetString("Noborder") %></option>
    <option value="Exactfit"><%= GetString("Exactfit") %></option>
</select>
												</td>
											</tr>
										</table>
								</div>
						</fieldset>
						<fieldset>
						<legend><%= GetString("Layout") %></legend>
							<div style="padding:3px;">
										<table border="0" cellpadding="1" cellspacing="0" class="normal" width="100%">
											<tr>
												<td style="width:100px"><%= GetString("Alignment") %>:</td>
												<td>
<select name="Align" style="width:80px;" id="Align" onchange="do_preview()" onpropertychange="do_preview()">
	<option id="optNotSet" selected="selected" value=""><%= GetString("NotSet") %></option>
	<option id="optLeft" value="left"><%= GetString("Left") %></option>
	<option id="optRight" value="right"><%= GetString("Right") %></option>
	<option id="optTexttop" value="textTop"><%= GetString("Texttop") %></option>
	<option id="optAbsMiddle" value="absMiddle"><%= GetString("Absmiddle") %></option>
	<option id="optBaseline" value="baseline"><%= GetString("Baseline") %></option>
	<option id="optAbsBottom" value="absBottom"><%= GetString("Absbottom") %></option>
	<option id="optBottom" value="bottom"><%= GetString("Bottom") %></option>
	<option id="optMiddle" value="middle"><%= GetString("Middle") %></option>
	<option id="optTop" value="top"><%= GetString("Top") %></option>
</select>
												</td>
											</tr>
											<tr>
												<td valign="middle" style="width:110px"><%= GetString("Horizontal") %>:</td>
												<td>
<input type="text" size="2" name="HSpace" onchange="do_preview()" onpropertychange="do_preview()" onkeypress="return CancelEventIfNotDigit()" style="width:80px" id="HSpace" />
												</td>
											</tr>
											<tr>
												<td valign="middle"><%= GetString("Vertical") %>:</td>
												<td>
<input type="text" size="2" name="VSpace" onchange="do_preview()" onpropertychange="do_preview()" onkeypress="return CancelEventIfNotDigit()" style="width:80px" id="VSpace" />
												</td>
											</tr>
										</table>
									</div>
						</fieldset>
					</td>
					<td style="width:2px">
					</td>
					<td style="vertical-align:top">
						<fieldset style="margin-bottom:5px">
							<legend>
								<%= GetString("Insert") %></legend>
							<div style="padding:3px;">
										<table border="0" cellpadding="1" cellspacing="0" class="normal">
											<tr>
												<td valign="middle" style="width:100px">
													<%= GetString("Url") %>:</td>
												<td>
<input type="text" id="TargetUrl" size="43" name="TargetUrl" /></td>
											</tr>
											<tr>
												<td colspan="2">
<input type="checkbox" checked="checked" id="chk_Loop" onchange="do_preview()" onpropertychange="do_preview()" />&nbsp;<%= GetString("Loop") %>&nbsp;&nbsp;
<input type="checkbox" checked="checked" id="chk_Autoplay" onchange="do_preview()" onpropertychange="do_preview()" />&nbsp;<%= GetString("Autoplay") %>&nbsp;&nbsp;
<input type="checkbox" checked="checked" id="chk_Transparency" onchange="do_preview()" onpropertychange="do_preview()" />&nbsp;<%= GetString("Transparency") %><br />
<input type="checkbox" checked="checked" id="chk_FlashMenu" onchange="do_preview()" onpropertychange="do_preview()" />&nbsp;<%= GetString("FlashMenu") %>&nbsp;
<input type="checkbox" checked="checked" id="chk_Fullscreen" onchange="do_preview()" onpropertychange="do_preview()" />&nbsp;<%= GetString("FlashFullscreen") %>&nbsp;
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
							<legend><%= GetString("Upload") %> (<%= GetString("MaxFileSizeAllowed") %> <%=MaxFlashSize%>K)</legend>
<iframe src="upload.asp?<%=setting %>&FP=<%=Current_FlashGalleryPath%>&Type=Flash&Theme=<%=Theme%>" id="upload_image" frameborder="0" scrolling="auto"></iframe>
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
	    <script type="text/javascript" src="../Scripts/Dialog/Dialog_InsertFlash.js"></script>
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
		        row_click(sFileName);
	        }
        	
	        function Refresh(path)
	        {
		        browse_Frame.location="browse_flash.asp?<%=setting %>&Theme=<%=Theme%>&MP=<%=FlashGalleryPath%>&loc="+path+"";
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
		        upload_image.location="upload.asp?<%=setting %>&FP="+path+"&Theme=<%=Theme%>&Type=Flash";
	        }
	        function ResetImage()
	        {	
		        document.getElementById("img_demo").src="../Images/1x1.gif";
	        }
	    </script>	
	</body>	
</html>