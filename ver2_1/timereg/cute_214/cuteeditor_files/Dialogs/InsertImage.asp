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
	    <title><%= GetString("InsertImage") %></title>
		<meta http-equiv="EXPIRES" content="0" />
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<!--[if IE]>
			<link href="../Style/IE.css" type="text/css" rel="stylesheet" />
		<![endif]-->
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>
        <style type="text/css">
	    #upload_image {VISIBILITY: inherit; Z-INDEX: 2;width:350px;height:65px;margin:10px 0 0 10px}
	    #browse_Frame{border:1px solid #a0a0a0;width:270px;height:250px}
	    #divouter{border:1px solid #a0a0a0; vertical-align: top; overflow: auto; width: 310px;height: 250px; background-color: white; text-align: center}
	    #divpreview{background-color: white; height: 100%; width: 100%}
	    #bordercolor_Preview{margin-bottom:-3px}
		.image-align{vertical-align:middle;margin:3px;}
		html, body,#ajaxdiv,#Form1 {height: 100%;}
		</style>
	<%
		Response.Expires = -1
		
		Dim Current_ImageGalleryPath
		Current_ImageGalleryPath=ImageGalleryPath
		if Request.QueryString("MP") <> "" then
			Current_ImageGalleryPath = Request.QueryString("MP")
		End if
	%>
	</head>
	<body>	    
		<div id="ajaxdiv">
		    <% 
                dim fs1,fullfolderpath
                fullfolderpath=Server.MapPath(Current_ImageGalleryPath)
                set fs1=Server.CreateObject("Scripting.FileSystemObject")
                if fs1.FolderExists(fullfolderpath)<>true then
                    Response.write("<div style='margin:5px'>Folder "&Current_ImageGalleryPath&" does not exist! Please use ImageGalleryPath property to specify the Image gallery path!</div>")
                    Response.End
                end if
                set fs1=nothing
            %>	
			<table border="0" cellspacing="2" cellpadding="0" width="100%">
	            <tr>
		            <td>
		            </td>
		            <td valign="bottom" style="width:200px">
<img src="../Images/newfolder.gif" id="btn_CreateDir" onclick="CreateDir_click();" title="<%= GetString("Createdirectory") %>" <% if CBool(AllowCreateFolder) then %> class="cursor dialogButton" <% else %> class="CuteEditorButtonDisabled" <% end if%> onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
<img class="cursor dialogButton" src="../Images/zoom_in.gif" id="btn_zoom_in" onclick="Zoom_In();" title="<%= GetString("ZoomIn") %>" onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
<img class="cursor dialogButton" src="../Images/zoom_out.gif" id="btn_zoom_out" onclick="Zoom_Out();" title="<%= GetString("ZoomOut") %>" onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
<img class="cursor dialogButton" src="../Images/bestfit.gif" id="btn_bestfit" onclick="BestFit();" title="<%= GetString("BestFit") %>" onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
<img class="cursor dialogButton" src="../Images/Actualsize.gif" id="btn_Actualsize" onclick="Actualsize();" title="<%= GetString("ActualSize") %>" onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
		            </td>	
	            </tr>
				<tr>
					<td style="width:260px;height:240px;">
<iframe src="browse_Img.asp?<%=setting %>&MP=<%=Current_ImageGalleryPath%>&Theme=<%=Theme%>" id="browse_Frame" frameborder="0" scrolling="auto"></iframe>		
					</td>
					<td style="width:326px">
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
							<legend><%= GetString("Layout") %></legend>
							<table border="0" cellpadding="2" cellspacing="0" width="100%">
								<tr>
									<td style="width:72px;white-space:nowrap"><%= GetString("Alignment") %>:</td>
									<td style="text-align:left">
<select name="ImgAlign" style="width:80px;" id="Align" onchange="do_preview()" onpropertychange="do_preview()">
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
									<td><%= GetString("Bordersize") %>:</td>
									<td style="text-align:left">
<input type="text" size="2" name="Border" onchange="do_preview()" onpropertychange="do_preview()" onkeypress="return CancelEventIfNotDigit()" style="width:80px;" id="Border" />
									</td>
								</tr>
								<tr>
									<td><%= GetString("BorderColor") %>:</td>
									<td style="text-align:left">
<input type="text" id="bordercolor" name="bordercolor" size="7" style="width:57px;" />
<img title="" src="../Images/colorpicker.gif" id="bordercolor_Preview" class="cursor" />
									</td>
								</tr>
								<tr>
									<td colspan="2">
										<table border="0" cellpadding="0" cellspacing="0" class="normal">
											<tr>
												<td style="width:76px; white-space:nowrap" ><%= GetString("Width") %>:</td>
												<td>
<input type="text" size="2" id="inp_width" onkeyup="checkConstrains('width');" onkeypress="return CancelEventIfNotDigit()" style="width:80px" />
												</td>
												<td rowspan="2" align="right" valign="middle">
												    <img src="../Images/locked.gif" id="imgLock" width="25" height="32"	title="<%= GetString("ConstrainProportions") %>" />
												</td>
											</tr>
											<tr>
												<td><%= GetString("Height") %>:</td>
												<td>
<input type="text" size="2" id="inp_height" onkeyup="checkConstrains('height');" onkeypress="return CancelEventIfNotDigit()" style="width:80px;" />
												</td>
											</tr>
											<tr>
												<td colspan="2">
<input type="checkbox" id="constrain_prop" checked="checked" onclick="javascript:toggleConstrains();" />
												    <%= GetString("ConstrainProportions") %>													
												</td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
						</fieldset>
						<fieldset>
							<legend><%= GetString("Spacing") %></legend>
							<table border="0" cellpadding="4" cellspacing="0" width="100%">
								<tr>
									<td>
										<table border="0" cellpadding="1" cellspacing="0" class="normal" width="100%">
											<tr>
												<td valign="middle" style="width:100px; white-space:nowrap"><%= GetString("Horizontal") %>:</td>
												<td>
<input type="text" size="2" name="HSpace" onchange="do_preview()" onpropertychange="do_preview()" onkeypress="return CancelEventIfNotDigit()" style="width:80px" id="HSpace" />
												</td>
											</tr>
											<tr>
												<td valign="middle"><%= GetString("Vertical") %>:</td>
												<td>
<input type="text" size="2" name="VSpace" onchange="do_preview()" onpropertychange="do_preview()" onkeypress="return CancelEventIfNotDigit()" style="width:80px" id="VSpace" /></td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
						</fieldset>
					</td>
					<td style="width:2px"></td>
					<td valign="top">
						<fieldset>
							<legend><%= GetString("Insert") %></legend>
							<table border="0" cellpadding="4" cellspacing="0">
								<tr>
									<td>
										<table border="0" cellpadding="1" cellspacing="0" class="normal">
											<tr>
												<td valign="middle">
													<%= GetString("Url") %>:</td>
												<td colspan="3">
													<input type="text" id="TargetUrl" onchange="do_preview()" onpropertychange="do_preview()"
														size="43" name="TargetUrl" /></td>
												<td></td>
											</tr>
											<tr>
												<td valign="middle"><%= GetString("Alternate") %>:</td>
												<td valign="middle"><input type="text" id="AlternateText" size="22" name="AlternateText" /></td>
												<td valign="middle" style="white-space:nowrap" >&nbsp;<%= GetString("ID") %>:</td>
												<td><input type="text" id="inp_id" size="12" /></td>
												<td></td>
											</tr>
											<tr>
												<td valign="middle" style="white-space:nowrap" ><%= GetString("longDesc") %>:</td>
												<td valign="middle" colspan="3"><input type="text" id="longDesc" size="43" name="longDesc" /></td>
												<td><img alt="" src="../Images/Accessibility.gif" /></td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
						</fieldset>
						<%
						    Dim Style_Display_None
							if Not CBool(AllowUpload) then
							    Style_Display_None="Style='Display:none'"
						    end if
						%>
						<fieldset id="fieldsetUpload" <%= Style_Display_None %> >
							<legend><%= GetString("Upload") %> (<%= GetString("MaxFileSizeAllowed") %> <%=MaxImageSize%>K)</legend>
<iframe src="upload.asp?<%=setting %>&Theme=<%=Theme%>&FP=<%=Current_ImageGalleryPath%>&Type=Image" id="upload_image" frameborder="0" scrolling="auto"></iframe>
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
	<script type="text/javascript" src="../Scripts/Dialog/Dialog_InsertImage.js"></script>
		<script type="text/javascript">	
	    var OK = "<%= GetString("OK")%>";
	    var Cancel = "<%= GetString("Cancel")%>";
	    var InputRequired = "<%= GetString("InputRequired")%>";
	    var ValidID = "<%= GetString("ValidID")%>";
	    var ValidColor = "<%= GetString("ValidColor")%>";
	    var SelectImagetoInsert = "<%= GetString("SelectImagetoInsert")%>";
	    
	    var currentfolder = "browse_Img_gallery.asp?<%=setting %>&MP=<%=Current_ImageGalleryPath %>&Theme=<%=Theme%>";	
	    
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
		    browse_Frame.location="browse_Img.asp?<%=setting %>&Theme=<%=Theme%>&MP=<%=ImageGalleryPath%>&loc="+path+"";
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
		    upload_image.location="upload.asp?<%=setting %>&FP="+path+"&Theme=<%=Theme%>&Type=Image";
	    }
	    function ResetImage()
	    {	
		    document.getElementById("img_demo").src="../Images/1x1.gif";
	    }
	</script>
	</body>
</html>