<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<% 

dim ps
ps=array("Font","Text","Border","Layout","Background","Other") 

dim activepanel
activepanel=request.QueryString("Style") &""
if activepanel ="" then
	activepanel="Font"
end if

dim activetext
activetext=""
%>

<script type="text/javascript" src="../Scripts/Dialog/Dialog_Tag_Style.js"></script>

<table border="0" cellspacing="0" cellpadding="0" style="margin-top:10px;width:99%">
	<tr>
		<td style="width:94px;vertical-align:top" id="navtd">
			<%
			
			dim px, iscurrent
			iscurrent = false
			For Each px in ps
                dim url				                        
                url=request.ServerVariables("URL")
                				                        
                url=url&"?Tag="&request.QueryString("Tag")     
                url=url&"&Tab=Style"			                        
                url=url&"&Style="&px	                        
                url=url&"&UC="&request.QueryString("UC")	           
                url=url&"&Theme="&request.QueryString("Theme")	
                url=url&"&setting="&request.QueryString("setting")	
			
			    if activepanel = px then
			        activetext=px
                    Response.Write "<div class='buttons'><a class='current' tabindex='-1' href='"&url&"'>"
		            Response.Write "<img alt='' align='absmiddle' src='../Images/style."&px&".gif' style='border:0; vertical-align:inherit;'>"
			        Response.Write px
	                Response.Write "</a><div>"
                else				   
                    Response.Write "<div class='buttons'><a tabindex='-1' href='"&url&"'>"
		            Response.Write "<img alt='' align='absmiddle' src='../Images/style."&px&".gif' style='border:0; vertical-align:inherit;'>"
			        Response.Write px
	                Response.Write "</a><div>"
			    end if 			    
			Next 
			%>
		</td>
		<td style="white-space:nowrap;width:10px" ></td>
		<td>
		    <%
             ' response.Write Server.MapPath("Tag/Tag_Style_"&activepanel&".asp")
                Server.Execute "Tag/Tag_Style_"&activepanel&".asp"          
            %>	
		</td>
	</tr>
</table>
<script type="text/javascript">
Window_GetElement(window,"navtd",true).ondblclick=function()
{
	if(event.shiftKey)
		alert(element.style.cssText);
}
</script>