<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<fieldset style="padding:4px;">
	<legend>
	<%= GetString("Cursor") %>
	</legend>
	<select id="sel_cursor">
		<option value=""><%= GetString("NotSet") %></option>
		<option value="Default"><%= GetString("Default") %></option>
		<option value="Move"><%= GetString("Move") %></option>
		<option value="Text">Text</option>
		<option value="Wait">Wait</option>
		<option value="Help">Help</option>
		<option value="Pointer">Pointer</option>
		<option value="Hand">Hand</option>
		<option value="Progress">Progress </option>
		<option value="not-allowed">Not-Allowed</option>
		<!-- x-resize -->
	</select>
</fieldset>

<fieldset style="padding:4px;">
	<legend>
	<%= GetString("Filter") %>
	</legend>
	<input type="text" id="inp_filter" style="width:240px" /> <!--button filter builder-->
</fieldset>

<div id="outer" style="height:100px; margin-bottom:10px; padding:5px;"><div id="div_demo"><%= GetString("DemoText") %></div></div>
<br />

<script type="text/javascript">

var sel_cursor=Window_GetElement(window,"sel_cursor",true);
var inp_filter=Window_GetElement(window,"inp_filter",true);
var outer=Window_GetElement(window,"outer",true);
var div_demo=Window_GetElement(window,"div_demo",true);

function UpdateState()
{
	div_demo.style.cssText=element.style.cssText;
}

function SyncToView()
{
	sel_cursor.value=element.style.cursor;
	if(Browser_IsWinIE())
	{
		inp_filter.value=element.style.filter;
	}
}
function SyncTo(element)
{
	element.style.cursor=sel_cursor.value;
	if(Browser_IsWinIE())
	{
		element.style.filter=inp_filter.value;
	}
}

</script>