<%@ CODEPAGE=65001 %> 
<% Response.CodePage=65001%> 
<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<fieldset>
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
		<!-- x-resize -->
	</select>
</fieldset>

<fieldset>
	<legend>
	<%= GetString("Filter") %>
	</legend>
	<input type="text" id="inp_filter" style="width:240px" /> <!--button filter builder-->
</fieldset>
<div id="outer"><div id="div_demo"><%= GetString("DemoText") %></div></div>

<script type="text/javascript">

var OxO8243=["sel_cursor","inp_filter","outer","div_demo","cssText","style","cursor","value","filter"];var sel_cursor=Window_GetElement(window,OxO8243[0x0],true);var inp_filter=Window_GetElement(window,OxO8243[0x1],true);var outer=Window_GetElement(window,OxO8243[0x2],true);var div_demo=Window_GetElement(window,OxO8243[0x3],true); function UpdateState(){ div_demo[OxO8243[0x5]][OxO8243[0x4]]=element[OxO8243[0x5]][OxO8243[0x4]] ;}  ; function SyncToView(){ sel_cursor[OxO8243[0x7]]=element[OxO8243[0x5]][OxO8243[0x6]] ;if(Browser_IsWinIE()){ inp_filter[OxO8243[0x7]]=element[OxO8243[0x5]][OxO8243[0x8]] ;} ;}  ; function SyncTo(element){ element[OxO8243[0x5]][OxO8243[0x6]]=sel_cursor[OxO8243[0x7]] ;if(Browser_IsWinIE()){ element[OxO8243[0x5]][OxO8243[0x8]]=inp_filter[OxO8243[0x7]] ;} ;}  ;

</script>