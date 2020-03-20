<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "../include_Security.asp" --><div class="normal">
The Style - List-XXX
</div>
<div id="outer"><div id="div_demo"><%= GetString("DemoText") %></div></div>
<script type="text/javascript">

function UpdateState()
{
	div_demo.style.cssText=element.style.cssText;
}

function SyncToView()
{
	
}
function SyncTo(element)
{

}

</script>
