<%@ Page Language="C#" AutoEventWireup="true" CodeFile="importer_timer.cs" Inherits="importer_timer" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">



<html xmlns="http://www.w3.org/1999/xhtml">

<head runat="server">
    <title>TimeOut import timer</title>
</head>
<body style="font-family:Arial" leftmargin="20" topmargin="20">

    


  <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	
	</td></tr></table>

    <form id="form1" runat="server">
     
         <!--
        HER: <=System.Reflection.Assembly.GetExecutingAssembly().Location%><br />
        HEr 2: <=AppDomain.CurrentDomain.SetupInformation.ConfigurationFile %><br />
        HEr 3: <=AppDomain.CurrentDomain.SetupInformation.ApplicationName %><br />
        HEr 5: <HttpContext.Current.GetSection("system.web/httpRuntime")%><br />
        -->
      

        <%
            
           
            
            //var section = HttpContext.Current.GetSection("system.web/httpRuntime") as System.Web.Configuration.HttpRuntimeSection;

            //Configuration config = WebConfigurationManager.OpenWebConfiguration("~/webconfig.config");
            //HttpRuntimeSection o = config.GetSection("system.web/httpRuntime") as HttpRuntimeSection;     

            //string userName = WebConfigurationManager.system.web.httpRuntime["httpRuntime"];
            
        %>

    <div>
    <h4>Importer timer til TimeOut MAKS 1000 linjer <span style="font-size:small; font-weight:normal;"><a href="../inc/xls/indlaestimerTemplate.csv">Download excel template her...</a></span></h4>
    *Bemærk at der skal være en overskriftslinje(header) med kolonne navne i excel filen.<br />
        <asp:Label ID="lblUploadStatus" runat="server" Text=""></asp:Label><br />
        
        <div>
         <asp:FileUpload ID="fu" runat="server" />
          <asp:Button ID="btnUpload" runat="server" Text="Upload" 
                onclick="btnUpload_Click" />
        </div><br />

        <div>
        <table>
        <tr>
        <td>
        Dine initialer:
        </td>
        <td>
            <asp:DropDownList ID="ddlMedarbejderId" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlMedarbejderId_SelectedIndexChanged" 
                ondatabound="ddlMedarbejderId_DataBound" AppendDataBoundItems="True">
                <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlMedarbejderId" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>      
        <td>
            <asp:Label ID="lblMedarbejderId" runat="server" Text=""></asp:Label>
        </td>
        </tr>
        <tr>
        <td>
        Job id:
        </td>
        <td>
            <asp:DropDownList ID="ddlJobId" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlJobId_SelectedIndexChanged" 
                ondatabound="ddlJobId_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlJobId" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblJobId" runat="server" Text=""></asp:Label>
        </td>
        </tr>
        <tr>
        <td>
        Timer
        </td>
        <td>
            <asp:DropDownList ID="ddlTimer" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlTimer_SelectedIndexChanged" 
                ondatabound="ddlTimer_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlTimer" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblTimer" runat="server" Text=""></asp:Label>
        </td>
        </tr>
        <tr>
        <td>
        Dato:
        </td>
        <td>
            <asp:DropDownList ID="ddlDato" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlDato_SelectedIndexChanged" 
                ondatabound="ddlDato_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlDato" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td> 
         <td>
            <asp:Label ID="lblDato" runat="server" Text=""></asp:Label>
        </td>
        </tr>
        <tr>
        <td>
        Aktivitetsnavn:
        </td>
        <td>
            <asp:DropDownList ID="ddlAktNavn" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlAktNavn_SelectedIndexChanged" 
                ondatabound="ddlAktNavn_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator5" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlAktNavn" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblAktNavn" runat="server" Text=""></asp:Label>
        </td>
        </tr>
        <tr>
        <td>
        Kommentar:
        </td>
        <td>
        <asp:DropDownList ID="ddlTimerKom" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlTimerKom_SelectedIndexChanged" 
                ondatabound="ddlTimerKom_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator6" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlTimerKom" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
        <td>
        <asp:Label ID="lblTimerKom" runat="server" Text=""></asp:Label>
        </td>
        </tr>
        </table>
        </div>
        <br />
        Filnavn
            <asp:TextBox ID="txtFileName" runat="server"></asp:TextBox>
        <asp:RequiredFieldValidator ID="RequiredFieldValidator7" runat="server" ErrorMessage="*" ForeColor="Red" ControlToValidate="txtFileName" ValidationGroup="Send"></asp:RequiredFieldValidator>
            <asp:Button ID="btnAdd" runat="server" Text="Send" onclick="btnAdd_Click" ValidationGroup="Send" /> til TimeOut    
            <br /><br />
        <asp:Label ID="lblStatus" runat="server" Text=""></asp:Label>         
    </div>
    </form>
</body>
</html>
