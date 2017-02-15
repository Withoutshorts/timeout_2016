<%@ Page Language="C#" AutoEventWireup="true" CodeFile="importer_job_monitor.cs" Inherits="importer_job_monitor" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<html xmlns="http://www.w3.org/1999/xhtml">

<head runat="server">
    <title>TimeOut import job - Monitor</title>
</head>
<body style="font-family:Arial" leftmargin="20" topmargin="20">

    


  <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	
	</td></tr></table>

    <form id="form1" runat="server">
        <asp:HiddenField id="importtype" value="2" runat=Server />


          <label ID="lbl_importtype" runat="server">..</label>
        
        

    <div>
    <h4>Importer job til TimeOut - maks 8000 linjer <span style="font-size:small; font-weight:normal;"><a href="../inc/xls/indlaesjobTemplate_monitor.csv">Download excel template her...</a></span></h4>
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
        Kundenavn:
        </td>
        <td>
            <asp:DropDownList ID="ddlKundenavn" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlKundenavn_SelectedIndexChanged" 
                ondatabound="ddlKundenavn_DataBound" AppendDataBoundItems="True">
                <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator8" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlKundenavn" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>      
        <td>
            <asp:Label ID="lblKundenavn" runat="server" Text=""></asp:Label>
        </td>
        </tr>


             

      
        <tr>
        <td>
        Jobnavn:
        </td>
        <td>
            <asp:DropDownList ID="ddlJobnavn" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlJobnavn_SelectedIndexChanged" 
                ondatabound="ddlJobnavn_DataBound" AppendDataBoundItems="True">
                <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlJobnavn" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>      
        <td>
            <asp:Label ID="lblJobnavn" runat="server" Text=""></asp:Label>
        </td>
        </tr>

        <tr>
        <td>
        Jobnr.:
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
        Job startdato:
        </td>
        <td>
            <asp:DropDownList ID="ddlstDato" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlstDato_SelectedIndexChanged" 
                ondatabound="ddlstDato_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlstDato" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td> 
         <td>
            <asp:Label ID="lblstDato" runat="server" Text=""></asp:Label>
        </td>
        </tr>

         <tr>
        <td>
        Job slutdato:
        </td>
        <td>
            <asp:DropDownList ID="ddlslDato" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlslDato_SelectedIndexChanged" 
                ondatabound="ddlslDato_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator5" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlslDato" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td> 
         <td>
            <asp:Label ID="lblslDato" runat="server" Text=""></asp:Label>
        </td>
        </tr>



      <tr>
        <td>
        Sortering:
        </td>
        <td>
        <asp:DropDownList ID="ddlSort" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlSort_SelectedIndexChanged" 
                ondatabound="ddlSort_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator9" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlSort" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
        <td>
        <asp:Label ID="lblSort" runat="server" Text=""></asp:Label>
        </td>
        </tr>

        <tr>
        <td>
        Forretningsområde:
        </td>
        <td>
        <asp:DropDownList ID="ddlFomr" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlFomr_SelectedIndexChanged" 
                ondatabound="ddlFomr_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator10" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlFomr" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
        <td>
        <asp:Label ID="lblFomr" runat="server" Text=""></asp:Label>
        </td>
        </tr>

         <tr>
        <td>
        Aktivitetsnavn:
        </td>
        <td>
        <asp:DropDownList ID="ddlAktnavn" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlAktnavn_SelectedIndexChanged" 
                ondatabound="ddlAktnavn_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator11" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlAktnavn" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
        <td>
        <asp:Label ID="lblAktnavn" runat="server" Text=""></asp:Label>
        </td>
        </tr>

         <tr>
        <td>
        Stk. antal:
        </td>
        <td>
            <asp:DropDownList ID="ddlAntal" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlAntal_SelectedIndexChanged" 
                ondatabound="ddlAntal_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlAntal" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblAntal" runat="server" Text=""></asp:Label>
        </td>
        </tr>


         <tr>
        <td>
        Akt. startdato:
        </td>
        <td>
        <asp:DropDownList ID="ddlAktstdato" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlAktstdato_SelectedIndexChanged" 
                ondatabound="ddlAktstdato_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator12" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlAktstdato" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
        <td>
        <asp:Label ID="lblAktstdato" runat="server" Text=""></asp:Label>
        </td>
        </tr>

           <tr>
        <td>
        Akt. slutdato:
        </td>
        <td>
        <asp:DropDownList ID="ddlAktsldato" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlAktsldato_SelectedIndexChanged" 
                ondatabound="ddlAktsldato_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator13" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlAktsldato" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
        <td>
        <asp:Label ID="lblAktsldato" runat="server" Text=""></asp:Label>
        </td>
        </tr>

            <tr>
        <td>
        Aktivitet varenr: (Rap. nr.)
        </td>
        <td>
        <asp:DropDownList ID="ddlAktvarenr" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlAktvarenr_SelectedIndexChanged" 
                ondatabound="ddlAktvarenr_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator14" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlAktvarenr" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
        <td>
        <asp:Label ID="lblAktvarenr" runat="server" Text=""></asp:Label>
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
        <asp:Label ID="lblStatus" runat="server" Text="..."></asp:Label>         
    </div>
    </form>
</body>
</html>
