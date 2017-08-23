<%@ Page Language="C#" AutoEventWireup="true" CodeFile="importer_med.cs" Inherits="importer_med" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">



<html xmlns="http://www.w3.org/1999/xhtml">

<head runat="server">
    <title>TimeOut import - Medarb.linjer</title>
</head>
<body style="font-family:Arial" leftmargin="20" topmargin="20">

    


  <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	
	</td></tr></table>

    <form id="form1" runat="server">
    <div>
    <h4>Importer medarbejdere til TimeOut - maks 10000 linjer <!--<span style="font-size:small; font-weight:normal;"><a href="../inc/xls/indlaesjobTemplate.csv">Download excel template her...</a></span>--></h4>
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
         <asp:label runat="server" ID="feltnr1navn">No. (Init):</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlMinit" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlMinit_SelectedIndexChanged" 
                ondatabound="ddlMinit_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator5" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlMinit" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblMinit" runat="server" Text=""></asp:Label>
        </td>
        </tr>

        <tr>
        <td>
          <asp:label runat="server" ID="feltnr2navn">Navn:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlNavn" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlNavn_SelectedIndexChanged" 
                ondatabound="ddlNavn_DataBound" AppendDataBoundItems="True">
                <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlNavn" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>      
        <td>
            <asp:Label ID="lblNavn" runat="server" Text=""></asp:Label>
        </td>
        </tr>
        <tr>
        <td>
       
          <asp:label runat="server" ID="feltnr3navn">Email:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlEmail" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlEmail_SelectedIndexChanged" 
                ondatabound="ddlEmail_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlEmail" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblEmail" runat="server" Text=""></asp:Label>
        </td>
        </tr>


        <tr>
        <td>
         <asp:label runat="server" ID="feltnr4navn">Norm time:</asp:label>
        </td>
        <td>
        <asp:DropDownList ID="ddlNorm" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlNorm_SelectedIndexChanged" 
                ondatabound="ddlNorm_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator6" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlNorm" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
        <td>
        <asp:Label ID="lblNorm" runat="server" Text=""></asp:Label>
        </td>
        </tr>

         <tr>
        <td>
         <asp:label runat="server" ID="feltnr5navn">Employment date:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlAnsatdato" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlAnsatdato_SelectedIndexChanged" 
                ondatabound="ddlAnsatdato_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator7" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlAnsatdato" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblAnsatdato" runat="server" Text=""></asp:Label>
        </td>
        </tr>

          

        <tr>
        <td>
         <asp:label runat="server" ID="feltnr6navn">Termination date:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlOpsagtdato" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlOpsagtdato_SelectedIndexChanged" 
                ondatabound="ddlOpsagtdato_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlOpsagtdato" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblOpsagtdato" runat="server" Text=""></asp:Label>
        </td>
        </tr>

             <tr>
        <td>
        <asp:label runat="server" ID="feltnr7navn">Blocked:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlMansat" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlMansat_SelectedIndexChanged" 
                ondatabound="ddlMansat_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlMansat" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblMansat" runat="server" Text=""></asp:Label>
        </td>
        </tr>
        

             <tr>
        <td>
         <asp:label runat="server" ID="feltnr8navn">Expence vendor no.:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlEvn" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlEvn_SelectedIndexChanged" 
                ondatabound="ddlEvn_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator8" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlEvn" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblEvn" runat="server" Text=""></asp:Label>
        </td>
        </tr>

             <tr>
        <td>
         <asp:label runat="server" ID="feltnr9navn">Costcenter:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlCostcenter" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlCostcenter_SelectedIndexChanged" 
                ondatabound="ddlCostcenter_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator10" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlCostcenter" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblCostcenter" runat="server" Text=""></asp:Label>
        </td>
        </tr>

             <tr>
        <td>
         <asp:label runat="server" ID="feltnr10navn">Linemanager:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlLinemanager" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlLinemanager_SelectedIndexChanged" 
                ondatabound="ddlLinemanager_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator11" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlLinemanager" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblLinemanager" runat="server" Text=""></asp:Label>
        </td>
        </tr>

             <tr>
        <td>
         <asp:label runat="server" ID="feltnr11navn">Country code:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlCcode" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlCcode_SelectedIndexChanged" 
                ondatabound="ddlCcode_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator12" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlCcode" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblCcode" runat="server" Text=""></asp:Label>
        </td>
        </tr>

             <tr>
        <td>
         <asp:label runat="server" ID="feltnr12navn">Weblang:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlWeblang" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlWeblang_SelectedIndexChanged" 
                ondatabound="ddlWeblang_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator13" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlWeblang" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblWeblang" runat="server" Text=""></asp:Label>
        </td>
        </tr>
        

       
       


        </table>
        </div>
        <br />
        Filnavn:
            <asp:TextBox ID="txtFileName" runat="server"></asp:TextBox>
        <asp:RequiredFieldValidator ID="RequiredFieldValidator9" runat="server" ErrorMessage="*" ForeColor="Red" ControlToValidate="txtFileName" ValidationGroup="Send"></asp:RequiredFieldValidator>
            <asp:Button ID="btnAdd" runat="server" Text="Send" onclick="btnAdd_Click" ValidationGroup="Send" /> til TimeOut    
            <br /><br />
        <asp:Label ID="lblStatus" runat="server" Text="..."></asp:Label>         
    </div>
    </form>
</body>
</html>
