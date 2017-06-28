<%@ Page Language="C#" AutoEventWireup="true" CodeFile="importer_omk.cs" Inherits="importer_omk" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">



<html xmlns="http://www.w3.org/1999/xhtml">

<head runat="server">
    <title>TimeOut import - Omkostninger</title>
</head>
<body style="font-family:Arial" leftmargin="20" topmargin="20">

    


  <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	
	</td></tr></table>

    <form id="form1" runat="server">
    <div>
    <h4>Importer omkostninger til TimeOut - maks 10000 linjer <!--<span style="font-size:small; font-weight:normal;"><a href="../inc/xls/indlaesjobTemplate.csv">Download excel template her...</a></span>--></h4>
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
         <asp:label runat="server" ID="feltnr1navn">Bogføringsdato:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlDato" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlDato_SelectedIndexChanged" 
                ondatabound="ddlDato_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator6" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlDato" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblDato" runat="server" Text=""></asp:Label>
        </td>
        </tr>

        <tr>
        <td>
          <asp:label runat="server" ID="feltnr2navn">Beskrivelse:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlBesk" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlBesk_SelectedIndexChanged" 
                ondatabound="ddlBesk_DataBound" AppendDataBoundItems="True">
                <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlBesk" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>      
        <td>
            <asp:Label ID="lblBesk" runat="server" Text=""></asp:Label>
        </td>
        </tr>
        
        
        <tr>
        <td>
         <asp:label runat="server" ID="feltnr3navn">Konto:</asp:label>
        </td>
        <td>
        <asp:DropDownList ID="ddlKonto" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlKonto_SelectedIndexChanged" 
                ondatabound="ddlKonto_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlKonto" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
        <td>
        <asp:Label ID="lblKonto" runat="server" Text=""></asp:Label>
        </td>
        </tr>


         <tr>
        <td>
         <asp:label runat="server" ID="feltnr4navn">Jobnr:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlJobId" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlJobId_SelectedIndexChanged" 
                ondatabound="ddlJobId_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator5" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlJobId" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblJobId" runat="server" Text=""></asp:Label>
        </td>
        </tr>

         <tr>
        <td>
         <asp:label runat="server" ID="feltnr5navn">Medarbejder (Initialer):</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlMinit" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlMinit_SelectedIndexChanged" 
                ondatabound="ddlMinit_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator7" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlMinit" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblMinit" runat="server" Text=""></asp:Label>
        </td>
        </tr>

          
        <tr>
        <td>
         <asp:label runat="server" ID="feltnr6navn">Beløb:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlBelob" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlBelob_SelectedIndexChanged" 
                ondatabound="ddlBelob_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator8" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlBelob" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblBelob" runat="server" Text=""></asp:Label>
        </td>
        </tr>

         <tr>
        <td>
        
               <asp:label runat="server" ID="Label3">Valuta:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlValuta" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlValuta_SelectedIndexChanged" 
                ondatabound="ddlValuta_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlValuta" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblValuta" runat="server" Text=""></asp:Label>
        </td>
        </tr>


   <tr>
        <td>
              <asp:label runat="server" ID="feltnr7navn">Løbenr. Nav:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlAktnr" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlAktnr_SelectedIndexChanged" 
                ondatabound="ddlAktnr_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlAktnr" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblAktnr" runat="server" Text=""></asp:Label>
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
