<%@ Page Language="C#" AutoEventWireup="true" CodeFile="importer_akt.cs" Inherits="importer_akt" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">



<html xmlns="http://www.w3.org/1999/xhtml">

<head runat="server">
    <title>TimeOut import - Akt./ Sagslinjer</title>
</head>
<body style="font-family:Arial" leftmargin="20" topmargin="20">

    


  <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	
	</td></tr></table>

    <form id="form1" runat="server">
    <div>
    <h4>Importer aktiviteter til TimeOut - maks 10000 linjer <!--<span style="font-size:small; font-weight:normal;"><a href="../inc/xls/indlaesjobTemplate.csv">Download excel template her...</a></span>--></h4>
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
         <asp:label runat="server" ID="feltnr1navn">Job:</asp:label>
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
          <asp:label runat="server" ID="feltnr2navn">Beskrivelse:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlAktnavn" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlAktnavn_SelectedIndexChanged" 
                ondatabound="ddlAktnavn_DataBound" AppendDataBoundItems="True">
                <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlAktnavn" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>      
        <td>
            <asp:Label ID="lblAktnavn" runat="server" Text=""></asp:Label>
        </td>
        </tr>
        <tr>
        <td>
       
          <asp:label runat="server" ID="feltnr3navn">Løbenr. NAV:</asp:label>
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


        <tr>
        <td>
         <asp:label runat="server" ID="feltnr4navn">Konto:</asp:label>
        </td>
        <td>
        <asp:DropDownList ID="ddlKonto" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlKonto_SelectedIndexChanged" 
                ondatabound="ddlKonto_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator6" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlKonto" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
        <td>
        <asp:Label ID="lblKonto" runat="server" Text=""></asp:Label>
        </td>
        </tr>

         <tr>
        <td>
         <asp:label runat="server" ID="feltnr5navn">Type:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlLinjetype" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlLinjetype_SelectedIndexChanged" 
                ondatabound="ddlLinjetype_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator7" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlLinjetype" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblLinjetype" runat="server" Text=""></asp:Label>
        </td>
        </tr>

          <%if (importtype != "t2")
              {  %>


        <tr>
        <td>
         <asp:label runat="server" ID="feltnr6navn">Timer/Stk.:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlAkttimer" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlAkttimer_SelectedIndexChanged" 
                ondatabound="ddlAkttimer_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlAkttimer" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblAkttimer" runat="server" Text=""></asp:Label>
        </td>
        </tr>

             <tr>
        <td>
        <asp:label runat="server" ID="feltnr7navn">Stk. pris:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlAkttpris" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlAkttpris_SelectedIndexChanged" 
                ondatabound="ddlAkttpris_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlAkttpris" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblAkttpris" runat="server" Text=""></asp:Label>
        </td>
        </tr>
        

             <tr>
        <td>
         <asp:label runat="server" ID="feltnr8navn">Beløb:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlAktsum" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlAktsum_SelectedIndexChanged" 
                ondatabound="ddlAktsum_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator8" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlAktsum" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblAktsum" runat="server" Text=""></asp:Label>
        </td>
        </tr>

               <tr>
        <td>
         <asp:label runat="server" ID="feltnr9navn">Dato:</asp:label>
        </td>
        <td>
            <asp:DropDownList ID="ddlAktstDato" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlAktstDato_SelectedIndexChanged" 
                ondatabound="ddlAktstDato_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator10" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlAktstDato" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblAktstDato" runat="server" Text=""></asp:Label>
        </td>
        </tr>
        

        <%}; %>
       
        
       


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
