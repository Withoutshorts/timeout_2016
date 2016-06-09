<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="HandlingSmtpExceptions.aspx.vb" Inherits="HandlingSmtpExceptions" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <h4>Handling SMTP Exceptions</h4>
    <p>
        This demo shows how to display an error message if there is a problem sending an email. It also
        shows how to assign the relay server settings to the <code>SmtpClient</code> class programmatically
        (versus having them spelled out in <code>Web.config</code>). To illustrate this, enter some invalid
        relay server settings or invalid email address values.
    </p>
    
    <table border="0">
        <tr>
            <td><b>From Email:</b></td>
            <td><asp:TextBox runat="server" ID="UsersEmail" Columns="30"></asp:TextBox></td>
        </tr>
        <tr>
            <td><b>Relay Server Hostname:</b></td>
            <td><asp:TextBox runat="server" ID="Hostname" Columns="30"></asp:TextBox></td>
        </tr>
        <tr>
            <td><b>Port:</b></td>
            <td><asp:TextBox runat="server" ID="Port" Columns="4"></asp:TextBox></td>
        </tr>
        <tr>
            <td><b>Username:</b></td>
            <td><asp:TextBox runat="server" ID="Username" Columns="20"></asp:TextBox></td>
        </tr>
        <tr>
            <td><b>Password:</b></td>
            <td><asp:TextBox runat="server" ID="Password" Columns="15" TextMode="Password"></asp:TextBox></td>
        </tr>
        <tr>
            <td colspan="2" align="center">
                <asp:Button runat="server" ID="SendEmail" Text="Send Test Email" />
                &nbsp;
                <asp:Button ID="LoadBadValues" runat="server" Text="Load Invalid Values" /></td>
        </tr>
    </table>
</asp:Content>

