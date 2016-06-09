<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="HtmlEmail.aspx.vb" Inherits="HtmlEmail" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <h4>Send an HTML-Formatted Email</h4>
    
    <p>
        This demo shows how an email message can include HTML formatting...
    </p>
    <p>
        <b>Your Email Address: </b>
        <asp:TextBox runat="server" ID="UsersEmail" Columns="30"></asp:TextBox>
    </p>
    <p>
        <asp:Button ID="SendHTMLEmail" runat="server" Text="Send an HTML-Formatted Email" />
        &nbsp;
    </p>
</asp:Content>

