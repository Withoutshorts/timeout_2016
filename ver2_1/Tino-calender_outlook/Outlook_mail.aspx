<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Outlook_mail.aspx.cs" Inherits="Outlook_mail" EnableEventValidation="false" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head  runat="server">
    <meta charset="utf-8" />

    <!-- New: Bootstrap Date-Picker Plugin -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.4.1/css/bootstrap-datepicker3.css" />

    <!--New: For jquery template-->
    <script src="http://code.jquery.com/jquery-latest.min.js"></script>
    <script src="http://ajax.microsoft.com/ajax/jquery.templates/beta1/jquery.tmpl.min.js"></script>

    <!--New: Bootstrap muliselect dropdown-->
    <link rel="stylesheet" href="http://netdna.bootstrapcdn.com/bootstrap/3.0.2/css/bootstrap.min.css" />
    <script src="http://netdna.bootstrapcdn.com/bootstrap/3.0.2/js/bootstrap.min.js"></script>

    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-multiselect/0.9.13/css/bootstrap-multiselect.css" />
    <script src="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-multiselect/0.9.13/js/bootstrap-multiselect.js"></script>

    <!-- New: Bootstrap Date-Picker Plugin -->
    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.4.1/js/bootstrap-datepicker.min.js"></script>

    <!-- jQuery validation -->
    <script type="text/javascript" src="https://cdn.jsdelivr.net/jquery.validation/1.15.1/jquery.validate.min.js"></script>

    <!-- BarDragger. draggable bar element in between panels -->
    <script src="https://static.codepen.io/assets/editor/pen/index-85a7604f435488c27ecf9076db52aa007293ef1665c535f985df3891dc5b1797.js"></script>
    <script src="https://static.codepen.io/assets/common/browser_support-f97f3ef9d187e4d0436e96e4d002603225517239e71f56f5a0441cce3a0d5be4.js"></script>
    <!-- 2005, 2014 jQuery Foundation  -->
    <script src="https://static.codepen.io/assets/common/everypage-eb67a83772e5fbc39195f098ad2e9fafa910022bcae0317bc5b3873698665a9e.js"></script>
    <script src="https://static.codepen.io/assets/packs/js/vendor-1aaaa35c84754a5df90e.chunk.js"></script>
    <script src="https://static.codepen.io/assets/packs/js/2-1aaaa35c84754a5df90e.chunk.js"></script>
    <script src="https://static.codepen.io/assets/packs/js/everypage-1aaaa35c84754a5df90e.js"></script>
    <script src="https://static.codepen.io/assets/packs/js/processorRouter-1aaaa35c84754a5df90e.js"></script>

    <!--  jQuery UI - v1.11.3 - 2015-03-05 -->
    <script src="https://static.codepen.io/assets/editor/global/commonLibs-5aa21c051f554186721dfb2e22efa0262d769cfc8be7ee60a1c6e0e44c2bcd81.js"></script>
    <script src="https://static.codepen.io/assets/editor/global/codemirror-41b003ba7076ac1285f386080f36114ec6e4c72796832faf66404c01df1eb104.js"></script>
    <script src="https://static.codepen.io/assets/libs/emmet-codemirror-plugin-222bd3cb88c16bb29433e34064a6dce2845b15c040718116c240719eaafc143f.js"></script>
  


       <title>Her kan du tilføje din ferie til din Outlook kalender.</title>

</head>
<body>


<asp:Button ID="btnOutlookCalender" Text="Get Calender" runat="server" OnClick="Send_email_with_outlookCalender" />



</body>




</html>