﻿<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>

<%

strConn = "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=" & Server.mappath("./YourSheet.xls") & ";UID=admin;"


 %>
    <form id="form1" runat="server">
    <div>
    
    </div>
    </form>
</body>
</html>
