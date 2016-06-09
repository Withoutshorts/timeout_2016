

<%@  Import Namespace="System"%>
<%@  Import Namespace="System.Data"%>
<%@  Import Namespace="System.Data.Odbc" %>
<%@  Import Namespace="System.Web.Services" %>



<!--
 Page Language="VB" Inherits="cati" Src="class_cati.vb" Debug="true"
-->




<%@ Page Language="VB"  Inherits="cac" Src="cac.vb" Debug="true" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<script runat="server" language="vbscript">

    

    Sub Page_load()
        
       
        'If Not IsPostBack Then
       
     

 
        'Dim objLibaryDR As OdbcDataReader
        Dim objLibaryConn As OdbcConnection
        Dim objLibaryCmd As OdbcCommand
        Dim objDataSet As New DataSet

        'Dim proxySample As New CATIService.CATI()  ' Proxy object.
        
        
        
        Dim strConn2 As String = "Driver={MySQL ODBC 3.51 Driver};Server=81.19.249.35;Database=timeout_epi;User=outzource;Password=SKba200473;"
        'Dim strConn2 As String = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_intranet;User=outzource;Password=SKba200473;"
       
        Dim strSQL As String = "SELECT id, medarbejderId, jobId, timer, dato, komm FROM job_imp_timer WHERE id BETWEEN 0 AND 1"


        objLibaryConn = New OdbcConnection(strConn2)
        'objLibaryCmd = New OdbcCommand(strSQL, objLibaryConn)
        'objLibaryConn.open()


        Dim objAdapter As New OdbcDataAdapter(strSQL, objLibaryConn)
        objAdapter.Fill(objDataSet, "job_imp_timer")

        Dim objDataView As New DataView(objDataSet.Tables("job_imp_timer"))
        
        
        
      
        'dgNameList.DataSource = objDataView
        'dgNameList.DataBind()
        
        'Call upDateCatiTimer2(objDataSet)
        
        errLabel.Text = strSQLaTxt & meID
        
        

        'Call upDateCatiTimer(objDataSet)


        
        
        

        'End If
        
    End Sub

   
    
</script>


<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    

    <h4>Reading Data from the connection
    <asp:Label ID="datasrc" runat="server"></asp:Label> to the DataGrid</h4>

    <asp:DataGrid ID="dgNameList"  autogeneratecolumns="true" runat="server" />

    <asp:Label ID="lb_aktid" runat="server"></asp:Label> 
    <asp:Label ID="errLabel" runat="server"></asp:Label> 
    

       
    
</body>
</html>
