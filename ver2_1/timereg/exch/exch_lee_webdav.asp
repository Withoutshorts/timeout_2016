<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/header_hvd_inc.asp"-->


<script>
var strCookies;
     var strRequest;
     var objXmlResult;
     
     try
     {
         strCookies = doFBALogin("https://exchangeserver/public", "exchangeserver", "domain\\username ", "password");
     
         strRequest = "<D:searchrequest xmlns:D = \"DAV:\" ><D:sql>SELECT \"DAV:displayname\" FROM \"https://exchangeserver/public\" WHERE \"DAV:ishidden\" = false</D:sql></D:searchrequest>";
    
        objXmlResult = Search("https://exchangeserver/public", strRequest, strCookies);
   
        // Do something with the XML object
    } 
    catch (e) 
    {
        // Handle the error!
    }

	
	
	
	
	function Search(strUrl, strQuery, strCookies)
    {
         // Send the request to the server
        var xmlHTTP = new ActiveXObject("Microsoft.xmlhttp");
         var i;
     
         xmlHTTP.open("SEARCH", strUrl, false);
         xmlHTTP.setRequestHeader("Content-type:", "text/xml"); 
     
        // Add the cookies to the request. According to Microsoft article Q234486, the cookie has to be set two times.
        xmlHTTP.setRequestHeader("Cookies", "Necessary according to Q234486");
        xmlHTTP.setRequestHeader("Cookies", strCookies);
        xmlHTTP.send(strQuery);
    
        return xmlHTTP.responseXML;
    }
    
    
    function doFBALogin(strDestination, strServer, strUsername, strPassword) 
    {
        var xmlHTTP = new ActiveXObject("Microsoft.xmlhttp");
        var strUrl
            var strContent;
        var arrHeader;
        var strCookies;
        var intCookies;
        var i;
    
        // Prepare the URL for FBA login
        strUrl = "https://" + strServer + "/exchweb/bin/auth/owaauth.dll";
    
        xmlHTTP.open("POST", strUrl, false); 
    
        xmlHTTP.setRequestHeader("Content-type:", "application/x-www-form-urlencoded"); 
    
        // Generate the body for FBA login
        strContent = "destination=" + strDestination + "&username=" + strUsername + "&password=" + strPassword;
        xmlHTTP.send(strContent);
    
        // Get all response headers
        arrHeader = xmlHTTP.getAllResponseHeaders().split("
		
		strCookies = "";
    intCookies = 0;
    
        
		// Iterate through the collection of returned headers
        for (i = 0; i<arrHeader.length; i++) 
        {
            // If the entry is a cookies, extract the name/value
           if (arrHeader[i].indexOf("Set-Cookie") != -1) 
           {
               strCookies += arrHeader[i].substr(12) + "; ";
                intCookies++;
            }
        }
    
        // Check if two cookies have been returned. Otherwise throw an exception
        if (intCookies < 2) throw "Could not log in to OWA!";
        return strCookies;
    }
</script>


<!--#include file="../inc/regular/footer_inc.asp"-->

