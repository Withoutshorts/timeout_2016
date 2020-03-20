


function SendSMS()
{
    var username = "8C051w0E";
    var password = "JU1larVD";



    var data = JSON.stringify({
        "source": "LINK",
        "destination": "+4526842000",
        "userData": "Hello world",
        "platformId": "COOL",
        "platformPartnerId": "30978",
        "useDeliveryReport": false
    });

    var xhr = new XMLHttpRequest();
    xhr.withCredentials = true;

    //xhr.addEventListener("readystatechange", function () {
    //    if (this.readyState === 4) {
     //       console.log(this.responseText);
    //    }
    //});

    //xhr.open("POST", "https://wsx.sp247.net/sms/send");

    //xhr.setRequestHeader("authorization", "Basic " + btoa(username + ":" + password));
    //xhr.setRequestHeader("content-type", "application/json");
    //xhr.setRequestHeader("Access-Control-Allow-Origin", "https://wsx.sp247.net");
    //xhr.setRequestHeader("Access-Control-Allow-Origin", "https://timeout.cloud");
    //xhr.setRequestHeader("Access-Control-Allow-Credentials", "true");
    //xhr.setRequestHeader("Access-Control-Allow-Origin: *");
     //xhr.setRequestHeader("content-type", "text/plain");
    //xhr.setRequestHeader("cache-control", "no-cache");


    xhr.open("POST", "https://gatewayapi.com/rest/mtsms");
    
    xhr.setRequestHeader("authorization", "Basic " + btoa(username + ":" + password));
    xhr.setRequestHeader("content-type", "application/json");
    xhr.setRequestHeader("Access-Control-Allow-Origin", "*"); 
    //xhr.setRequestHeader("Access-Control-Allow-Credentials","true"); 
    //xhr.setRequestHeader("Access-Control-Allow-Origin: *");
     //xhr.setRequestHeader("content-type", "text/plain");
    //xhr.setRequestHeader("cache-control", "no-cache");
    

    xhr.send(data);

    alert("done")
}
