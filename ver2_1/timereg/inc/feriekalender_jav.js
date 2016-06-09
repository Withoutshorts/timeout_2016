





$(document).ready(function() {

  

    

});

$(window).load(function() {
    // run code
    $("#loadbar").hide(1000);
});





	
	function showdatoinfo(id){
	    //alert(id)
	    document.getElementById("divid_info_"+id).style.visibility = "visible"
	    document.getElementById("divid_info_"+id).style.display = ""
	}
	
function hidedatoinfo(id){
    document.getElementById("divid_info_"+id).style.visibility = "hidden"
    document.getElementById("divid_info_"+id).style.display = "none"
}
	
