





$(document).ready(function() {

  

    

});

$(window).load(function() {
    // run code
    $("#loadbar").hide(1000);
});




	    $(document).ready(function() {

	        if (document.getElementById("FM_visprjob_ell_sum2").checked == true) { 
	            $("#ysiu").css("display", "");
	            $("#ysiu").css("visibility", "visible");
	        }
            
        
	        $("#FM_visprjob_ell_sum2").click(function() {
	            $("#ysiu").css("display", "");
	            $("#ysiu").css("visibility", "visible");
	            $("#ysiu").show("slow");
	        });


	        $("#FM_visprjob_ell_sum0").click(function() {
	            $("#ysiu").css("display", "none");
	            $("#ysiu").css("visibility", "hidden");
	            $("#ysiu").hide("slow");
	        });

	        $("#FM_visprjob_ell_sum1").click(function() {
	            $("#ysiu").css("display", "none");
	            $("#ysiu").css("visibility", "hidden");
	            $("#ysiu").hide("slow");
	        });

	    });


	



	  
	
	
function BreakItUp2()
{
    //Set the limit for field size.
    var FormLimit = 102399
	
    //Get the value of the large input object.
    var TempVar = new String
    TempVar = document.theForm2.BigTextArea.value
	
    //If the length of the object is greater than the limit, break it
    //into multiple objects.
    if (TempVar.length > FormLimit)
    {
        document.theForm2.BigTextArea.value = TempVar.substr(0, FormLimit)
        TempVar = TempVar.substr(FormLimit)
	
        while (TempVar.length > 0)
        {
            var objTEXTAREA = document.createElement("TEXTAREA")
            objTEXTAREA.name = "BigTextArea"
            objTEXTAREA.value = TempVar.substr(0, FormLimit)
            document.theForm2.appendChild(objTEXTAREA)
	
            TempVar = TempVar.substr(FormLimit)
        }
    }
}
	
function visjob() {
    id = document.getElementById("id").value
    //alert("default.asp?typethis="+tp+"&clicked="+cl+"&FM_fab="+fab+"")
    window.location.href = "materiale_stat.asp?menu=stat&id="+id
}	
	
function clearJobsog(){
    document.getElementById("FM_jobsog").value = ""
}


function clearK_Jobans() {
    //alert("her")
    document.getElementById("cFM_jobans").checked = false
    document.getElementById("cFM_jobans2").checked = false
    document.getElementById("cFM_kundeans").checked = false
}


function setshowprogrpval() {
    document.getElementById("showprogrpdiv").value = 1
}


function closeprogrpdiv() {
    document.getElementById("progrpdiv").style.visibility = "hidden"
    document.getElementById("progrpdiv").style.display = "none"
}


function curserType(tis) {
    document.getElementById(tis).style.cursor = 'hand'
}




function gkAlly(val) {

    antal_x = document.getElementById("antal_x").value
    //antal_x = 10
    for (i = 0; i < antal_x; i++) {
       
        txt = document.getElementById("mfidgk_id_" + i).options[document.getElementById("mfidgk_id_" + i).selectedIndex].text
        //alert(document.getElementById("mfidgk_id_" + i).value + " i:"+ i +" # " + txt)
        //document.getElementById("mfidgk_id_" + i).selectedIndex = val
        document.getElementById("mfidgk_id_" + i).value = val

        if (val == 0) {
            document.getElementById("mfidgk_id_" + i).options[document.getElementById("mfidgk_id_" + i).selectedIndex].text = "Nej"
        } else {
            document.getElementById("mfidgk_id_" + i).options[document.getElementById("mfidgk_id_" + i).selectedIndex].text = "Ja"
        }
    }
}

function afAlly(val) {

    antal_x = document.getElementById("antal_x").value
    //antal_x = 10
    for (i = 0; i < antal_x; i++) {

        txt = document.getElementById("mfid_id_" + i).options[document.getElementById("mfid_id_" + i).selectedIndex].text
        //alert(document.getElementById("mfidgk_id_" + i).value + " i:"+ i +" # " + txt)
        //document.getElementById("mfidgk_id_" + i).selectedIndex = val
        document.getElementById("mfid_id_" + i).value = val

        if (val == 0) {
            document.getElementById("mfid_id_" + i).options[document.getElementById("mfid_id_" + i).selectedIndex].text = "Nej"
        } else {
            document.getElementById("mfid_id_" + i).options[document.getElementById("mfid_id_" + i).selectedIndex].text = "Ja"
        }
    }
}         
        
        
        

	
