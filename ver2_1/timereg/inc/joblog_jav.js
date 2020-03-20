


	




$(document).ready(function() {


   
    $("#joblog_uge3").click(function () {
        if ($(this).is(':checked') == true) {

            $("#bruguge").attr('checked', false);
            $("#brugmd").attr('checked', true);
        }

    });


    $("#FM_jobsog").keyup(function () {
        //alert("her")
        //$("#FM_job option:selected").val()
        $("#FM_job option[value='0']").attr('selected', 'selected');
    });


    
    $("#bruguge").click(function () {

        if ($(this).is(':checked') == true) {
       
            $("#brugmd").attr('checked', false);

            if ($("#joblog_uge3").is(':checked') == true) {
                $("#joblog_uge1").attr('checked', true);
            }
        }

    });

    $("#brugmd").click(function () {

        if ($(this).is(':checked') == true) {

            $("#bruguge").attr('checked', false);
        }

    });

    $("#chkalle_0").click(function() {
        if ($(this).is(':checked') == true) {
            $(".akt_afst").attr('checked', true);
        } else {
            $(".akt_afst").attr('checked', false);
        }

    });

    $("#chkalle_1").click(function() {
        if ($(this).is(':checked') == true) {
            $(".akt_udspec").attr('checked', true);
        } else {
            $(".akt_udspec").attr('checked', false);
        }

    });

    $("#chkalle_2").click(function() {
        if ($(this).is(':checked') == true) {
            $(".akt_flex").attr('checked', true);
        } else {
            $(".akt_flex").attr('checked', false);
        }

    });

    $("#chkalle_3").click(function() {
        if ($(this).is(':checked') == true) {
            $(".akt_ferie").attr('checked', true);
        } else {
            $(".akt_ferie").attr('checked', false);
        }

    });

    $("#chkalle_4").click(function() {
        if ($(this).is(':checked') == true) {
            $(".akt_overarb").attr('checked', true);
        } else {
            $(".akt_overarb").attr('checked', false);
        }

    });

    $("#chkalle_5").click(function() {
        if ($(this).is(':checked') == true) {
            $(".akt_syg").attr('checked', true);
        } else {
            $(".akt_syg").attr('checked', false);
        }

    });




    $("#load").hide(1000)


});







 function checkAll(val) {


     //FM_akttype_id_1
     //alert(val)
     antal_v = document.getElementById("antal_v_" + val).value
     //alert(antal_v)

     if (document.getElementById("chkalle_" + val).checked == true) {
         tval = true
     } else {
         tval = false
     }

     //alert(tval)

     for (i = 0; i < antal_v; i++)
         //alert(i)
         document.getElementById("FM_akttype_id_" + val + "_" + i).checked = tval
 }
	
function mOvr(divId,src,clrOver) {
    src.bgColor = clrOver;	
}
	
function clearJobsog(){
    document.getElementById("FM_jobsog").value = ""
}
	
	
	
function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	
	
	
	
	
function showjoblogsubdiv(){
    document.getElementById("joblogsub").style.visibility = "visible"
    document.getElementById("joblogsub").style.display = "" 
}
	
function hidejoblogsubdiv(){
    document.getElementById("joblogsub").style.visibility = "hidden"
    document.getElementById("joblogsub").style.display = "none" 
}
	
function closeprogrpdiv(){
    document.getElementById("progrpdiv").style.visibility = "hidden"
    document.getElementById("progrpdiv").style.display = "none" 
}
	
function setshowprogrpval(){
    document.getElementById("showprogrpdiv").value = 1
}
	
function curserType(tis){
    document.getElementById(tis).style.cursor='hand'
}
	
	
function checkAllGK(val){
	
    antal_v = document.getElementById("antal_v").value 
	
    //for (i = 0; i < antal_v ; i++)
     //   document.getElementById("FM_godkendt" + val + "_" + i).checked = true

    $(".FM_godkendt_"+ val +"").attr('checked', true);
}





function clearK_Jobans() {
    //alert("her")
    document.getElementById("cFM_jobans").checked = false
    document.getElementById("cFM_jobans2").checked = false
    document.getElementById("cFM_kundeans").checked = false
}
	
	
