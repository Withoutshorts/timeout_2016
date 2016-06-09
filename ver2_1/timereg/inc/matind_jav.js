function rydSog() {
    document.getElementById("FM_sog").value = ""
}


function opd_jobid(id) {
    thisJobidoptSel = document.getElementById("jobid_sel").selectedIndex
    ////alert(thisJobidoptSel)
    document.getElementById("jobid_" + id).value = document.getElementById("jobid_sel").options[thisJobidoptSel].value
}


function beregnsalgspris(matid) {
    avaId = document.getElementById("FM_avagruppe_" + matid).value
    avaVal = document.getElementById("avagrpval_" + avaId).value.replace(",", ".")


    kobspris = document.getElementById("FM_kobspris_" + matid).value.replace(",", ".")
    nysalgspris = ((kobspris / 1 * avaVal) / 100 + kobspris / 1)
    nysalgspris = Math.round(nysalgspris * 100) / 100
    document.getElementById("FM_salgspris_" + matid).value = nysalgspris
    document.getElementById("FM_salgspris_" + matid).value = document.getElementById("FM_salgspris_" + matid).value.replace(".", ",")
}





function beregnsalgsprisOTF() {
    avaId = document.getElementById("gruppe").value
    avaVal = document.getElementById("avagrpval_" + avaId).value.replace(",", ".")


    kobspris = document.getElementById("pris").value.replace(",", ".")
    nysalgspris = ((kobspris / 1 * avaVal) / 100 + kobspris / 1)
    nysalgspris = Math.round(nysalgspris * 100) / 100
    document.getElementById("FM_salgspris").value = nysalgspris
    document.getElementById("FM_salgspris").value = document.getElementById("FM_salgspris").value.replace(".", ",")

}






         $(document).ready(function() {
                    


             $("#intkode").change(function () {


                 var thisval = this.value
                 //alert(thisval)

                 if (thisval == 2) {
                     $("#tr_slgs").css("visibility", "visible");
                 } else {
                     $("#tr_slgs").css("visibility", "hidden");
                 }


             });
        
       

             $("#sel_valuta").change(function() {

                 $(".s_valuta").val($("#sel_valuta").val()) 

             });


             $("#sel_internkode").change(function() {

                 $(".s_internkode").val($("#sel_internkode").val()) 

             });


                

             $("#sel_regdato").keyup(function() {
                 $(".s_regdato").val($("#sel_regdato").val()) 
             });




                
             $("#sel_matgrp").change(function() {

                  
                 $(".s_matgrp").val($("#sel_matgrp").val()) 
                
                 avaId = $("#sel_matgrp").val()
                 avaVal = $("#avagrpval_"+avaId).val()
                 lto = $("#lto").val()
                                
        
        
                 antalx = $("#antal_x").val() 
                 for (i=0;i<=antalx;i++) {
                                        
                     //alert("dd")
                     matid = $("#FM_matid_"+i).val() 
        
                     kobspris = $("#FM_kobspris_"+matid).val().replace(",",".") 
                     nysalgspris = ((kobspris/1 * avaVal) / 100 + kobspris/1)
                                        
                     if (lto == "dencker") { // 3 decimaler
                         nysalgspris = Math.round(nysalgspris*1000)/1000
                     } else {
                         nysalgspris = Math.round(nysalgspris*100)/100
                     } 
                                        
                     nysalgspris = String(nysalgspris).replace(".",",") 
	                                    
                     //alert(nysalgspris)
                     $("#FM_salgspris_"+matid).val(nysalgspris)
	                                                       
                        
                 }



             });
                
                
                    


                	
  

         });
















