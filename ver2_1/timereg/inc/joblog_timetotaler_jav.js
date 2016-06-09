
	function BreakItUp()
	{
	  //Set the limit for field size.
	  var FormLimit = 102399 
	
	  //Get the value of the large input object.
	  var TempVar = new String
	  TempVar = document.theForm.BigTextArea.value

	  //alert(TempVar.length)
        
	  //If the length of the object is greater than the limit, break it
	  //into multiple objects.
	  if (TempVar.length > FormLimit)
	  {
	    document.theForm.BigTextArea.value = TempVar.substr(0, FormLimit)
	    TempVar = TempVar.substr(FormLimit)
	
	    while (TempVar.length > 0)
	    {
	      var objTEXTAREA = document.createElement("TEXTAREA")
	      objTEXTAREA.name = "BigTextArea"
	      objTEXTAREA.value = TempVar.substr(0, FormLimit)
	      document.theForm.appendChild(objTEXTAREA)

	      TempVar = TempVar.substr(FormLimit)

          //alert("her")
	    }
	  }
	}
	


	function BreakItUp2()
	{
	  //Set the limit for field size.
	  var FormLimit = 302399
	
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


	


	function clearJobsog() {
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

	$(window).load(function () {
	    // run code
	    $(".load").hide(1000);

	    

	});



	$(document).ready(function () {



	    //$("#sbm_pivot").click(function () {

	    //    $("#csv_pivot").val('1')

	    //});

	    //$("#sbm_csv").click(function () {

	    //    $("#csv_pivot").val('0')

	    //});


	    if ($("#directexp").is(':checked') == false) {
	        $("#csv_pivot0").attr('checked', true);
	        $("#csv_pivot1").attr('disabled', true);
	    }


	    $("#directexp").click(function () {

	        if ($("#directexp").is(':checked') == false) {
	            $("#csv_pivot0").attr('checked', true);
	            $("#csv_pivot1").attr('disabled', true);
	        } else {
	            $("#csv_pivot1").attr('disabled', false);
	        }

	    });

	    


	    $("#redaktor").submit(function () {
	        //alert("Er du sikker?")
	        var r = confirm("Er du sikker du vil opdatere timepriserne?");
	        if (r == true) {
	            return true
	        }
	        else {
	            return false
	        }


	    });


	  

	    function vissp_redaktor() {
	        if ($("#vis_fakbare_res1").is(':checked') == true) {

	            $("#sp_redaktor").css("display", "");
	            $("#sp_redaktor").css("visibility", "visible");
	            $("#sp_redaktor").show(1000);

	        } else {

	            $("#sp_redaktor").hide(1000);

	        }
	    }


	    $("#FM_jobsog").keyup(function () {
	        //alert("her")
	        //$("#FM_job option:selected").val()
	        $("#FM_job option[value='0']").attr('selected', 'selected');
	    });


	    $("#vis_aktnavn").click(function () {
	        if ($("#vis_aktnavn").is(':checked') == false) {
	            document.getElementById("FM_aktnavnsog").value = ""
	        }
	    });

	    $("#FM_udspec").click(function () {
	        udspec_all();
	    });

	    $(".s_tp").mouseover(function () {
	        $(this).css("cursor", "pointer");

	    });

	    $(".s_tp").click(function () {
	        var thisid = this.id
	        //alert(thisid)
	        var thisvallngt = thisid.length

	        var thisvaltrim = thisid.slice(5, thisvallngt)
	        var thisval = thisvaltrim

	        //alert(thisval)

	        var nthisval_lngt = thisval.length
	        var x_val_lngt = thisval.indexOf("__");
	        //alert(x_val_lngt)

	        var x_val = thisval.slice(x_val_lngt + 2, nthisval_lngt)
	        //alert(x_val)

	        var ncls = thisval.slice(0, x_val_lngt)
	        //alert(ncls)

	        var timerThis = $("#f_tp_" + x_val).val()
	        $(".f_tp_" + ncls).val(timerThis)

	        //alert(timerThis)

	    });


	    function udspec_all() {
	        if ($("#FM_udspec").is(':checked') == true) {
	            $("#FM_udspec_all").removeAttr('disabled')
	            //$("#FM_udspec_all input:checked").attr('checked', true)
	            //$('input[name=FM_udspec_all]').attr('checked', true)
	        } else {
	            $("#FM_udspec_all").attr('disabled', true);
	            //$("#FM_udspec_all").attr('checked', false);
	            $("#FM_udspec_all").removeAttr('checked')
	        }
	    }


	    udspec_all();
	    //vissp_redaktor();

	});