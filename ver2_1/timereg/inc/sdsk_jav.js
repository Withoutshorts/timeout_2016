







    $(document).ready(function () {


        

        $(".txtField").change(function () {


            //alert("fer")

            var thisid = this.id
            var thisvallngt = thisid.length
            var thisvaltrim = thisid.slice(13, thisvallngt)
            thisval = thisvaltrim

            jobid = $("#jobid_" + thisval).val()

            jobstatus = $(this).val()



            $.post("?jobid=" + jobid + "&jobstatus=" + jobstatus, { control: "FM_ajaxtype", AjaxUpdateField: "true" }, function (data) {
            });

            $("#sp_stopd_" + thisval).css("visibility", "visible")


            setTimeout(function () {
                // Do something after 5 seconds

                $("#sp_stopd_" + thisval).css("visibility", "hidden")
            }, 1000);

        });


        if (screen.width > 1480) {

            gblDivWdt = screen.width - 474
        }
        else {
            gblDivWdt = screen.width - 74
        }






        $("#sdskmain").width(gblDivWdt);



        //$("select[name*=ajax]").AjaxUpdateField({parent : "tr", subselector : "td:first > input[name=rowId]"});
        //<%if sortBy = "9" AND kview <> "j" AND print <> "j" then%>                  
        //        $("#incidentlist").table_sort({ items: 'tbody > tr:gt(1):has(:input[name=rowId])', IdControlNode : "td:first > :input[name=rowId]"});
        //       <%end if %>





        $(".ajx_estimat").click(function () {

            var thisid = this.id

            var thisvallngt = thisid.length
            var thisvaltrim = thisid.slice(14, thisvallngt)
            this_id_val = thisvaltrim

            //alert(this_id_val)

            //var thisval = this.value
            var thisval = $("#FM_ajaxesttid_" + this_id_val).val()




            //alert(oprvalTot)

            //var thisC = thisval
            //alert(thisval)

            $.post("?FM_estimat=" + thisval, { control: "FM_ajaxesttid", AjaxUpdateField: "true", cust: this_id_val }, function (data) {

                //window.location.reload();
                //$("#fajl").val(data);

                TONotifie("Estimat opdateret", true)

                // Beregner total //

                var oprvalTot = $("#totestimat").val().replace(".", "")
                oprvalTot = oprvalTot.replace(",", ".")
                oprvalTot = Math.round(oprvalTot * 100) / 100
                //alert(oprvalTot)

                thisval = thisval.replace(".", "")
                thisval = thisval.replace(",", ".")
                thisval = Math.round(thisval * 100) / 100
                //alert(thisval)

                var oprval = $("#FM_ajaxesttid_opr_" + this_id_val).val().replace(".", "")
                oprval = oprval.replace(",", ".")
                oprval = Math.round(oprval * 100) / 100
                //alert(oprval)

                var estimatTot = (oprvalTot / 1 - oprval / 1) + (thisval / 1)
                estimatTot = Math.round(estimatTot * 100) / 100
                estimatTot = String(estimatTot)
                estimatTot = estimatTot.replace(".", ",")

                $("#totestimat").val(estimatTot)

                thisval = String(thisval)
                thisval = thisval.replace(".", ",")
                $("#FM_ajaxesttid_opr_" + this_id_val).val(thisval)

            });

        });



    });





$(window).load(function() {
    // run code
    //$("#loadbar").hide(1000);
});



var newwindow = '';

function popitup(url) {

    url = url + "&kundeid=" + document.getElementById("FM_kontakt").value

    if (!newwindow.closed && newwindow.location) {
        newwindow.location.href = url;
    }
    else {
        newwindow = window.open(url, 'name', 'height=500,width=350,left=100,top=100');
        if (!newwindow.opener) newwindow.opener = self;
    }
    if (window.focus) { newwindow.focus() }
    return false;
}

function renssog() {
    document.getElementById("sog").value = ""
}


function BreakItUp() {
    //Set the limit for field size.
    var FormLimit = 302399

    //Get the value of the large input object.
    var TempVar = new String
    TempVar = document.theForm.BigTextArea.value

    //If the length of the object is greater than the limit, break it
    //into multiple objects.
    if (TempVar.length > FormLimit) {
        document.theForm.BigTextArea.value = TempVar.substr(0, FormLimit)
        TempVar = TempVar.substr(FormLimit)

        while (TempVar.length > 0) {
            var objTEXTAREA = document.createElement("TEXTAREA")
            objTEXTAREA.name = "BigTextArea"
            objTEXTAREA.value = TempVar.substr(0, FormLimit)
            document.theForm.appendChild(objTEXTAREA)

            TempVar = TempVar.substr(FormLimit)
        }
    }
}

function visduedate() {
    //alert(document.getElementById("FM_useduedate").checked)
    if (document.getElementById("FM_useduedate").checked == true) {
        document.getElementById("FM_useduedate_dag").disabled = false
        document.getElementById("FM_useduedate_md").disabled = false
        document.getElementById("FM_useduedate_aar").disabled = false
        document.getElementById("FM_duetime").disabled = false
        //document.getElementById("FM_esttid").disabled = false
    } else {
        document.getElementById("FM_useduedate_dag").disabled = true
        document.getElementById("FM_useduedate_md").disabled = true
        document.getElementById("FM_useduedate_aar").disabled = true
        document.getElementById("FM_duetime").disabled = true
        //document.getElementById("FM_esttid").disabled = true
    }
}

//function setEmneFocus(){
//alert("her")
//document.getElementById("sdsk_emne").focus() 
//}




function showbesk(besk_id) {
    var entryDiv = $("#b_" + besk_id);
    (entryDiv.css("display") == "none") ? entryDiv.stop().slideDown(200) : entryDiv.stop().slideUp(200);
    /*if (lastOpen != 0) {
    document.getElementById("b_"+lastOpen).style.display = "none";
    document.getElementById("b_"+lastOpen).style.visibility = "hidden";
    }
    
    document.getElementById("b_"+besk_id).style.display = "";
    document.getElementById("b_"+besk_id).style.visibility = "visible";
    
    document.getElementById("FM_div_id_txtbox_"+besk_id).focus()
    //document.getElementById("FM_div_id_txtbox_"+besk_id).scrollintoview()

    
    document.getElementById("divopen").value = besk_id*/
    return false;
}

function lukemldiv() {
    document.getElementById("emldiv").style.display = "none";
    document.getElementById("emldiv").style.visibility = "hidden";
}