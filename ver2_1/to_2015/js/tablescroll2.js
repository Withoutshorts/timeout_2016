


$(document).ready(function () {

    //alert($("#tablebody").find("tr").length)
   /*alert(document.getElementById('example').rows[0].cells.length)
    alert(document.getElementById('example').rows[1].cells.length)
    alert(document.getElementById('example').rows[2].cells.length)
    alert(document.getElementById('example').rows[3].cells.length)
    alert(document.getElementById('example').rows[4].cells.length)
    alert(document.getElementById('example').rows[5].cells.length) */

    var headerColumnCount = document.getElementById('example').rows[0].cells.length

    i = 0
   while (i < headerColumnCount - 1) {
        // alert("awd")
        $(".tr_linje").append("<td style='display:none;'></td>")
        i = i + 1
    }

    var table_lto = $('#table_lto').val();
    //alert(table_lto)

    if (table_lto == "oko")
    {
        fixedColumns_ = 6
    } 
    else
    {
        fixedColumns_ = 5
    }

    var table = $('#example').DataTable({
    scrollY: "600px",
    scrollX: true,
    scrollCollapse: true,
    paging: false,
    ordering: false,
    fixedColumns: {
        leftColumns: fixedColumns_
    }

    });

   // alert("hephep")

   /* $("#fitScreen_chb").click(function () {

        alert("Hawd")
        var table = $('#example').DataTable();
        table.columns.adjust().draw();
        alert("hep")
    }); */

   /* alert("hephep")
   var headerColumnCount = document.getElementById('example').rows[0].cells.length
    //alert(headerColumnCount)
    var totaltrs = $("#tablebody").find("tr").length

    //alert(totaltrs)

    var bodyColumnCount = $(".medarbfelter_tr").find("td").length
   // alert(bodyColumnCount)
    var bodyheaderDiff = headerColumnCount - bodyColumnCount */

   // alert(bodyheaderDiff)
    
  
  /*  $('#example > tbody > tr').each(function () {
        thisid = this.id

        while (this.cells.length < headerColumnCount)
        {
            var newTdElement = document.createElement("td")
            newTdElement.style.display = "none"
            newTdElement.className = "addedTdElement"
            this.appendChild(newTdElement)
        }


    }); /*

   //makeTable();

   /* $("#fitScreen_chb").click(function () {

        alert("fitfit")

        $('.addedTdElement').each(function () {

            this.remove();
           // alert("removed")
        });

        var headerColumnCount = document.getElementById('example').rows[0].cells.length
        //alert(headerColumnCount)
        var totaltrs = $("#tablebody").find("tr").length

        //alert(totaltrs)

        var bodyColumnCount = $(".medarbfelter_tr").find("td").length
        // alert(bodyColumnCount)
        var bodyheaderDiff = headerColumnCount - bodyColumnCount

        // alert(bodyheaderDiff)

        $('#example').DataTable().destroy();

        $('#example > tbody > tr').each(function () {
            thisid = this.id

            while (this.cells.length < headerColumnCount) {
                var newTdElement = document.createElement("td")
                newTdElement.style.display = "none"
                newTdElement.className = "addedTdElement"
                this.appendChild(newTdElement)
            }


        });

        //makeTable();

    }); */
    
/*
  function makeTable() {

        tableWidth = document.getElementById('example').offsetWidth
        alert(tableWidth)


        if (tableWidth > 900)
        {
            var table = $('#example').DataTable({
                scrollY: "300px",
                scrollX: "1250px",
                scrollCollapse: true,
                paging: false,
                ordering: false,
                fixedColumns: {
                    leftColumns: 1
                }

            });
        }
        else 
        {       
            var table = $('#example').DataTable({
                scrollY: "300px",
                // scrollX: true,
                scrollCollapse: true,
                paging: false,
                ordering: false//,
                // fixedColumns: {
                //  leftColumns: 1
                // }

            });
        } 
  

        alert("Mogens2")

        alert("te")

        alert("awd")
    } */
});

