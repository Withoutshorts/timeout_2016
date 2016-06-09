// JScript File

$(window).load(function() {
    // run code


//if jq_func
//getAktlisten()

   

});

$(document).ready(function () {

    $("select[name*=ajax]").AjaxUpdateField({ parent: "tr", subselector: "td:first > input[name=rowId]" });
    $("#incidentlist").table_sort({ items: 'tbody > tr:gt(1):has(:input[name=rowId])', IdControlNode: "td:first > :input[name=rowId]" });



    $("#a_sel_all").click(function () {
        $(".cls_todo_rel").attr('checked', true);
    });


    $("#a_sel_none").click(function () {
        $(".cls_todo_rel").attr('checked', false);
    });





});  


