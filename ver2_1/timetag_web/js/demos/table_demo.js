$(function() {
  var table_1 = $('#table-1').dataTable ({
    "sAjaxSource": "./js/demos/table_data.json",
    "aoColumns": [
        { "mData": "id" },
        { "mData": "sort" },
        { "mData": "navn" },
        { "mData": "hovedgr"},
        { "mData": "sstgr"},
        { "mData": "medl" },
        { "mData": "timepr" },
        { "mData": "kstpr" },
        { "mData": "timer" }
    ],
    "fnInitComplete": function(oSettings, json) {
      $(this).parents ('.dataTables_wrapper').find ('.dataTables_filter input').prop ('placeholder', 'Søg i tabel...').addClass ('form-control input-sm')
    }
  })
})