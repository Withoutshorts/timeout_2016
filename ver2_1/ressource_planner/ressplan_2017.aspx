<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ressplan_2017.aspx.cs" Inherits="ressplan_2017" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Ressource Planning</title>
    <%--    <script type="text/javascript" src="js/jquery-2.2.3.js"></script>--%>
    <%--<script type="text/javascript" src="http://code.jquery.com/jquery-1.10.0.min.js"></script>--%>
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/r/bs-3.3.5/dt-1.10.9/datatables.min.css" />
    <%--<script type="text/javascript" src="js/jquery.inputmask.bundle.js"></script>--%>
    <%--<link href="css/font-awesome.min.css" rel="stylesheet" />--%>

    <%--New css--%>

    <!-- Bruges til LESS styling af TO menu -->
    <link href='//fonts.googleapis.com/css?family=Open+Sans:400,300,600,700,800' rel='stylesheet' type='text/css' />
    <link href="css/chronograph.css" rel="stylesheet" type="text/css" />
    <!-- Google Font: Open Sans -->
    <link rel="stylesheet" href="//fonts.googleapis.com/css?family=Open+Sans:400,400italic,600,600italic,800,800italic" />
    <link rel="stylesheet" href="//fonts.googleapis.com/css?family=Oswald:400,300,700" />
    <!-- Font Awesome CSS -->
    <link rel="stylesheet" href="css/font-awesome.min.css" />
    <!-- Bootstrap CSS -->
    <link rel="stylesheet" href="css/bootstrap.min.css" />
    <!-- Plugin CSS -->
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/r/bs-3.3.5/dt-1.10.9/datatables.min.css" />
    <%--<link rel="stylesheet" href="css/bootstrap-datepicker3.css" />--%>
    <link rel="stylesheet" href="//cdnjs.cloudflare.com/ajax/libs/jasny-bootstrap/3.1.3/css/jasny-bootstrap.min.css" />
    <!-- App CSS -->
    <link rel="stylesheet" href="css/mvpready-admin.css" />
    <%--<link rel="stylesheet" href="css/mvpready-flat.css" />--%>
    <!-- Custom styles for TimeOut  -->
    <link href="css/mpvready-style-timeout.css" rel="stylesheet" />

    <!-- New-->
    <link href="css/jquery.multiselect.css" rel="stylesheet" />

    <!-- New: Bootstrap Date-Picker Plugin -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.4.1/css/bootstrap-datepicker3.css" />


    <!-- Core JS -->
    <%--<script src="js/libs/bootstrap.min.js"></script>--%>
    <!-- Plugin JS -->
    <%--<script src="//cdnjs.cloudflare.com/ajax/libs/jasny-bootstrap/3.1.3/js/jasny-bootstrap.min.js"></script>--%>

    <%--New: For jquery template--%>
    <script src="http://code.jquery.com/jquery-latest.min.js"></script>
    <script src="http://ajax.microsoft.com/ajax/jquery.templates/beta1/jquery.tmpl.min.js"></script>


    <%--New: Bootstrap muliselect dropdown--%>
    <link rel="stylesheet" href="http://netdna.bootstrapcdn.com/bootstrap/3.0.2/css/bootstrap.min.css" />
    <script src="http://netdna.bootstrapcdn.com/bootstrap/3.0.2/js/bootstrap.min.js"></script>

    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-multiselect/0.9.13/css/bootstrap-multiselect.css" />
    <script src="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-multiselect/0.9.13/js/bootstrap-multiselect.js"></script>

    <!-- New: Bootstrap Date-Picker Plugin -->
    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.4.1/js/bootstrap-datepicker.min.js"></script>

    <!-- New: Moment JS -->
    <script type="text/javascript" src="js/moment.js"></script>

    <%-- jQuery validation --%>
    <script type="text/javascript" src="https://cdn.jsdelivr.net/jquery.validation/1.15.1/jquery.validate.min.js"></script>


    <style>
        .container{width:auto; margin:0px 20px;}
        table th, table td { font-size: 11px; min-width: 35px;background-color: white !important; }
        table td {padding: 5px !important;}
        table th, .v-middle {vertical-align: middle !important;}
        .weekend-color {background-color: lightgray !important;}
        .today-color {background-color: #ffd1d1 !important;}
        .bg-color {background-color: #f5f5f5 !important;}
        .pointer {cursor: pointer;}
        .ddcls { vertical-align: middle; border-top: hidden; width: 140px;border-right-width: 1px !important; }
        .month-header { text-align: center; border-top: hidden; background-color: white; font-size: 15px; }
        .day-header, .week-header { text-align: center; border-top: hidden; background-color: white; }
        .day-header .header, .week-header .header {font-size: 15px;}
        .month-day-ths, .month-date-ths { background-color: #f5f5f5; text-align: center; vertical-align: middle !important;}
        .center-align {text-align: center;}
        .header-prev-margin { margin-right: -27px; font-size: 22px; line-height: 23px; }
        .header-next-margin { margin-left: -29px; font-size: 22px; line-height: 23px; }
        .header-prev-margin a, .header-next-margin a {position: relative;z-index: 9;}
        .header-div-marg {margin-left: 276px;}
        .dayheader-div-marg {margin-left: 276px;}
        .monthheader-div-marg {margin-left: 468px;}
        .nodays-color {background-color: #e7e3e3;}
        .data-row, .data-row th {background-color: white !important;}
        .multiselect-container > li > a > label { padding: 8px 20px 8px 40px !important; margin-bottom: 3px !important; margin-top: 3px !important; }
        .wrapper > .content {margin: 0px;}
        *[class*='dropdown-'] {height: auto !important;}
        .activity-modal {font-size: 12px;}
        .activity-modal .modal-body > .row {margin-bottom: 10px;}
        .lh-34 {line-height: 34px;}
        .form-control {font-size: 11px !important;}
        .multiselect-selected-text { font-size: 11px; float: left; line-height: 20px; font-weight: bold; width: 110px; text-align: left; color: #555; text-overflow: ellipsis; overflow: hidden; }
        .error {border: 1px solid #b12121 !important;}
        .error-txt{color:#b12121;}
        .main-ac-data { min-height: 35px; padding: 2px; color: black; font-size: 12px; margin-bottom: 3px; width: 100%; position: relative;float:left; }
        .main-ac-data .ac-name { text-overflow: ellipsis; overflow: hidden; text-align: left; line-height: 15px; white-space: nowrap; }
        .main-ac-data .ac-icon { position: absolute; bottom: 0; right: 3px; }
        .main-ac-data .ac-icon a.glyphicon.glyphicon-refresh { color: #2b2b2b; font-size: 10px; font-weight: bold; display: inline-block; text-decoration: none; }
        .main-ac-data .ac-icon a.imp { width: 11px; height: 11px; background-color: #2b2b2b; display: inline-block; border-radius: 100%; margin-right: 5px; font-size: 8px; font-weight: bold; text-decoration: none; text-align:center; }
        .day-data-td {min-width:90px;}
        .day-data-td .main-ac-data{width:100%; max-width:71px;}
        .week-data-td {min-width:145px;}
        .week-data-td .main-ac-data{width:100%; max-width:none;}
        .month-full-width{max-width:40px; min-height:40px}
        .month-half-width{max-width:16px; margin: 0 2px 4px; min-height:50px;}
        .month-half-width .ac-icon {margin-left:2px;}
        .month-half-width .ac-icon a.imp{margin-right:0px;}
        .week-full-width{max-width:200px !important;}
        .week-half-width{max-width:45px !important; margin: 0 2px 4px;}
        .day-full-width{max-width:130px !important;}
        .day-half-width{max-width:60px !important; margin: 0 2px 4px;}
        .modal-title{color:white;}
        .modal-header .close{margin-top:6px;}
        .pl0{padding-left:0px;}
        .rec-row{line-height:32px;}
        .brd-table { min-height: 400px;border-left: 1px solid #ccc;border-right: 1px solid #ccc;border-bottom: 1px solid #ccc;}
        .brd-table > table{ margin-top:0px !important;}
        .fit-to-screen{ font-size: 14px;font-weight: bold;}
        .fit-to-screen > input{ float: left;margin-top: 22px;}
        .fit-to-screen > div{ float: left;margin-top: 21px;margin-left: 5px;}
        #daytbl .header-div-marg, #weektbl .header-div-marg{margin-left:276px;}
        table.dataTable {margin-bottom:20px !important;}
    </style>
</head>
<body>

    <%--Modal popup--%>

    <!--<label id="lblsql" runat="server">SQL:  </label>-->

    <form id="taskform" runat="server" name="taskModalForm">
        <div class="wrapper">
            <div class="content">
                <div class="container">
                    <div class="portlet">
                        <div>
                            <h3 class="portlet-title">
                                <u>Ressource Planner</u>
                                <div class="pull-right fit-to-screen">
                                    <input type="checkbox" checked="checked" onchange="setFitToScreen(this)" />
                                    <div>Fit To Screen</div>
                                </div>                                
                            </h3>                            
                        </div>
                        
                        <div class="portlet-body">
                            <div id="activityModal" class="modal fade activity-modal" tabindex='-1'>
                                <div class="modal-dialog">
                                    <div class="modal-content">
                                        <div class="modal-header">
					                        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">×</span></button>
					                        <h3 class="modal-title">Update Activity</h3>
				                        </div>

                                        <div class="modal-body">
                                            <br />

                                            <div class="row">
                                                <div class="col-lg-2 lh-34">Start:</div>
                                                <div class="col-lg-4">
                                                    <div class='input-group date start'>
                                                        <input type="text" id="dtstart" class="form-control input-small" name="dts" value="" placeholder="dd-mm-yyyy" />
                                                        <span class="input-group-addon input-small">
                                                            <span class="fa fa-calendar open-datetimepicker"></span>
                                                        </span>
                                                    </div>
                                                </div>
                                                <div class="col-lg-2 lh-34"></div>
                                                <div class="col-lg-4">
                                                    <input type="time" id="fromtime" name="fromtime" class="form-control input-small" placeholder="00:00" />
                                                </div>
                                            </div>
                                            <div class="row">
                                                <div class="col-lg-2 lh-34">End:</div>
                                                <div class="col-lg-4">
                                                    <div class='input-group date end'>
                                                        <input type="text" id="dtend" class="form-control input-small" name="dts" value="" placeholder="dd-mm-yyyy" />
                                                        <span class="input-group-addon input-small">
                                                            <span class="fa fa-calendar"></span>
                                                        </span>
                                                    </div>
                                                    <p class="error-txt" style="display:none;">End Date should be more than start date!</p>
                                                </div>
                                                <div class="col-lg-2 lh-34"></div>
                                                <div class="col-lg-4">
                                                    <input type="time" id="totime" name="totime" class="form-control input-small" placeholder="00:00" />
                                                </div>
                                            </div>
                                            <!--
                                            <div class="row">
                                                <div class="col-lg-2 lh-34">Recurrence:</div>
                                                <div class="col-lg-4">
                                                    <select class="form-control input-small" id="recurrencedd" name="recurrencedd" onchange="recurrenceChange()">                                                        
                                                        <option value="-1">No Recurrence</option>
                                                        <option value="0">Daily</option>
                                                        <option value="1">Weekly</option>
                                                        <option value="2">Every Month</option>
                                                        <option value="3">Every 2. Month</option>
                                                        <option value="4">Every 3. Month</option>
                                                        <option value="5">Every 6. Month</option>
                                                        <option value="6" onclick="selmaaned.call">Every Year</option>
                                                    </select>
                                                </div>
                                                <div class="col-lg-2 lh-34">Important:</div>
                                                <div class="col-lg-4 lh-34">
                                                    <input id="chkimp" name="chkimp" type="checkbox" value="0" />
                                                </div>
                                            </div>
                                            
                                            <div class="row no-recurrence">
                                                <div class="col-lg-2 lh-34">End after:</div>
                                                <div class="col-lg-2">
                                                    <input id="noOfRecurrence" name="noOfRecurrence" maxlength="2" type="text" class="form-control input-small" />
                                                </div>
                                                <div class="col-lg-2 pl0 v-middle rec-row">
                                                    Recurrences
                                                </div>
                                            </div>
                                             -->
                                            <div class="row"></div>
                                            <div class="row">
                                                <div class="col-lg-2 lh-34">Employee:</div>
                                                <div class="col-lg-4">
                                                    <select id="empchk" name="empchk" class="form-control input-small">
                                                        <option value="1">Hans</option>
                                                        <option value="2">Per</option>
                                                        <option value="3">Søren Karlsen</option>
                                                    </select>
                                                </div>
                                                <div class="col-lg-2 lh-34">Customer:</div>
                                                <div class="col-lg-4">
                                                    <input id="custtxt" name="custtxt" type="text" value="Outzource Aps" class="form-control input-small" readonly="readonly" disabled="disabled" />
                                                </div>
                                            </div>
                                            <div class="row">
                                                <div class="col-lg-2 lh-34">Heading:</div>
                                                <div class="col-lg-4">
                                                    <select id="headtxt" name="headtxt" class="form-control input-small">
                                                        <option value="1">Activity</option>
                                                        <option value="2">Project</option>
                                                        <option value="3">Customer</option>
                                                    </select>
                                                </div>

                                                <div class="col-lg-2 lh-34">Project:</div>
                                                <div class="col-lg-4">
                                                    <input id="projecttxt" name="projecttxt" type="text" value="Design" class="form-control input-small" readonly="readonly" disabled="disabled" />
                                                </div>
                                            </div>
                                            <div class="row">
                                                <div class="col-lg-2 lh-34">Type:</div>
                                                <div class="col-lg-4">
                                                    <input id="actType" name="actType" type="text" value="" class="form-control input-small" readonly="readonly" disabled="disabled" />
                                                </div>

                                                <div class="col-lg-2 lh-34">Activity:</div>
                                                <div class="col-lg-4">
                                                    <input id="prodtxt" name="prodtxt" type="text" value="Produkion" class="form-control input-small" readonly="readonly" disabled="disabled" />
                                                </div>
                                            </div>
                                        </div>

                                        <div class="modal-footer" style="text-align: left">
                                            <button type="button" onclick="saveActivity()" class="btn btn-default">Submit</button>
                                        </div>
                                    </div>
                                </div>
                            </div>

                            <div id="maindiv">
                                <div id="headerdiv" class="table-responsive brd-table"></div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

        </div>
    </form>

    <input type="hidden" name="hidweek" id="hidweek" />
    <input type="hidden" name="hidmonth" id="hidmonth" />
    <input type="hidden" name="hidyear" id="hidyear" />
    <input type="hidden" name="hidstartdate" id="hidstartdate" />
    <input type="hidden" name="hidprecallmonth" id="hidprecallmonth" />
    <input type="hidden" name="hidprecallyear" id="hidprecallyear" />
    <input type="hidden" name="hidprecallstartdate" id="hidprecallstartdate" />

    <script id="mainTemplate" type="text/x-jquery-tmpl">
    </script>

    <script id="dayheaderTemplate" type="text/x-jquery-tmpl">
        <table id="daytbl" class="table dataTable table-bordered ui-datatable removetbl">

            <tr>
                <th class="ddcls">
                    <select id="viewtypedaydd" onchange="viewTypeChange();" class="viewtype form-control input-small"></select>
                </th>
                <th colspan="70" class="day-header">
                    <%--<a href="javascript:void(0);"><</a> &nbsp Monday 01-01-16 &nbsp <a href="javascript:void(0);">></a>--%>

                    <div class="row">
                        <div class="col-md-6 dayheader-div-marg">
                            <div class="col-md-1 header-prev-margin">
                                <a href="javascript:void(0);" onclick="previous()">
                                    <i class="glyphicon glyphicon-menu-left"></i>
                                </a>
                            </div>
                            <div class="col-md-6 header">
                                <span id="currentdayView" class="currentdateview" data-type="day" data-day="" data-date="" data-month="" data-year="" data-week=""></span>
                            </div>
                            <div class="col-md-1 header-next-margin">
                                <a href="javascript:void(0);" onclick="next()">
                                    <i class="glyphicon glyphicon-menu-right"></i>
                                </a>
                            </div>
                        </div>
                        <div class="col-md-3">
                            <div class="pull-right">
                                <a href="javascript:void(0);" onclick="changeDayGroup('day');" class="btn btn-secondary btn-sm"><b>Day</b></a>
                                <a href="javascript:void(0);" onclick="changeDayGroup('week');" class="btn btn-default btn-sm"><b>Week</b></a>
                                <a href="javascript:void(0);" onclick="changeDayGroup('month');" class="btn btn-default btn-sm"><b>Month</b></a>
                            </div>
                        </div>
                    </div>
                </th>
            </tr>
            <tr>
                <th class="ddcls">
                    <select id="sortbydaydd" onchange="sortViewByChange();" class="sortby form-control input-small"></select>
                </th>
                <th colspan="69" style="border-right: hidden"></th>
            </tr>
            <tr id="timetr" class="last-header">
                <th id="th1" class="ddcls">
                    <select id="oliverdaydd" onchange="resourceChange()" multiple="multiple" class="oliver form-control input-small"></select>
                </th>
            </tr>

        </table>
    </script>
    <script id="weekheaderTemplate" type="text/x-jquery-tmpl">
        <table id="weektbl" class="table dataTable table-bordered ui-datatable test removetbl">
            <tr>
                <th class="ddcls">
                    <select id="viewtypeweekdd" onchange="viewTypeChange();" class="viewtype form-control input-small"></select>
                </th>
                <th colspan="29" class="week-header">
                    <%--<a href="javascript:void(0);"><</a> &nbsp Uge 1 - Januar 2016 &nbsp <a href="javascript:void(0);">></a>--%>

                    <div class="row">
                        <div class="col-md-6 header-div-marg">
                            <div class="col-md-1 header-prev-margin">
                                <a href="javascript:void(0);" onclick="previous()">
                                    <i class="glyphicon glyphicon-menu-left"></i>
                                </a>
                            </div>
                            <div class="col-md-7 header">
                                <span id="currentweekView" class="currentdateview" data-type="week" data-week="" data-month="" data-year="" data-startdate="" data-nextstartdate=""></span>
                            </div>
                            <div class="col-md-1 header-next-margin">
                                <a href="javascript:void(0);" onclick="next()">
                                    <i class="glyphicon glyphicon-menu-right"></i>
                                </a>
                            </div>
                        </div>
                        <div class="col-md-3">
                            <div class="pull-right">
                                <a href="javascript:void(0);" onclick="changeDayGroup('day');" class="btn btn-default btn-sm"><b>Day</b></a>
                                <a href="javascript:void(0);" onclick="changeDayGroup('week');" class="btn btn-secondary btn-sm"><b>Week</b></a>
                                <a href="javascript:void(0);" onclick="changeDayGroup('month');" class="btn btn-default btn-sm"><b>Month</b></a>
                            </div>
                        </div>
                    </div>
                </th>
            </tr>
            <tr id="weektbldaystr">
                <th class="ddcls">
                    <select id="sortbyweekdd" onchange="sortViewByChange();" class="sortby form-control input-small"></select>
                </th>

            </tr>
            <tr id="weektbldatetr" class="last-header">
                <th class="ddcls">
                    <select id="oliverweekdd" onchange="resourceChange()" multiple="multiple" class="oliver form-control input-small"></select>
                </th>
            </tr>
        </table>
    </script>
    <script id="monthheaderTemplate" type="text/x-jquery-tmpl">
        <table id="monthtbl" class="table dataTable table-striped table-bordered ui-datatable removetbl">
            <tr>
                <th class="ddcls" style="border-right: hidden"></th>
                <th colspan="72" class="month-header">
                    <div class="row">
                        <div class="col-md-6  monthheader-div-marg">
                            <div class="col-md-1 header-prev-margin">
                                <a href="javascript:void(0);" onclick="previous()">
                                    <i class="glyphicon glyphicon-menu-left"></i>
                                </a>
                            </div>
                            <div class="col-md-3">
                                <span id="currentDateView" class="currentdateview" data-type="month" data-month="" data-year=""></span>
                            </div>
                            <div class="col-md-1 header-next-margin">
                                <a href="javascript:void(0);" onclick="next()">
                                    <i class="glyphicon glyphicon-menu-right"></i>
                                </a>
                            </div>
                        </div>
                        <div class="col-md-2">
                            <div class="pull-right">
                                <a href="javascript:void(0);" onclick="changeDayGroup('day');" class="btn btn-default btn-sm"><b>Day</b></a>
                                <a href="javascript:void(0);" onclick="changeDayGroup('week');" class="btn btn-default btn-sm"><b>Week</b></a>
                                <a href="javascript:void(0);" onclick="changeDayGroup('month');" class="btn btn-secondary btn-sm"><b>Month</b></a>
                            </div>
                        </div>
                    </div>


                </th>
            </tr>
            <tr id="monthtbltr">
                <th class="ddcls">
                    <select id="viewtypemonthdd" onchange="viewTypeChange();" class="viewtype form-control input-small"></select>
                </th>
            </tr>
            <tr id="monthtbldaystr">
                <th class="ddcls">
                    <select id="sortbymonthdd" onchange="sortViewByChange();" class="sortby form-control input-small"></select>
                </th>
            </tr>
            <tr id="monthtbldate" class="last-header">
                <th class="ddcls">
                    <select data-style="btn-primary" id="olivermonthdd" onchange="resourceChange()" multiple="multiple" class="oliver form-control input-small"></select>
                </th>
            </tr>
        </table>
    </script>

    <script type="text/javascript" src="js/ressplan_20173.js"></script>

      
</body>
</html>
