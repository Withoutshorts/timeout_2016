

var sortbySelectedValue = '';
var viewtypeSelectedValue = '';
var oliverSelectedValue = '';
var date = new Date();
var months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
var weekDays = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
var daysview = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
var currentView = 'month';
var bookingId = 0;
var noEmployeeCalLabel = 'Debugging';
var dayStart = 8;
var dayEnd = 17;
var holidayList = [];

/*
Color codes: 
Invoiceble: #9fc6e7
Non-Invoiceble: #fbe983
Absence: #fa573c
Holiday: #b3dc6c
*/
var actTypes = {
    1: '#9fc6e7',       // Invoiceble
    2: '#fbe983',       // Non-Invoiceble
    6: '#fbe983',       // Sales
    20: '#fa573c',      // Absence
    21: '#fa573c',      // Absence
    11: '#b3dc6c',      // Holiday
    13: '#b3dc6c',      // Holiday
    14: '#b3dc6c',      // Holiday
    18: '#b3dc6c',      // Holiday
    19: '#b3dc6c'       // Holiday
};

$(document).ready(function () {
    $.ajax({
        type: "POST",
        url: "ressplan_2017.aspx/GetHolidays",
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        success: function (response) {
            holidayList = response.d;
                    
            $("#mainTemplate").tmpl().appendTo("#maindiv");
            changeDayGroup('month');

            $('#fromtime').on('change', function () {
                if ($('#dtstart').val() == $('#dtend').val() && $(this).val() > $('#totime').val()) {
                    $(this).addClass('error');
                    $(this).parent().find('p').remove()
                    $(this).parent().append('<p>Start time should be less than end time!</p>');
                }
                else {
                    $(this).removeClass('error');
                    $(this).parent().find('p').remove();
                }
            });

            $('#totime').on('change', function () {
                if ($('#dtstart').val() == $('#dtend').val() && $(this).val() < $('#fromtime').val()) {
                    $(this).addClass('error');
                    $(this).parent().find('p').remove();
                    $(this).parent().append('<p>End time should be more than start time!</p>');
                }
                else {
                    $(this).removeClass('error');
                    $(this).parent().find('p').remove();
                }
            });

            bindEmployees();
        },
    });            
});

function loadTemplate(templatename) {
    $(".removetbl").remove();
    $(templatename).tmpl().appendTo("#headerdiv");
    populateDropdown();

    $('.oliver').multiselect({
        selectAllValue: 'multiselect-all',
        //enableCaseInsensitiveFiltering: true,
        //enableFiltering: true,
        ////maxHeight: '300',
        buttonWidth: '140'
    });
}

function changeDayGroup(type) {
            
    switch (type) {
        case 'month':

            var month = $("#hidmonth").val();
            var year = $("#hidyear").val();

            if (month == "" || year == "") {
                bindMonthView(date.getMonth(), date.getFullYear());
            }
            else {
                bindMonthView(month, year);
            }

            break;

        case 'week':

            var month = $("#hidmonth").val();
            var year = $("#hidyear").val();

            if (month == "" || year == "") {
                bindWeekView(1, date.getMonth(), date.getFullYear(), 1);
            }
            else {
                var currDate = new Date();
                var todayWeekNo = 0;

                if (currDate.getMonth() == parseInt(month) && currDate.getFullYear() == parseInt(year)) {
                    todayWeekNo = new Date(currDate.getFullYear(), currDate.getMonth(), currDate.getDate()).getWeek();
                }
                else if (currentView == 'day') {
                    var currDay = $("#currentdayView").data('date');
                    var currMonth = $("#currentdayView").data('month');
                    currMonth = parseInt(currMonth) - 1;
                    var currYear = $("#currentdayView").data('year');

                    todayWeekNo = new Date(currYear, currMonth, currDay).getWeek();
                }
                else if (currDate.getMonth() != parseInt(month) || currDate.getFullYear() != parseInt(year)) {
                    todayWeekNo = new Date(year, month, 1).getWeek();
                }
                        
                if (todayWeekNo == 0) {
                    todayWeekNo = 1;
                }
                else if ($("#currentdayView").data('day') == '0') {
                    todayWeekNo--;
                }

                bindWeekView(todayWeekNo, month, year, 1);
            }                    

            break;

        case 'day':

            var month = $("#hidmonth").val();
            var year = $("#hidyear").val();
            var day = moment(year + '-' + (parseInt(month) + 1) + '-' + '01').format('d');

            var currDate = new Date();
            if (currDate.getMonth() == parseInt(month) && currDate.getFullYear() == parseInt(year)) {
                var $todayDay = $('#headerdiv .today-color');
                if ($todayDay) {
                    $todayDay.click();
                }
                else {
                    bindDayView(day, 1, parseInt(month) + 1, year, 1);
                }
            }
            else {                        
                bindDayView(day, 1, parseInt(month) + 1, year, 1);
            }
                    
                                        
            break;
    }
}

function previous() {

    var currentDateView = $(".currentdateview");
    var type = $(currentDateView).data('type');

    switch (type) {
        case 'month':

            var month = $(currentDateView).data('month');
            var year = $(currentDateView).data('year');

            if (month == 0) {
                month = 11;
                year = parseInt(year) - 1;
            } else {
                month = parseInt(month) - 1;
            }

            bindMonthView(month, year);
            break;

        case 'week':
            var week = $(currentDateView).data('week');
            var month = $(currentDateView).data('month');
            var year = $(currentDateView).data('year');
            var startdate = $(currentDateView).data('startdate');

            /* New Code Start */

            //var currentMonth = new Date(year, parseInt(month) - 1);
            //var firstDay = currentMonth.getDay();
            //// Note : Start the day with Monday
            //if (firstDay == 0) {
            //    firstDay = 6; // Set the Sunday as last day of week
            //} else {
            //    firstDay = firstDay - 1;
            //}

            //var lastDate = new Date(year, month, 0).getDate();
            //var currentWeek = 1;
            //var firstDayOfLastWeek = 1;

            //for (var i = 0; i < lastDate ; i++) {
            //    var weekDay = (i + firstDay);
            //    var weekDayIndex = weekDay % 7;
            //    if (weekDayIndex == 0 && i != 0) {
            //        currentWeek = parseInt(currentWeek) + 1;
            //        firstDayOfLastWeek = i + 1;
            //    }
            //}

            //lastweek = currentWeek;
            //firstDayOfLastWeek = firstDayOfLastWeek;

            ///* New Code End */
            //if (week == 1) {

            //    if (month == 0) {
            //        year = parseInt(year) - 1;
            //        month = 11;
            //        week = lastweek;
            //    }
            //    else {
            //        month = parseInt(month) - 1;
            //        week = lastweek;
            //    }

            //} else {
            //    week = parseInt(week) - 1;
            //}

            //if (startdate < 7) {

            //    if (week == 1) {
            //        var startdate = 1;
            //    }
            //    else {
            //        var startdate = firstDayOfLastWeek;
            //    }

            //} else {
            //    startdate = parseInt(startdate) - 7;
            //    if (startdate == 0) {
            //        startdate = 1;
            //    }
            //}

            week = parseInt(week) - 1;
            if (week == 0) {
                week = moment(year + '-' + (parseInt(month) + 1) + '-' + 1).subtract(1, 'week').format('w');
                year = parseInt(year) - 1;
                month = 11;
            }

            bindWeekView(week, month, year, startdate);
            break;

        case 'day':
            var day = $(currentDateView).data('day');
            var date = $(currentDateView).data('date');
            var month = $(currentDateView).data('month');
            var year = $(currentDateView).data('year');
            var week = $(currentDateView).data('week');
            var lastDate = new Date(year, month, 0).getDate();

            if (day == 0) {
                day = 6;
            } else {
                day = parseInt(day) - 1;
            }

            if (lastDate >= date) {
                /* For Date */
                date = parseInt(date) - 1;
                if (date == 0) {
                    var lastMonthLastDate = new Date(year, parseInt(month) - 1, 0).getDate();
                    date = lastMonthLastDate;
                    month = parseInt(month) - 1;
                }

                /* For Month */
                if (date == lastMonthLastDate && month == 0) {
                    month = 12;
                    year = parseInt(year) - 1;
                } else {
                    month = month;
                }
                        
                /* For Year */
                year = year;

            } else {
                date = parseInt(date) - 1;
                month = month;
                year = year;
            }

            bindDayView(day, date, month, year, week);

            break;

        default:
    }
}

function next() {
    var currentDateView = $(".currentdateview");
    var type = $(currentDateView).data('type');

    switch (type) {
        case 'month':
            var month = $(currentDateView).data('month');
            var year = $(currentDateView).data('year');

            if (month == 11) {
                month = 0;
                year = parseInt(year) + 1;
            }
            else {
                month = parseInt(month) + 1;
            }

            bindMonthView(month, year);
            break;

        case 'week':
            var week = $(currentDateView).data('week');
            var month = $(currentDateView).data('month');
            var year = $(currentDateView).data('year');
            var startdate = $(currentDateView).data('startdate');
            var nextstartdate = $(currentDateView).data('nextstartdate');

            week = parseInt(week) + 1;

            if(week > 52){
                var arrDatesOfWeek = getDateRangeOfWeek(week, month, year);
                var nextMonth = jQuery.grep(arrDatesOfWeek, function (a) {
                    return a.getDate() == 31 && a.getMonth() == 11;
                });

                if (nextMonth.length == 0) {
                    week = 1;
                    month = 0;
                    year = parseInt(year) + 1;
                }
            }

            //var lastDate = new Date(year, parseInt(month) + 1, 0).getDate();

            //if (nextstartdate > lastDate) {
            //    nextstartdate = 1;
            //    if (nextstartdate == 1) {
            //        week = 1;
            //        if (month == 11) {
            //            month = 0;
            //            year = parseInt(year) + 1;
            //        } else {
            //            month = parseInt(month) + 1;
            //        }
            //    }
            //    else {
            //        week = parseInt(week) + 1;
            //    }
            //} else {
            //    week = parseInt(week) + 1;
            //}                    

            bindWeekView(week, month, year, nextstartdate);
            break;

        case 'day':
            var day = $(currentDateView).data('day');
            var date = $(currentDateView).data('date');
            var month = $(currentDateView).data('month');
            var year = $(currentDateView).data('year');
            var week = $(currentDateView).data('week');
            var lastDate = new Date(year, month, 0).getDate();

            if (parseInt(day) + 1 == 7) {
                day = 0;
            } else {
                day = parseInt(day) + 1;
            }

            if (lastDate <= date) {
                /* For Date */
                date = 1;

                /* For Month */
                if (parseInt(month) + 1 == 13) {
                    month = 1;
                    /* For Year */
                    year = parseInt(year) + 1;
                } else {
                    month = parseInt(month) + 1;
                }



            } else {
                date = parseInt(date) + 1;
                month = month;
                year = year;
            }

            bindDayView(day, date, month, year, week);

            break;

        default:

    }
}

function bindMonthView(month, year) {
    currentView = 'month';
    $("#hidprecallmonth").val(month);
    $("#hidprecallyear").val(year);

    loadTemplate("#monthheaderTemplate");

    var currentMonth = new Date(year, month);
    var firstDay = currentMonth.getDay();

    // Note : Start the day with Monday
    if (firstDay == 0) {
        firstDay = 6; // Set the Sunday as last day of week
    } else {
        firstDay = firstDay - 1;
    }

    var lastDate = new Date(year, parseInt(month) + 1, 0).getDate();
    var currentWeek = 1;
    var weekOfYear = new Date(year, parseInt(month), 1).getWeek();
    var prevYear = year;

    if (new Date(year, parseInt(month), 1).getDay() == 0) {
        weekOfYear = moment(year + '-' + (parseInt(month) + 1) + '-' + 1).subtract(1, 'week').format('w')
        //prevYear = year - 1;
    }

    var weekHeader = $('<tr />');
    var dayHeader = $('<tr />');
    var dateHeader = $('<tr />');
    //var dataHeader = $('<tr />');

    $(weekHeader).append($('<th />').text('Week ' + weekOfYear).attr('colspan', (7 - firstDay)).attr('data-week-no', weekOfYear).attr('onClick', 'bindWeekView(' + weekOfYear + ',' + month + ', ' + prevYear + ', 1)').addClass("pointer center-align"));

    if (month == 0) {
        weekOfYear = 0;
    }

    for (var i = 0; i < lastDate ; i++) {
        var weekDay = (i + firstDay);
        var weekDayIndex = weekDay % 7;

        if (weekDayIndex == 0 && i != 0) {
            currentWeek = parseInt(currentWeek) + 1;
            weekOfYear++;
            var startdate = parseInt(i) + 1;
            $(weekHeader).append($('<th />').text('Week ' + weekOfYear).attr('colspan', 7).attr('data-week-no', weekOfYear).attr('onClick', 'bindWeekView(' + weekOfYear + ',' + month + ', ' + year + ', ' + startdate + ')').addClass("pointer center-align"));
        }

        var dayOfWeek = (weekDay + 1) % 7;

        var thDayHeader = $('<th />').append('<span>' + weekDays[weekDayIndex] + '</span>').attr('onClick', 'bindDayView(' + (dayOfWeek) + ',' + (i + 1) + ',' + (parseInt(month) + 1) + ', ' + year + ', ' + currentWeek + ')').addClass("pointer month-day-ths");
        var thDateHeader = $('<th />').append('<span>' + (i + 1) + '</span>').attr('onClick', 'bindDayView(' + (dayOfWeek) + ',' + (i + 1) + ',' + (parseInt(month) + 1) + ', ' + year + ', ' + currentWeek + ')').addClass("pointer month-date-ths");
                
        var todayDate = moment(year + '-' + (parseInt(month) + 1) + '-' + (i + 1));                
        if (todayDate.format('YYYY-MM-DD') == moment().format('YYYY-MM-DD')) {
            $(thDayHeader).addClass("today-color center-align").attr('data-week-no', currentWeek);
            $(thDateHeader).addClass("today-color center-align");
        }
        else if (weekDays[weekDayIndex] == "Sat" || weekDays[weekDayIndex] == "Sun") {
            $(thDayHeader).addClass("weekend-color center-align");
            $(thDateHeader).addClass("weekend-color center-align");
        }
        else {
            $.each(holidayList, function (i, e) {
                if (todayDate.format('YYYY-MM-DD') == moment(e.HolidayDate).format('YYYY-MM-DD')) {
                    $(thDayHeader).addClass("weekend-color center-align");
                    $(thDateHeader).addClass("weekend-color center-align");
                }
            });
        }

        $(dayHeader).append(thDayHeader);
        $(dateHeader).append(thDateHeader);
        //$(dataHeader).append(thDataHeader);
    }

    $('#monthtbltr th:first').after($(weekHeader).html());
    $('#monthtbldaystr th:first').after($(dayHeader).html());
    $('#monthtbldate th:first').after($(dateHeader).html());
    //$('#monthtbldata th:first').after($(dataHeader).html());

    $("#hidmonth").val(month);
    $("#hidyear").val(year);

    $('#currentDateView').html(months[month] + ' ' + year);
    $('#currentDateView').data('month', month);
    $('#currentDateView').data('year', year);
}

function bindWeekView(week, month, year, startdate)
{
    currentView = 'week';
    $("#hidprecallmonth").val(month);
    $("#hidprecallyear").val(year);            
    $('#currentweekView').data('startdate', startdate);
            
    loadTemplate("#weekheaderTemplate");
            
    var dayHeader = $('<tr />');
    var dateHeader = $('<tr />');

    var thDayHeader;
    var thDateHeader;

    var arrDatesOfWeek = getDateRangeOfWeek(week, month, year);
            
    for (var i = 0; i < arrDatesOfWeek.length; i++) {
        var dayOfWeek = i + 1;
        if (dayOfWeek > 6) {
            dayOfWeek = 0;
        }

        var tempDate = arrDatesOfWeek[i];
                
        thDayHeader = $('<th />').text(weekDays[i]).attr('onClick', 'bindDayView(' + dayOfWeek + ',' + tempDate.getDate() + ',' + (tempDate.getMonth() + 1) + ', ' + tempDate.getFullYear() + ', ' + tempDate.getWeek() + ')').addClass("pointer center-align week-day-ths");
        thDateHeader = $('<th />').text(tempDate.getDate()).attr('onClick', 'bindDayView(' + dayOfWeek + ',' + tempDate.getDate() + ',' + (tempDate.getMonth() + 1) + ', ' + tempDate.getFullYear() + ', ' + tempDate.getWeek() + ')').addClass("pointer center-align week-date-ths");
                
        if (moment(tempDate).format('YYYY-MM-DD') == moment().format('YYYY-MM-DD')) {
            $(thDayHeader).addClass("today-color center-align");
            $(thDateHeader).addClass("today-color center-align");
        }
        else if (i == 5 || i == 6) {
            $(thDayHeader).addClass("weekend-color center-align week-day-ths");
            $(thDateHeader).addClass("weekend-color center-align week-date-ths");
        }
        else {
            $.each(holidayList, function (i, e) {
                if (moment(tempDate).format('YYYY-MM-DD') == moment(e.HolidayDate).format('YYYY-MM-DD')) {
                    $(thDayHeader).addClass("weekend-color");
                    $(thDateHeader).addClass("weekend-color");
                }
            });
        }

        $(dayHeader).append(thDayHeader);
        $(dateHeader).append(thDateHeader);
    }

    var currMonth = moment(arrDatesOfWeek[0]).format('MMM') + ' ' + moment(arrDatesOfWeek[0]).format('DD') + ' - ';
    currMonth = currMonth + moment(arrDatesOfWeek[arrDatesOfWeek.length - 1]).format('MMM') + ' ' + moment(arrDatesOfWeek[arrDatesOfWeek.length - 1]).format('DD');
    currMonth = currMonth + ', ' + moment(arrDatesOfWeek[0]).format('YYYY');
    currMonth = currMonth + ' (Week ' + week + ')';
            
    $("#currentweekView").html(currMonth);

    $('#weektbldaystr th:first').after($(dayHeader).html());
    $('#weektbldatetr th:first').after($(dateHeader).html());

    $("#hidprecallstartdate").val(arrDatesOfWeek[0].getDate());
    $('#currentweekView').data('week', week);
    $('#currentweekView').data('month', month);
    $('#currentweekView').data('year', year);

    $("#hidmonth").val(arrDatesOfWeek[arrDatesOfWeek.length - 1].getMonth());
    $("#hidyear").val(arrDatesOfWeek[arrDatesOfWeek.length - 1].getFullYear());
}

function bindDayView(day, date, month, year, week) {
    currentView = 'day';
    $("#hidprecallmonth").val(month);
    $("#hidprecallyear").val(year);
            
    loadTemplate("#dayheaderTemplate");
    var appenddate = date;
    var appendmonth = month;

    if (appenddate.toString().length <= 1) {
        appenddate = "0" + appenddate;
    }
    if (appendmonth.toString().length <= 1) {
        appendmonth = "0" + appendmonth;
    }

    var firstDate = year + '-' + appendmonth + '-' + appenddate;
    var firstDay = moment(firstDate, "YYYY-MM-DD").format('dddd');

    $("#currentdayView").html(firstDay + ", " + appenddate + "-" + appendmonth + "-" + year);
    $("#currentdayView").data('day', day);
    $("#currentdayView").data('date', date);
    $("#currentdayView").data('month', month);
    $("#currentdayView").data('year', year);
    $("#currentdayView").data('week', week);

    $("#hidmonth").val(month - 1);
    $("#hidyear").val(year);
    $("#hidweek").val(week);
            
    var dateHeader = $('<tr />');

    for (var i = dayStart; i <= dayEnd; i++) {
        $(dateHeader).append($('<th />').text(i + ":00").addClass("center-align day-ths"));
    }

    //$('#timetr').html($(dateHeader).html());
    $('#timetr th:first').after($(dateHeader).html());
}

function populateDropdown() {
    /************ Sort by dropdown populate ********************/
    var sortbydd = ['SortBy', 'Customer', 'Job', 'Group', 'Employees'];
    var option = '';
    for (var i = 0; i < sortbydd.length; i++) {
        option += '<option value="' + i + '">' + sortbydd[i] + '</option>';
    }
    $('.sortby').append(option);

    if (sortbySelectedValue != "") {
        $(".sortby").val(sortbySelectedValue);
    }

    /********** View type dropdown *********************************/
    //var viewtypedd = ['View type', 'Project time', 'Absence', 'Important'];
    var viewtypedd = ['Project time'];
    var viewtypeoption = '';
    for (var i = 0; i < viewtypedd.length; i++) {
        viewtypeoption += '<option value="' + i + '">' + viewtypedd[i] + '</option>';
    }
    $('.viewtype').append(viewtypeoption);

    if (viewtypeSelectedValue != "") {
        $(".viewtype").val(viewtypeSelectedValue);
    }

    /************** oliver dropdown *************************/

    bindResources();
}

function viewTypeChange() {
    viewtypeSelectedValue = $(".viewtype").val();
    resourceChange();
}

function sortViewByChange() {
    oliverSelectedValue = '';
    bindResources();
}

// first 2 dropdown onchange event
function bindResources() {
    sortbySelectedValue = $(".sortby").val();
    viewtypeSelectedValue = $(".viewtype").val();

    if (sortbySelectedValue !== undefined) {
        $.ajax({
            type: "POST",
            url: "ressplan_2017.aspx/SortByTypeSelected",
            data: '{sortby: "' + sortbySelectedValue + '" }',
            contentType: "application/json; charset=utf-8",
            dataType: "json",
            success: function (response) {
                var olivereoption = '';
                olivereoption = response.d;
                if (olivereoption != "0") {

                    $(".oliver").empty().append(olivereoption);
                    $(".oliver").multiselect('rebuild');

                    if (oliverSelectedValue !== "" && oliverSelectedValue !== null) {
                        $('.oliver').multiselect('select', oliverSelectedValue);
                    }
                }
                else {
                    $(".oliver").empty();
                    $(".oliver").multiselect('rebuild');
                }

                resourceChange();
            },
        });
    }
}

// last dropdown onchange event and append data tr tag
function resourceChange() {
    var headerRowIndex = $('.removetbl tr.last-header').index();
    $('.removetbl tr:gt(' + headerRowIndex + ')').remove();

    oliverSelectedValue = $(".oliver").val();
            
    if (oliverSelectedValue != undefined && oliverSelectedValue != null && oliverSelectedValue.length > 0) {
        var selectedTexts = [];
        var selectedGroupId = [];

        var data = [];
        var $el = $(".oliver");
        $el.find('option:selected').each(function () {
            selectedTexts.push($(this).text());
            selectedGroupId.push($(this).val());
        });

        var month = $("#hidprecallmonth").val();
        var year = $("#hidprecallyear").val();
        var lastDate = new Date(year, parseInt(month) + 1, 0).getDate();

        var startDate = getSavedFormattedDate(new Date(year, month));
        var endDate = getSavedFormattedDate(new Date(year, parseInt(month) + 1, 0));

        var startDay = $("#hidprecallstartdate").val();
        var arrDates;

        if (currentView == 'week') {
            arrDates = getDateRangeOfWeek($('#currentweekView').data('week'), month, year);
            if (arrDates.length > 0) {
                startDate = getSavedFormattedDate(arrDates[0]);
                endDate = getSavedFormattedDate(arrDates[arrDates.length - 1]);
                lastDate = arrDates[arrDates.length - 1].getDate();
            }
        }
        else if (currentView == 'day') {
            startDay = $('#currentdayView').data('date');

            startDate = getSavedFormattedDate(new Date(year, parseInt(month) - 1, startDay));
            endDate = getSavedFormattedDate(new Date(year, parseInt(month) - 1, startDay));
        }

        if (sortbySelectedValue == 1) {
            // layout for customers
            if (currentView == 'month') {
                /********** Month view tr start (Month layout for employee data) ***********/
                $.ajax({
                    type: "POST",
                    url: "ressplan_2017.aspx/GetCustomerLayoutData",
                    data: "{'customerIds':" + JSON.stringify(selectedGroupId) + ", 'startDate':'" + startDate + "', 'endDate':'" + endDate + "', 'viewType': " + viewtypeSelectedValue + "}",
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    success: function (response) {
                        for (var j = 0; j < response.d.length; j++) {
                            var customer = response.d[j];

                            var dataMonths = $('<tr />').addClass("data-row");
                            $(dataMonths).append($('<td />').attr("colspan", parseInt(lastDate) + 1).append('<b>' + customer.CustomerName + '</b><br/>' + customer.JobName).addClass(""));
                            $('#monthtbl tbody').append($(dataMonths));

                            var activities = customer.Activities;

                            for (var index = 0; index < activities.length; index++) {
                                dataMonths = $('<tr />').addClass("data-row");
                                        
                                $(dataMonths).append($('<td class="v-middle" />').text(activities[index].ActivityName));
                                bindMonthViewDataCustomerJob(lastDate, year, month, activities[index], $(dataMonths));

                                $('#monthtbl tbody').append($(dataMonths));
                            }
                        }

                        arrangeMonthViewData();
                    }
                });
            }
            else if (currentView == 'week') {
                /********** week view tr start (Week layout for employee data) ***********/
                $.ajax({
                    type: "POST",
                    url: "ressplan_2017.aspx/GetCustomerLayoutData",
                    data: "{'customerIds':" + JSON.stringify(selectedGroupId) + ", 'startDate':'" + startDate + "', 'endDate':'" + endDate + "', 'viewType': " + viewtypeSelectedValue + "}",
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    success: function (response) {
                        for (var j = 0; j < response.d.length; j++) {
                            var customer = response.d[j];

                            var dataWeeks = $('<tr />').addClass("data-row");
                            $(dataWeeks).append($('<td class="v-middle" />').attr("colspan", 8).append('<b>' + customer.CustomerName + '</b><br/>' + customer.JobName));
                            $('#weektbl tbody').append($(dataWeeks));

                            var activities = customer.Activities;

                            for (var index = 0; index < activities.length; index++) {
                                dataWeeks = $('<tr />').addClass("data-row");
                                $(dataWeeks).append($('<td class="v-middle" />').text(activities[index].ActivityName));
                                bindWeekViewDataCustomerJob(arrDates, activities[index], $(dataWeeks));
                                $('#weektbl tbody').append($(dataWeeks));
                            }
                        }

                        arrangeWeekViewData();
                    }
                });
            }
            else if (currentView == 'day') {
                /********** day view tr start (Day layout for employee data) ***********/
                $.ajax({
                    type: "POST",
                    url: "ressplan_2017.aspx/GetCustomerLayoutData",
                    data: "{'customerIds':" + JSON.stringify(selectedGroupId) + ", 'startDate':'" + startDate + "', 'endDate':'" + endDate + "', 'viewType': " + viewtypeSelectedValue + "}",
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    success: function (response) {
                        for (var j = 0; j < response.d.length; j++) {
                            var customer = response.d[j];

                            var dataDays = $('<tr />').addClass("data-row");
                            $(dataDays).append($('<td class="v-middle" />').attr("colspan", 18).append('<b>' + customer.CustomerName + '</b><br/>' + customer.JobName));
                            $('#daytbl tbody').append($(dataDays));
                                    
                            var activities = customer.Activities;

                            for (var index = 0; index < activities.length; index++) {
                                dataDays = $('<tr />').addClass("data-row");
                                $(dataDays).append($('<td class="v-middle" />').text(activities[index].ActivityName));
                                bindDayViewDataCustomerJob(startDate, activities[index], $(dataDays));
                                $('#daytbl tbody').append($(dataDays));
                            }
                        }

                        arrangeDayViewData();
                    }
                });
            }
        }
        else if (sortbySelectedValue == 2) {
            //layout for job

                 

            if (currentView == 'month') {
                /************ Month View *************/

             

                $.ajax({
                    type: "POST",
                    url: "ressplan_2017.aspx/GetJobLayoutData",
                    data: "{'jobIds':" + JSON.stringify(selectedGroupId) + ", 'startDate':'" + startDate + "', 'endDate':'" + endDate + "', 'viewType': " + viewtypeSelectedValue + "}",
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    success: function (response) {
                        for (var j = 0; j < response.d.length; j++) {
                            var customer = response.d[j];
                         
                            var dataMonths = $('<tr />').addClass("data-row");
                            $(dataMonths).append($('<td />').attr("colspan", parseInt(lastDate) + 1).append('<b>' + customer.CustomerName + '</b><br/>' + customer.JobName).addClass(""));
                            //$(dataMonths).append($('<td />').attr("colspan", 8).append('<b>Kundenavn</b><br/>Jobnavn'));
                            $('#monthtbl tbody').append($(dataMonths));

                            var activities = customer.Activities;
                          
                            for (var index = 0; index < activities.length; index++) {
                                dataMonths = $('<tr />').addClass("data-row");

                                $(dataMonths).append($('<td class="v-middle" />').text(activities[index].ActivityName));
                                bindMonthViewDataCustomerJob(lastDate, year, month, activities[index], $(dataMonths));

                                //alert("HER 2 - " + index + "lastDate: " + lastDate + ", year:" + year + ", month:" + month + ", act: " + activities[index].ActivityName);

                                $('#monthtbl tbody').append($(dataMonths));
                            } 
                        }

                        arrangeMonthViewData();
                       
                        
                    },
                });

               
            }
            else if (currentView == 'week') {
                /************ Week view **************/
                $.ajax({
                    type: "POST",
                    url: "ressplan_2017.aspx/GetJobLayoutData",
                    data: "{'jobIds':" + JSON.stringify(selectedGroupId) + ", 'startDate':'" + startDate + "', 'endDate':'" + endDate + "', 'viewType': " + viewtypeSelectedValue + "}",
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    success: function (response) {
                        for (var j = 0; j < response.d.length; j++) {
                            var customer = response.d[j];

                            var dataWeeks = $('<tr />').addClass("data-row");
                            $(dataWeeks).append($('<td class="v-middle" />').attr("colspan", 8).append('<b>' + customer.CustomerName + '</b><br/>' + customer.JobName));
                            $('#weektbl tbody').append($(dataWeeks));

                            var activities = customer.Activities;

                            for (var index = 0; index < activities.length; index++) {
                                dataWeeks = $('<tr />').addClass("data-row");
                                $(dataWeeks).append($('<td class="v-middle" />').text(activities[index].ActivityName));
                                //alert("HER W - " + index + " act: " + activities[index].ActivityName);
                                bindWeekViewDataCustomerJob(arrDates, activities[index], $(dataWeeks));
                                $('#weektbl tbody').append($(dataWeeks));
                            }
                        }

                        arrangeWeekViewData();
                    }
                });
            }
            else if (currentView == 'day') {
                /************ day view **************/
                $.ajax({
                    type: "POST",
                    url: "ressplan_2017.aspx/GetJobLayoutData",
                    data: "{'jobIds':" + JSON.stringify(selectedGroupId) + ", 'startDate':'" + startDate + "', 'endDate':'" + endDate + "', 'viewType': " + viewtypeSelectedValue + "}",
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    success: function (response) {
                        for (var j = 0; j < response.d.length; j++) {
                            var customer = response.d[j];

                            var dataDays = $('<tr />').addClass("data-row");
                            $(dataDays).append($('<td class="v-middle" />').attr("colspan", 18).append('<b>' + customer.CustomerName + '</b><br/>' + customer.JobName));
                            $('#daytbl tbody').append($(dataDays));

                            var activities = customer.Activities;

                            for (var index = 0; index < activities.length; index++) {
                                dataDays = $('<tr />').addClass("data-row");
                                $(dataDays).append($('<td class="v-middle" />').text(activities[index].ActivityName));
                                bindDayViewDataCustomerJob(startDate, activities[index], $(dataDays));
                                $('#daytbl tbody').append($(dataDays));
                            }
                        }

                        arrangeDayViewData();
                    }
                });
            }
        }
        else if (sortbySelectedValue == 3) {
            //layout for group
            if (currentView == 'month') {
                /************ Month View *************/
                $.ajax({
                    type: "POST",
                    url: "ressplan_2017.aspx/GetGroupLayoutData",
                    data: "{'groupIds':" + JSON.stringify(selectedGroupId) + ", 'startDate':'" + startDate + "', 'endDate':'" + endDate + "', 'viewType': " + viewtypeSelectedValue + "}",
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    success: function (response) {

                        for (var j = 0; j < selectedTexts.length; j++) {

                            var dataMonths = $('<tr />').addClass("data-row");
                            $(dataMonths).append($('<td class="v-middle" />').attr("colspan", parseInt(lastDate) + 1).append('<b>' + selectedTexts[j] + '</b>'));
                            $('#monthtbl tbody').append($(dataMonths));

                            var arrEmployees = $.grep(response.d, function (v) {
                                return v.ProjectId === parseInt(selectedGroupId[j]);
                            });
                                    
                            for (var p = 0; p < arrEmployees.length; p++) {
                                var employee = arrEmployees[p];
                                var activities = employee.Activities;

                                dataMonths = $('<tr />').addClass("data-row");

                                $(dataMonths).append($('<td class="v-middle" />').text(employee.Name));
                                bindMonthViewData(lastDate, year, month, activities, $(dataMonths));

                                $('#monthtbl tbody').append($(dataMonths));
                            }
                        }

                        arrangeMonthViewData();
                    },
                });
            }
            else if (currentView == 'week') {
                /************ Week view **************/
                $.ajax({
                    type: "POST",
                    url: "ressplan_2017.aspx/GetGroupLayoutData",
                    data: "{'groupIds':" + JSON.stringify(selectedGroupId) + ", 'startDate':'" + startDate + "', 'endDate':'" + endDate + "', 'viewType': " + viewtypeSelectedValue + "}",
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    success: function (response) {

                        for (var j = 0; j < selectedTexts.length; j++) {

                            var dataWeeks = $('<tr />').addClass("data-row");
                            $(dataWeeks).append($('<td class="v-middle" />').attr("colspan", 8).append('<b>' + selectedTexts[j] + '</b>').addClass(""));
                            $('#weektbl tbody').append($(dataWeeks));
                                    
                            var arrEmployees = $.grep(response.d, function (v) {
                                return v.ProjectId === parseInt(selectedGroupId[j]);
                            });

                            for (var p = 0; p < arrEmployees.length; p++) {
                                var employee = arrEmployees[p];
                                var activities = employee.Activities;
                                dataWeeks = $('<tr />').addClass("data-row");                                        
                                $(dataWeeks).append($('<td class="v-middle" />').text(employee.Name));
                                bindWeekViewData(arrDates, activities, $(dataWeeks));
                                $('#weektbl tbody').append($(dataWeeks));
                            }
                        }

                        arrangeWeekViewData();
                    }
                });
            }
            else if (currentView == 'day') {
                /************ day view **************/
                $.ajax({
                    type: "POST",
                    url: "ressplan_2017.aspx/GetGroupLayoutData",
                    data: "{'groupIds':" + JSON.stringify(selectedGroupId) + ", 'startDate':'" + startDate + "', 'endDate':'" + endDate + "', 'viewType': " + viewtypeSelectedValue + "}",
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    success: function (response) {
                        for (var j = 0; j < selectedTexts.length; j++) {

                            var dataDays = $('<tr />').addClass("data-row");
                            $(dataDays).append($('<td class="v-middle" />').attr("colspan", 18).append('<b>' + selectedTexts[j] + '</b>').addClass(""));
                            $('#daytbl tbody').append($(dataDays));
                            var weekHeaderth = $('tr#weektbldaystr th.week-day-ths');
                            var weekHeaderDateth = $('tr#weektbldatetr th.week-date-ths');

                            var arrEmployees = $.grep(response.d, function (v) {
                                return v.ProjectId === parseInt(selectedGroupId[j]);
                            });

                            for (var p = 0; p < arrEmployees.length; p++) {
                                var employee = arrEmployees[p];
                                var activities = employee.Activities;
                                dataDays = $('<tr />').addClass("data-row");
                                $(dataDays).append($('<td class="v-middle" />').text(employee.Name));
                                        
                                bindDayViewData(startDate, activities, $(dataDays));

                                $('#daytbl tbody').append($(dataDays));
                            }
                        }

                        arrangeDayViewData();
                    }
                });
            }
        }
        else if (sortbySelectedValue == 4) {
            // layout for employees
            if (currentView == 'month') {
                /********** Month view tr start (Month layout for employee data) ***********/
                $.ajax({
                    type: "POST",
                    url: "ressplan_2017.aspx/GetEmployeeLayoutData",
                    data: "{'employeeIds':" + JSON.stringify(selectedGroupId) + ", 'startDate':'" + startDate + "', 'endDate':'" + endDate + "', 'viewType': " + viewtypeSelectedValue + "}",
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    success: function (response) {

                        for (var j = 0; j < response.d.length; j++) {
                            var employee = response.d[j];
                            var activities = employee.Activities;
                            var dataMonths = $('<tr />').addClass("data-row");
                            $(dataMonths).append($('<td class="v-middle" />').text(employee.Name));
                            bindMonthViewData(lastDate, year, month, activities, $(dataMonths));

                            $('#monthtbl tbody').append($(dataMonths));
                        }

                        arrangeMonthViewData();
                    }
                });
            }
            else if (currentView == 'week') {
                /********** week view tr start (Week layout for employee data) ***********/
                $.ajax({
                    type: "POST",
                    url: "ressplan_2017.aspx/GetEmployeeLayoutData",
                    data: "{'employeeIds':" + JSON.stringify(selectedGroupId) + ", 'startDate':'" + startDate + "', 'endDate':'" + endDate + "', 'viewType': " + viewtypeSelectedValue + "}",
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    success: function (response) {
                        for (var j = 0; j < response.d.length; j++) {
                            var employee = response.d[j];
                            var activities = employee.Activities;

                            var dataWeeks = $('<tr />').addClass("data-row");
                            $(dataWeeks).append($('<td class="v-middle" />').text(employee.Name));
                            bindWeekViewData(arrDates, activities, $(dataWeeks));
                            $('#weektbl tbody').append($(dataWeeks));
                        }

                        arrangeWeekViewData();
                    }
                });
            }
            else if (currentView == 'day') {
                /********** day view tr start (Day layout for employee data) ***********/
                $.ajax({
                    type: "POST",
                    url: "ressplan_2017.aspx/GetEmployeeLayoutData",
                    data: "{'employeeIds':" + JSON.stringify(selectedGroupId) + ", 'startDate':'" + startDate + "', 'endDate':'" + endDate + "', 'viewType': " + viewtypeSelectedValue + "}",
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    success: function (response) {
                        for (var j = 0; j < response.d.length; j++) {
                            var employee = response.d[j];
                            var activities = employee.Activities;

                            var dataDays = $('<tr />').addClass("data-row");
                            $(dataDays).append($('<td class="v-middle" />').text(employee.Name));

                            bindDayViewData(startDate, activities, $(dataDays));
                            $('#daytbl tbody').append($(dataDays));
                        }

                        arrangeDayViewData();
                    }
                });
            }
        }
    }
}

function editActivity(id) {
    bookingId = id;

    if (sortbySelectedValue == '1' || sortbySelectedValue == '2') {
        $('#headtxt').prop('disabled', 'disabled');
        $('#empchk').attr('multiple', 'multiple');
    }
    else {
        $('#headtxt').prop('disabled', false);
        $('#empchk').removeAttr('multiple');
    }

    //Get Employee options
    $.ajax({
        type: "POST",
        url: "ressplan_2017.aspx/EditActivity",
        data: "{'bookingId':" + bookingId + "}",
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        success: function (response) {

            $('#fromtime').val(response.d.StartTime);
            $('#totime').val(response.d.EndTime);

            $('#dtstart').val(getDisplayFormattedDate(response.d.StartDate));
            $('#dtstart').attr('date', getSavedFormattedDate(response.d.StartDate));

            $('#dtend').val(getDisplayFormattedDate(response.d.EndDate));
            $('#dtend').attr('date', getSavedFormattedDate(response.d.EndDate));

            $('#recurrencedd').val(response.d.Recurrence);
            if (response.d.Important == 1) {
                $('#chkimp').prop('checked', true);
            }
            else {
                $('#chkimp').prop('checked', false);
            }

            recurrenceChange();

            if ($('#recurrencedd').val() != '0' && $('#recurrencedd').val() != '-1') {
                $('#noOfRecurrence').val(response.d.NoOfRecurrence);
            }

            var employees = [];
            $.each(response.d.EmployeeList, function (i, e) {
                employees.push(e.Id);
            });

            $('#empchk').val(employees);
            $('#custtxt').val(response.d.Customer);
            $('#headtxt').val(response.d.Heading);

            $('#projecttxt').val(response.d.Project);

            var activityType = response.d.ActivityType;
            var selectedActType = 'Invoiceable';
            if (activityType == 2) {
                selectedActType = 'None Invoiceable';
            }
            else if (activityType == 20 || activityType == 21) {
                selectedActType = 'Holiday';
            }
            else if (activityType == 11 || activityType == 13 || activityType == 14 || activityType == 18 || activityType == 19) {
                selectedActType = 'Illness';
            }

            $('#actType').val(selectedActType);                    
            $('#prodtxt').val(response.d.ActivityName);
        },
    });

    $('#activityModal').modal('show');

    $('#activityModal').on('shown.bs.modal', function () {
        afterModelShown();
    });
}

//bind datapicker after modal shown
function afterModelShown() {
            
    $('.input-group.date.start').datepicker({
        format: 'dd-mm-yyyy',
        autoclose: true
    }).on('changeDate', function (e) {
        compareDate(e.date, $(this));
    });

    $('.input-group.date.end').datepicker({
        format: 'dd-mm-yyyy',
        autoclose: true
    }).on('changeDate', function (e) {
        compareDate(e.date, $(this));
    });
}

function compareDate(selectedDate, cntrl) {
    $(cntrl).find('input:text').attr('date', getSavedFormattedDate(selectedDate));

    var datestartObj = new Date($('.input-group.date.start').datepicker('getDate'));
    var momentStartObj = moment(datestartObj);
    var momentStartString = momentStartObj.format('YYYY-MM-DD');

    var dateEndObj = new Date($('.input-group.date.end').datepicker('getDate'));
    var momentEndObj = moment(dateEndObj);
    var momentEndString = momentEndObj.format('YYYY-MM-DD')

    var $endDateCntrl = $('.input-group.date.end');
    if (momentStartObj._d > momentEndObj._d) {
        $endDateCntrl.children('input').addClass('error');
        $endDateCntrl.parent().find('p').show();
    }
    else {
        $endDateCntrl.children('input').removeClass('error');
        $endDateCntrl.parent().find('p').hide();
    }
}

function timeValidation() {
    if ($('#dtstart').val() == $('#dtend').val() && $('#fromtime').val() > $('#totime').val()) {
        $('#totime').addClass('error');
        $('#totime').parent().append('<p class="error-txt">End time should less than start time.</p>');
    }
    else {
        if ($('#totime').hasClass('error')) {
            $('#totime').removeClass('error');
        }
    }
}

function saveActivity() {

   
    timeValidation();

    var taskData = JSON.stringify($('form').serializeArray());

    var model = {};
    model.BookingId = bookingId;
    
    model.StartDate = $('#dtstart').attr('date'); //'2017-04-02' //
    model.EndDate = $('#dtend').attr('date');
    model.StartTime = $('#fromtime').val();
    model.StartDateTime = model.StartDate + ' ' + model.StartTime;

    model.EndTime = $('#totime').val();
    model.EndDateTime = model.EndDate + ' ' + model.EndTime;

    if (sortbySelectedValue !== '1' && sortbySelectedValue !== '2') {
        var arrEmployees = [];
        arrEmployees.push($('#empchk').val());
        model.Employees = arrEmployees;
    }
    else {
        model.Employees = $('#empchk').val();
    }

    //new Date(model.StartDate + ' ' + model.StartTime);
    /*
    model.StartTime = $('#fromtime').val();
    model.EndTime = $('#totime').val();
    model.StartDateTime = new Date(model.StartDate + ' ' + model.StartTime);
    model.EndDateTime = new Date(model.EndDate + ' ' + model.EndTime);
    model.Recurrence = $('#recurrencedd').val();
    //alert ("model" + model.BookingId + model.Recurrence)
    if ($('#recurrencedd').val() != '0' && $('#recurrencedd').val() != '-1') {
        model.NoOfRecurrence = $('#noOfRecurrence').val();
    }
    else {
        model.NoOfRecurrence = 0;
    }

    model.Important = $('#chkimp').prop('checked') ? 1 : 0;

    if (sortbySelectedValue !== '1' && sortbySelectedValue !== '2') {
        var arrEmployees = [];
        arrEmployees.push($('#empchk').val());
        model.Employees = arrEmployees;
    }
    else {
        model.Employees = $('#empchk').val();
    }
            
    model.Heading = $('#headtxt').val() == null ? 0 : $('#headtxt').val();
    */

    // data: JSON.stringify({ 'model': model }),
    // data: "{'model':" + model + "}",
    if (!$('#activityModal').find('div.col-lg-4 input').hasClass('error')) {
        $.ajax({
            type: "POST",
            url: "ressplan_2017.aspx/SaveActivity",
            data: JSON.stringify({ 'model': model }),
            contentType: "application/json; charset=utf-8",
            dataType: "json",
            success: function (response) {
                if (response.d > 0) {
                    alert('Activity Updated successfully.');
                    bookingId = 0;
                    $('#activityModal').modal('hide');
                    resourceChange();
                }
                else {
                    alert('Error! while updating activity.');
                    //alert("Test");
                }
            },
        }); //alert("HER")
    }
}

function bindEmployees() {            
    //Get Employee options
    $.ajax({
        type: "POST",
        url: "ressplan_2017.aspx/GetAllEmployees",
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        success: function (response) {
            $('#empchk').empty();
            $(response.d).appendTo('#empchk');
        },
    });
}

function getDisplayFormattedDate(date) {
    var custDate = new Date(date);
    var day = custDate.getDate();
    var month = custDate.getMonth();
    month += 1;

    if (day < 10) {
        day = '0' + day;
    }

    if (month < 10) {
        month = '0' + month;
    }

    var year = custDate.getFullYear();

    var customDate = day + "-" + month + "-" + year;
    return customDate;
}

function getSavedFormattedDate(date) {
    var custDate = new Date(date);
    var day = custDate.getDate();
    var month = custDate.getMonth();
    month += 1;

    if (day < 10) {
        day = '0' + day;
    }

    if (month < 10) {
        month = '0' + month;
    }

    var year = custDate.getFullYear();

    var customDate = year + "-" + month + "-" + day;
    return customDate;
}

function bindMonthViewData(lastDate, year, month, activities, $element) {

    //alert("HER" + "lastDate: " + lastDate + "year: " + year + "month: " + month + "activities: "+ activities)
    var monthHeaderth = $('tr#monthtbldate th.month-date-ths');

    for (var i = 0; i < lastDate; i++) {

        var currDate = year //+ '-' + (parseInt(month) + 1) + '-' + (i + 1);

        //alert(currDate)
        if (parseInt(month) + 1 < 10) {
            currDate += '-0' + (parseInt(month) + 1);
        } else {
            currDate += '-' + (parseInt(month) + 1);
        }

        if ((i + 1) < 10) {
            currDate += '-0' + (i + 1);
        } else {
            currDate += '-' + (i + 1);
        }


        //alert("i:" + i)
        //alert("currDate: " + currDate)
        //alert("booking St. og endtime: " + v.startDateTime + " " + v.endDateTime + " currDate: " + currDate)


        //if (!$(monthHeaderth[i]).hasClass('weekend-color')) {
        var arrActivities = $.grep(activities, function (v) {
           //alert("booking St. og endtime: " + moment(v.StartDateTime) + " " + moment(v.EndDateTime) + " currDate: " + currDate)

            return moment(moment(v.EndDateTime).format('YYYY-MM-DD')) >= moment(currDate) && moment(moment(v.StartDateTime).format('YYYY-MM-DD')) <= moment(currDate);
            //return moment(v.EndDateTime) >= moment(currDate) && moment(moment(v.StartDateTime)) <= moment(currDate);
        });
                    
        arrActivities.sort(function (x, y) {
            return x.ActivityName.localeCompare(y.ActivityName);
        });

        var dataHTML = '';
                    
        for (var k = 0; k < arrActivities.length; k++) {
            var backColor = actTypes[arrActivities[k].ActivityType];
            var activityTime = moment(arrActivities[k].EndDateTime).diff(arrActivities[k].StartDateTime, 'minutes');
            var widthClass = 'month-full-width';
            if (activityTime <= 240) {
                widthClass = 'month-half-width';
            }

            dataHTML += '<div class="main-ac-data ' + widthClass + ' pointer" style="background:' + backColor + ';" onclick="editActivity(' + arrActivities[k].BookingId + ')">';

            var heading = arrActivities[k].ActivityName;
            if (arrActivities[k].Heading == 2) {
                heading = arrActivities[k].JobName;
            }
            else if (arrActivities[k].Heading == 3) {
                heading = arrActivities[k].CustomerName;
            }

            dataHTML += '<div class="ac-name">' + heading + '</div>';

            if (arrActivities[k].Important == 1 || arrActivities[k].Recurrence == 0) {
                dataHTML += '<div class="ac-icon">';

                if (arrActivities[k].Important == 1) {
                    dataHTML += '<a class="imp" style="color:' + backColor + ';" title="Important">!</a>';
                }

                if (arrActivities[k].Recurrence == 0) {
                    dataHTML += '<a class="glyphicon glyphicon-refresh" title="Daily"></a>';
                }

                dataHTML += '</div>';
            }

            dataHTML += '</div>';
        }

        var todayClass = '';
        if ($(monthHeaderth[i]).hasClass('today-color')) {
            todayClass = 'today-color';
        }
        else if ($(monthHeaderth[i]).hasClass('weekend-color')) {
            todayClass = 'weekend-color';
        }

        $element.append($('<td />').append(dataHTML).addClass(todayClass));
        //}
        //else {
        //    $element.append($('<td />').addClass('weekend-color'));
        //}
    }
}

function bindWeekViewData(arrDates, activities, $element) {
    var weekHeaderth = $('tr#weektbldaystr th.week-day-ths');
    var weekHeaderDateth = $('tr#weektbldatetr th.week-date-ths');
    startDay = $("#hidprecallstartdate").val();

    for (var i = 0; i < arrDates.length ; i++) {
        var todayClass = '';
        if ($(weekHeaderth[i]).hasClass('today-color')) {
            todayClass = 'today-color';
        }
        else if ($(weekHeaderth[i]).hasClass('weekend-color')) {
            todayClass = 'weekend-color';
        }

        //if (!$(weekHeaderth[i]).hasClass('weekend-color')) {
        var arrActivities = $.grep(activities, function (v) {
            return moment(moment(v.EndDateTime).format('YYYY-MM-DD')) >= moment(arrDates[i]) && moment(moment(v.StartDateTime).format('YYYY-MM-DD')) <= moment(arrDates[i]);
        });

        arrActivities.sort(function (x, y) {
            return x.ActivityName.localeCompare(y.ActivityName);
        });

        var dataHTML = '';

        for (var k = 0; k < arrActivities.length; k++) {
            var backColor = actTypes[arrActivities[k].ActivityType];
            var activityTime = moment(arrActivities[k].EndDateTime).diff(arrActivities[k].StartDateTime, 'minutes');
            var widthClass = 'week-full-width';
            if (activityTime <= 120) {
                widthClass = 'week-half-width';
            }

            dataHTML += '<div class="main-ac-data ' + widthClass + ' pointer" style="background:' + backColor + ';" onclick="editActivity(' + arrActivities[k].BookingId + ')">';

            var heading = arrActivities[k].ActivityName;
            if (arrActivities[k].Heading == 2) {
                heading = arrActivities[k].JobName;
            }
            else if (arrActivities[k].Heading == 3) {
                heading = arrActivities[k].CustomerName;
            }

            dataHTML += '<div class="ac-name">' + heading + '</div>';
                        
            if (arrActivities[k].Important == 1 || arrActivities[k].Recurrence == 0) {
                dataHTML += '<div class="ac-icon">';

                if (arrActivities[k].Important == 1) {
                    dataHTML += '<a class="imp" style="color:' + backColor + ';" title="Important">!</a>';
                }

                if (arrActivities[k].Recurrence == 0) {
                    dataHTML += '<a class="glyphicon glyphicon-refresh" title="Daily"></a>';
                }

                dataHTML += '</div>';
            }

            dataHTML += '</div>';
        }

        $element.append($('<td />').append(dataHTML).addClass("week-data-td " + todayClass));
        //}
        //else if ($(weekHeaderth[i]).hasClass('weekend-color')) {
        //    $element.append($('<td />').addClass("week-data-td weekend-color"));
        //}
        //else {
        //    $element.append($('<td />').addClass("week-data-td " + todayClass));
        //}

        startDay++;
    }
}

function bindDayViewData(startDate, activities, $element) {
    var currentDateView = $(".currentdateview");
    var day = $(currentDateView).data('day');
    var isHoliday = false;
    $.each(holidayList, function (i, e) {
        if (moment(startDate).format('YYYY-MM-DD') == moment(e.HolidayDate).format('YYYY-MM-DD')) {
            isHoliday = true;
        }
    });

    if (!isHoliday) {
        if (parseInt(day) == 0 || parseInt(day) == 6) {
            isHoliday = true;
        }
    }

    for (var i = dayStart; i <= dayEnd; i++) {
        //if (parseInt(day) !== 0 && parseInt(day) !== 6) {
        var currTime = i + '00:00';
        if (i < 10) {
            currTime = '0' + currTime;
        }

        var arrActivities = $.grep(activities, function (v) {
            return moment(moment(v.EndDateTime).format('YYYY-MM-DD')) >= moment(startDate) && moment(moment(v.StartDateTime).format('YYYY-MM-DD')) <= moment(startDate) &&
                    ((moment(v.StartDateTime).format('HH:mm:ss') == '00:00:00' && moment(v.EndDateTime).format('HH:mm:ss') == '00:00:00') ||
                        (moment(moment(v.StartDateTime).format('HH:mm:ss'), 'HH:mm:ss') <= moment(currTime, "HH:mm:ss") &&
                         moment(moment(v.EndDateTime).format('HH:mm:ss'), 'HH:mm:ss') >= moment(currTime, "HH:mm:ss")));
        });

        arrActivities.sort(function (x, y) {
            return x.ActivityName.localeCompare(y.ActivityName);
        });

        var dataHTML = '';

        for (var k = 0; k < arrActivities.length; k++) {
            var backColor = actTypes[arrActivities[k].ActivityType];
            var activityTime = moment(arrActivities[k].EndDateTime).diff(arrActivities[k].StartDateTime, 'minutes');
            var widthClass = 'day-full-width';
            if (activityTime <= 120) {
                widthClass = 'day-half-width';
            }

            dataHTML += '<div class="main-ac-data ' + widthClass + ' pointer" style="background:' + backColor + ';" onclick="editActivity(' + arrActivities[k].BookingId + ')">';

            var heading = arrActivities[k].ActivityName;
            if (arrActivities[k].Heading == 2) {
                heading = arrActivities[k].JobName;
            }
            else if (arrActivities[k].Heading == 3) {
                heading = arrActivities[k].CustomerName;
            }

            dataHTML += '<div class="ac-name">' + heading + '</div>';

            if (arrActivities[k].Important == 1 || arrActivities[k].Recurrence == 0) {
                dataHTML += '<div class="ac-icon">';

                if (arrActivities[k].Important == 1) {
                    dataHTML += '<a class="imp" style="color:' + backColor + ';" title="Important">!</a>';
                }

                if (arrActivities[k].Recurrence == 0) {
                    dataHTML += '<a class="glyphicon glyphicon-refresh" title="Daily"></a>';
                }

                dataHTML += '</div>';
            }

            dataHTML += '</div>';
        }

        var todayClass = '';
        if (startDate == moment().format('YYYY-MM-DD')) {
            todayClass = 'today-color';
        }
        else if (isHoliday) {
            todayClass = 'weekend-color';
        }

        $element.append($('<td />').append(dataHTML).addClass("day-data-td " + todayClass));
        //}
        //else {
        //    $element.append($('<td />').addClass("weekend-color"));
        //}
    }
}

function bindMonthViewDataCustomerJob(lastDate, year, month, activity, $element) {
    var monthHeaderth = $('tr#monthtbldate th.month-date-ths');
            
    //alert("HEJ 2")

    for (var i = 0; i < lastDate; i++) {

        //alert("HEJ")
        //var currDate = year + '-' + (parseInt(month) + 1) + '-' + (i + 1);
        //alert("HEJ 3" + currDate)

        var currDate = year;

        if (parseInt(month) + 1 < 10) {
            currDate += '-0' + (parseInt(month) + 1);
        } else {
            currDate += '-' + (parseInt(month) + 1);
        }

        if ((i + 1) < 10) {
            currDate += '-0' + (i + 1);
        } else {
            currDate += '-' + (i + 1);
        }
        

        //alert(currDate)

        

        //if (!$(monthHeaderth[i]).hasClass('weekend-color')) {
        var dataHTML = '';

        
        //alert(moment(moment(activity.EndDateTime).format('YYYY-MM-DD')) + " >= " + moment(currDate) + " Stdate: " + moment(activity.StartDateTime));

        if (moment(moment(activity.EndDateTime).format('YYYY-MM-DD')) >= moment(currDate) && moment(moment(activity.StartDateTime).format('YYYY-MM-DD')) <= moment(currDate)) {
            for (var k = 0; k < activity.EmployeeList.length; k++) {
                var backColor = actTypes[activity.ActivityType];
                var activityTime = moment(activity.EndDateTime).diff(activity.StartDateTime, 'minutes');
                var widthClass = 'month-full-width';
                if (activityTime <= 240) {
                    widthClass = 'month-half-width';
                }

                dataHTML += '<div class="main-ac-data ' + widthClass + ' pointer" style="background:' + backColor + ';" onclick="editActivity(' + activity.BookingId + ')">';
                var employeeName = activity.EmployeeList[k].Name;
                if (employeeName == '') {
                    employeeName = noEmployeeCalLabel;
                }
                dataHTML += '<div class="ac-name">' + employeeName + '</div>';

                if (activity.Important == 1 || activity.Recurrence == 0) {
                    dataHTML += '<div class="ac-icon">';

                    if (activity.Important == 1) {
                        dataHTML += '<a class="imp" style="color:' + backColor + ';" title="Important">!</a>';
                    }

                    if (activity.Recurrence == 0) {
                        dataHTML += '<a class="glyphicon glyphicon-refresh" title="Daily"></a>';
                    }

                    dataHTML += '</div>';
                }

                dataHTML += '</div>';
            }
        }
                    
        var todayClass = '';
        if ($(monthHeaderth[i]).hasClass('today-color')) {
            todayClass = 'today-color';
        }
        else if ($(monthHeaderth[i]).hasClass('weekend-color')) {
            todayClass = 'weekend-color';
        }

        $element.append($('<td />').append(dataHTML).addClass(todayClass));
        //}
        //else {
        //$element.append($('<td />').addClass('weekend-color'));
        //}
    }
}

function bindWeekViewDataCustomerJob(arrDates, activity, $element) {
    var weekHeaderth = $('tr#weektbldaystr th.week-day-ths');
            
    for (var i = 0; i < arrDates.length ; i++) {
        var todayClass = '';
        if ($(weekHeaderth[i]).hasClass('today-color')) {
            todayClass = 'today-color';
        }
        else if ($(weekHeaderth[i]).hasClass('weekend-color')) {
            todayClass = 'weekend-color';
        }

        //if (!$(weekHeaderth[i]).hasClass('weekend-color')) {
        var dataHTML = '';

        if (moment(moment(activity.EndDateTime).format('YYYY-MM-DD')) >= moment(arrDates[i]) && moment(moment(activity.StartDateTime).format('YYYY-MM-DD')) <= moment(arrDates[i])) {
            for (var k = 0; k < activity.EmployeeList.length; k++) {
                var backColor = actTypes[activity.ActivityType];
                var activityTime = moment(activity.EndDateTime).diff(activity.StartDateTime, 'minutes');
                var widthClass = 'week-full-width';
                if (activityTime <= 120) {
                    widthClass = 'week-half-width';
                }

                dataHTML += '<div class="main-ac-data ' + widthClass + ' pointer" style="background:' + backColor + ';" onclick="editActivity(' + activity.BookingId + ')">';

                var employeeName = activity.EmployeeList[k].Name;
                if (employeeName == '') {
                    employeeName = noEmployeeCalLabel;
                }
                dataHTML += '<div class="ac-name">' + employeeName + '</div>';

                if (activity.Important == 1 || activity.Recurrence == 0) {
                    dataHTML += '<div class="ac-icon">';

                    if (activity.Important == 1) {
                        dataHTML += '<a class="imp" style="color:' + backColor + ';" title="Important">!</a>';
                    }

                    if (activity.Recurrence == 0) {
                        dataHTML += '<a class="glyphicon glyphicon-refresh" title="Daily"></a>';
                    }

                    dataHTML += '</div>';
                }

                dataHTML += '</div>';
            }
        }                    

        $element.append($('<td />').append(dataHTML).addClass("week-data-td " + todayClass));
        //}
        //else if ($(weekHeaderth[i]).hasClass('weekend-color')) {
        //    $element.append($('<td />').addClass("week-data-td weekend-color"));
        //}
        //else {
        //    $element.append($('<td />').addClass("week-data-td " + todayClass));
        //}
    }
}

function bindDayViewDataCustomerJob(startDate, activity, $element) {
    var currentDateView = $(".currentdateview");
    var day = $(currentDateView).data('day');
    var startDateTime = moment(activity.StartDateTime);
    var endDateTime = moment(activity.EndDateTime);
    var isHoliday = false;
    $.each(holidayList, function (i, e) {
        if (moment(startDate).format('YYYY-MM-DD') == moment(e.HolidayDate).format('YYYY-MM-DD')) {
            isHoliday = true;
        }
    });

    if (!isHoliday) {
        if (parseInt(day) == 0 || parseInt(day) == 6) {
            isHoliday = true;
        }
    }

    for (var i = dayStart; i <= dayEnd; i++) {
        //if (parseInt(day) !== 0 && parseInt(day) !== 6) {
        var currTime = i + ':00:00';
        if (i < 10) {
            currTime = '0' + currTime;
        }

        var dataHTML = '';
                    
        //alert("booking St. og endtime: " + startDateTime.format('YYYY-MM-DD') + " "+ endDateTime.format('YYYY-MM-DD') + " startDate: " + startDate + " startTime: " + startTime + " currTime: " + currTime)

        if (moment(endDateTime.format('YYYY-MM-DD')) >= moment(startDate) && moment(startDateTime.format('YYYY-MM-DD')) <= moment(startDate)) {
            var startTime = startDateTime.format('HH:mm:ss');
            var EndTime = endDateTime.format('HH:mm:ss');

            if ((startTime == '00:00:00' && EndTime == '00:00:00') || 
                (moment(startTime, "HH:mm:ss") <= moment(currTime, "HH:mm:ss") && moment(EndTime, "HH:mm:ss") >= moment(currTime, "HH:mm:ss"))
                || (moment(endDateTime.format('YYYY-MM-DD')) > moment(startDate))) {

                for (var k = 0; k < activity.EmployeeList.length; k++) {
                    var backColor = actTypes[activity.ActivityType];
                    var activityTime = endDateTime.diff(startDateTime, 'minutes');
                    var widthClass = 'day-full-width';
                    if (activityTime <= 120) {
                        widthClass = 'day-half-width';
                    }

                    dataHTML += '<div class="main-ac-data ' + widthClass + ' pointer" style="background:' + backColor + ';" onclick="editActivity(' + activity.BookingId + ')">';

                    var employeeName = activity.EmployeeList[k].Name;
                    if (employeeName == '') {
                        employeeName = noEmployeeCalLabel;
                    }
                    dataHTML += '<div class="ac-name">' + employeeName + '</div>';

                    if (activity.Important == 1 || activity.Recurrence == 0) {
                        dataHTML += '<div class="ac-icon">';

                        if (activity.Important == 1) {
                            dataHTML += '<a class="imp" style="color:' + backColor + ';" title="Important">!</a>';
                        }

                        if (activity.Recurrence == 0) {
                            dataHTML += '<a class="glyphicon glyphicon-refresh" title="Daily"></a>';
                        }

                        dataHTML += '</div>';
                    }

                    dataHTML += '</div>';
                }
            }                        
        }

        var todayClass = '';
        if (startDate == moment().format('YYYY-MM-DD')) {
            todayClass = 'today-color';
        }
        else if (isHoliday) {
            todayClass = 'weekend-color';
        }

        $element.append($('<td />').append(dataHTML).addClass("day-data-td " + todayClass));
        //}
        //else {
        //    $element.append($('<td />').addClass("weekend-color"));
        //}
    }
}

function arrangeMonthViewData() {
    $('#monthtbl tbody').find('.month-half-width').each(function (index, item) {
        $(item).parent().append($(item));
    });
}

function arrangeWeekViewData() {
    $('#weektbl tbody').find('.week-half-width').each(function (index, item) {
        $(item).parent().append($(item));

        //alert("HEJ WEEK")
    });
}

function arrangeDayViewData() {
    $('#daytbl tbody').find('.day-half-width').each(function (index, item) {
        $(item).parent().append($(item));
    });
}

function recurrenceChange() {
    if ($('#recurrencedd').val() == '0' || $('#recurrencedd').val() == '-1') {
        $('.no-recurrence').hide();
    }
    else {
        $('.no-recurrence').show();
    }
}

function getDateRangeOfWeek(weekNo, month, year) {
    var d1 = new Date(year, month);
    numOfdaysPastSinceLastMonday = eval(d1.getDay() - 1);
    d1.setDate(d1.getDate() - numOfdaysPastSinceLastMonday);
    var weekNoToday = d1.getWeek();
    var weeksInTheFuture = eval(weekNo - weekNoToday);
    d1.setDate(d1.getDate() + eval(7 * weeksInTheFuture));
    var rangeIsFrom = eval(d1.getMonth() + 1) + "/" + d1.getDate() + "/" + d1.getFullYear();            
    var arrWeekDates = new Array();

    for (var i = 0; i <= 6; i++) {
        var tempDate = new Date(d1);
        tempDate.setDate(tempDate.getDate() + i);
        arrWeekDates.push(tempDate);
    }

    return arrWeekDates;
};

Date.prototype.getWeek = function () {
    var onejan = new Date(this.getFullYear(), 0, 1);
    return Math.ceil((((this - onejan) / 86400000) + onejan.getDay() + 1) / 7);
}

Date.prototype.getWeekOfMonth = function () {
    var month = this.getMonth()
        , year = this.getFullYear()
        , firstWeekday = new Date(year, month, 1).getDay() - 1
        , lastDateOfMonth = new Date(year, month + 1, 0).getDate()
        , offsetDate = this.getDate() + firstWeekday - 1
        , index = 1 // start index at 0 or 1, your choice
        , weeksInMonth = index + Math.ceil((lastDateOfMonth + firstWeekday - 7) / 7)
        , week = index + Math.floor(offsetDate / 7)
    ;

    return week;
    //return week === weeksInMonth ? index + 5 : week;
};

function getWeeksInMonth(year, month) {

    var monthStart = moment().year(year).month(month).date(1);
    var monthEnd = moment().year(year).month(month).endOf('month');
    var numDaysInMonth = moment().year(year).month(month).endOf('month').date();

    //calculate weeks in given month
    var weeks = Math.ceil((numDaysInMonth + monthStart.day()) / 7);
    var weekRange = [];
    var weekStart = moment().year(year).month(month).date(1);
    var i = 0;

    while (i < weeks) {
        var weekEnd = moment(weekStart);


        if (weekEnd.endOf('week').date() <= numDaysInMonth && weekEnd.month() == month) {
            weekEnd = weekEnd.endOf('week').format('LL');
        } else {
            weekEnd = moment(monthEnd);
            weekEnd = weekEnd.format('LL')
        }

        weekRange.push({
            'weekStart': weekStart.format('LL'),
            'weekEnd': weekEnd
        });


        weekStart = weekStart.weekday(7);
        i++;
    }

    return weekRange;
}

function setFitToScreen(cntrl){
    if ($(cntrl).is(':checked')) {
        $('#headerdiv').addClass('table-responsive brd-table');
    }
    else {
        $('#headerdiv').removeClass('table-responsive brd-table');
    }
}
