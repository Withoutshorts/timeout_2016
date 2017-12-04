

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
var splitCounter = 0;

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
    //alert("HER")

    //$("#sortbymonthdd").select('1');
    //$(".sortby option[id='3']").attr("selected", "selected");
    //$('#sortbymonthdd option').eq(2).prop('selected', true);
    //$(".viewtype").val('4');

    $.ajax({
        type: "POST",
        url: "ressplan_2017.aspx/GetHolidays",
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        success: function (response) {
            holidayList = response.d;

            $("#mainTemplate").tmpl().appendTo("#maindiv");
            changeDayGroup('month');

            bindEmployees();
        },
    });
});

function fromTimeChange(cntrl) {
    var currCounter = $(cntrl).data('counter');

    if ($('#dtstart_' + currCounter).val() == $('#dtend_' + currCounter).val() && $(cntrl).val() > $('#totime_' + currCounter).val()) {
        $(cntrl).addClass('error');
        $(cntrl).parent().find('p').remove()
        $(cntrl).parent().append('<p class="error-txt">Start time should be less than end time!</p>');
    }
    else {
        $(cntrl).removeClass('error');
        $(cntrl).parent().find('p').remove();
    }
}

function toTimeChange(cntrl) {
    var currCounter = $(cntrl).data('counter');

    if ($('#dtstart_' + currCounter).val() == $('#dtend_' + currCounter).val() && $(cntrl).val() < $('#fromtime_' + currCounter).val()) {
        $(cntrl).addClass('error');
        $(cntrl).parent().find('p').remove();
        $(cntrl).parent().append('<p class="error-txt">End time should be more than start time!</p>');
    }
    else {
        $(cntrl).removeClass('error');
        $(cntrl).parent().find('p').remove();
    }
}

function loadTemplate(templatename) {
    $(".removetbl").remove();
    $(templatename).tmpl().appendTo("#headerdiv");
    populateDropdown();

    $('.oliver').multiselect({
        selectAllValue: 'multiselect-all',
        enableCaseInsensitiveFiltering: true,
        enableFiltering: true
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

            if (week > 52) {
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

    for (var i = 0; i < lastDate; i++) {
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

function bindWeekView(week, month, year, startdate) {
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

    ddsortbyTypePresel = $("#ddsortbyTypePresel").val();

    //alert(ddsortbyTypePresel)

    var option = '';
    for (var i = 0; i < sortbydd.length; i++) {

        if (i == ddsortbyTypePresel) {
            sortByDDSel = "SELECTED"
        } else {
            sortByDDSel = ""
        }
        option += '<option value="' + i + '" ' + sortByDDSel + '>' + sortbydd[i] + '</option>';
    }

    $('.sortby').append(option);

    if (sortbySelectedValue != "") {
        $(".sortby").val(sortbySelectedValue);
    }

    /********** View type dropdown *********************************/
    var viewtypedd = ['View type', 'Project time', 'Absence', 'Important'];
    //var viewtypedd = ['Project time'];
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

                if (olivereoption != "0" && olivereoption != '') {
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

    //alert($("#ddjobidPresel").val())

    if ($("#ddjobidPresel").val() != "") {
        oliverSelectedValue = $("#ddjobidPresel").val()

        //alert(oliverSelectedValue)
        var $el = $(".oliver");

        //LOOP ALL CHECKBOXES in DropDown
        $el.find('option:not(:selected)').each(function () {
            if (oliverSelectedValue == $(this).val()) {

                // --> HOW TO CHECK THIS BOX?
                //alert("TT" + $(".oliver").val() + $(this).val());
                //$(this + 'option').prop('checked', true);

                //$('.oliver').attr('selected', 'selected');
                //$(this).prop('checked', true);
                //$(this).attr('selected', 'selected');

                $(this).attr('selected', 'selected');
                $(".oliver").multiselect('rebuild');

            }

        });

    } else {
        oliverSelectedValue = $(".oliver").val();
    }


    //return false;

    //alert(oliverSelectedValue)

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
                                bindMonthViewData(lastDate, year, month, activities, $(dataMonths), employee.Id, employee.Name);

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
                                bindWeekViewData(arrDates, activities, $(dataWeeks), employee.Id, employee.Name);
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

                                bindDayViewData(startDate, activities, $(dataDays), employee.Id, employee.Name);

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
                            bindMonthViewData(lastDate, year, month, activities, $(dataMonths), employee.Id, employee.Name);

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
                            bindWeekViewData(arrDates, activities, $(dataWeeks), employee.Id, employee.Name);
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

                            bindDayViewData(startDate, activities, $(dataDays), employee.Id, employee.Name);
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
    $('#lnkSplitBooking').hide();
    $('#main_split_Bookings').html('');
    splitCounter = 0;
    $('#chksplit').prop('checked', false);

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

            $('#fromtime_0').val(response.d.StartTime);
            $('#totime_0').val(response.d.EndTime);

            $('#dtstart_0').val(getDisplayFormattedDate(response.d.StartDate));
            $('#dtstart_0').attr('date', getSavedFormattedDate(response.d.StartDate));

            $('#dtend_0').val(getDisplayFormattedDate(response.d.EndDate));
            $('#dtend_0').attr('date', getSavedFormattedDate(response.d.EndDate));

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
        afterModelShown(0);
    });
}

//bind datapicker after modal shown
function afterModelShown(no) {

    $('.input-group.date.start_' + no).datepicker({
        format: 'dd-mm-yyyy',
        autoclose: true
    }).on('changeDate', function (e) {
        compareDate(e.date, $(this));
    });

    $('.input-group.date.end_' + no).datepicker({
        format: 'dd-mm-yyyy',
        autoclose: true
    }).on('changeDate', function (e) {
        compareDate(e.date, $(this));
    });
}

function compareDate(selectedDate, cntrl) {
    $(cntrl).find('input:text').attr('date', getSavedFormattedDate(selectedDate));

    var currCounter = $(cntrl).find('input:text').data('counter');

    var datestartObj = new Date($('.input-group.date.start_' + currCounter).datepicker('getDate'));
    var momentStartObj = moment(datestartObj);
    var momentStartString = momentStartObj.format('YYYY-MM-DD');

    var dateEndObj = new Date($('.input-group.date.end_' + currCounter).datepicker('getDate'));
    var momentEndObj = moment(dateEndObj);
    var momentEndString = momentEndObj.format('YYYY-MM-DD')

    var $endDateCntrl = $('.input-group.date.end_' + currCounter);
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
    if ($('#dtstart_0').val() == $('#dtend_0').val() && $('#fromtime_0').val() > $('#totime_0').val()) {
        $('#totime_0').addClass('error');
        $('#totime_0').parent().append('<p class="error-txt">End time should less than start time.</p>');
    }
    else {
        if ($('#totime_0').hasClass('error')) {
            $('#totime_0').removeClass('error');
        }
    }
}

function saveActivity() {
    timeValidation();

    var taskData = JSON.stringify($('form').serializeArray());
    console.log(taskData);

    var model = {};
    model.BookingId = bookingId;

    model.StartDate = $('#dtstart_0').attr('date'); //'2017-04-02' //
    model.EndDate = $('#dtend_0').attr('date');
    model.StartTime = $('#fromtime_0').val();
    model.StartDateTime = model.StartDate + ' ' + model.StartTime;

    model.EndTime = $('#totime_0').val();
    model.EndDateTime = model.EndDate + ' ' + model.EndTime;

    if (sortbySelectedValue !== '1' && sortbySelectedValue !== '2') {
        var arrEmployees = [];
        arrEmployees.push($('#empchk').val());
        model.Employees = arrEmployees;
    }
    else {
        model.Employees = $('#empchk').val();
    }

    model.Split = $('#chksplit').prop('checked') ? 1 : 0;
    model.Important = $('#chkimp').prop('checked') ? 1 : 0;

    var arrEmployeeActivities = [];
    if (splitCounter > 0) {
        for (var i = 1; i <= splitCounter; i++) {

            if ($('#split_booking_' + i).length > 0) {
                var employeeActivity = {};
                employeeActivity.EmployeeId = $('#empSplit_' + i).val();
                employeeActivity.EmployeeName = $('#empSplit_' + i + ' option:selected').text();

                var splitStartDate = $('#dtstart_' + i).attr('date');
                var splitStartTime = $('#fromtime_' + i).val();

                var splitEndDate = $('#dtend_' + i).attr('date');
                var splitEndTime = $('#totime_' + i).val();

                employeeActivity.StartDateTime = splitStartDate + ' ' + splitStartTime;
                employeeActivity.EndDateTime = splitEndDate + ' ' + splitEndTime;

                arrEmployeeActivities.push(employeeActivity);
            }
        }
    }

    model.EmployeeActivities = arrEmployeeActivities;
        
    if (!$('#activityModal').find('div.col-lg-4 input').hasClass('error')) {
        $.ajax({
            type: "POST",
            url: "ressplan_2017.aspx/SaveActivity",
            data: JSON.stringify({ 'model': model }),
            contentType: "application/json; charset=utf-8",
            dataType: "json",
            success: function (response) {
                if (response.d != null && response.d.length > 0) {
                    if (response.d[0] == "1") {
                        bookingId = 0;
                        $('#activityModal').modal('hide');
                        resourceChange();
                    }
                    else {
                        alert(response.d[1]);
                    }
                }
                else {
                    alert('Error! while updating activity.');
                }
            },
        });
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

function bindMonthViewData(lastDate, year, month, activities, $element, empId, employee) {

    var monthHeaderth = $('tr#monthtbldate th.month-date-ths');

    for (var i = 0; i < lastDate; i++) {

        var currDate = year //+ '-' + (parseInt(month) + 1) + '-' + (i + 1);

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

            dataHTML += '<div class="main-ac-data ' + widthClass + ' cur-default draggable" style="background:' + backColor + ';" data-booking-id=' + arrActivities[k].BookingId + '>';

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

        $element.append($('<td class="droppable" data-emp-id=' + empId + ' data-emp-name="' + employee + '" data-date="' + currDate + '" />').append(dataHTML).addClass(todayClass));
    }
}

function bindWeekViewData(arrDates, activities, $element, empId, employee) {
    var weekHeaderth = $('tr#weektbldaystr th.week-day-ths');
    var weekHeaderDateth = $('tr#weektbldatetr th.week-date-ths');
    startDay = $("#hidprecallstartdate").val();

    for (var i = 0; i < arrDates.length; i++) {
        var todayClass = '';
        if ($(weekHeaderth[i]).hasClass('today-color')) {
            todayClass = 'today-color';
        }
        else if ($(weekHeaderth[i]).hasClass('weekend-color')) {
            todayClass = 'weekend-color';
        }

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

            dataHTML += '<div class="main-ac-data ' + widthClass + ' cur-default draggable" style="background:' + backColor + ';" data-booking-id=' + arrActivities[k].BookingId + '>';

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

        $element.append($('<td data-emp-id=' + empId + ' data-emp-name="' + employee + '" data-date="' + moment(arrDates[i]).format('YYYY-MM-DD') + '" />').append(dataHTML).addClass("droppable week-data-td " + todayClass));

        startDay++;
    }
}

function bindDayViewData(startDate, activities, $element, empId, employee) {
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

        ////var arrActivities = [];
        ////for (var a = 0; a < activities.length; a++) {
        ////    var activity = activities[a];
            
        ////    if (moment(moment(activity.StartDateTime).format('YYYY-MM-DD')).isSameOrBefore(moment(startDate)) &&
        ////            moment(moment(activity.EndDateTime).format('YYYY-MM-DD')).isSameOrAfter(moment(startDate))) {

        ////        if (moment(moment(activity.StartDateTime).format('YYYY-MM-DD')).isSame(moment(startDate)) &&
        ////            moment(moment(activity.EndDateTime).format('YYYY-MM-DD')).isSame(moment(startDate))) {
        ////            if (moment(moment(activity.StartDateTime).format('HH:mm:ss'), 'HH:mm:ss').isSameOrBefore(moment(currTime, "HH:mm:ss")) &&
        ////                moment(moment(activity.EndDateTime).format('HH:mm:ss'), 'HH:mm:ss').isSameOrAfter(moment(currTime, "HH:mm:ss"))) {
        ////                arrActivities.push(activity);
        ////            }
        ////            else if (moment(moment(activity.EndDateTime).format('HH:mm:ss'), 'HH:mm:ss') < moment(currTime, "HH:mm:ss")) {
        ////                arrActivities.push(activity);
        ////            }
        ////        }
        ////        else if (moment(moment(activity.EndDateTime).format('YYYY-MM-DD')).isSame(moment(startDate))){
        ////            if (moment(activity.EndDateTime).format('HH:mm:ss') == '00:00:00' ||
        ////                moment(moment(activity.EndDateTime).format('HH:mm:ss'), 'HH:mm:ss').isSameOrAfter(moment(currTime, "HH:mm:ss"))) {

        ////                if (activity.BookingId == 37) {
        ////                    console.log('1');
        ////                }
        ////                arrActivities.push(activity);
        ////            }
        ////        }
        ////        else if (moment(moment(activity.StartDateTime).format('YYYY-MM-DD')).isSame(moment(startDate))){
        ////            if (moment(activity.StartDateTime).format('HH:mm:ss') == '00:00:00' ||
        ////                moment(moment(activity.StartDateTime).format('HH:mm:ss'), 'HH:mm:ss').isSameOrBefore(moment(currTime, "HH:mm:ss"))) {

        ////                if (activity.BookingId == 37) {
        ////                    console.log('2');
        ////                }
        ////                arrActivities.push(activity);
        ////            }
        ////        }
        ////        else if (moment(moment(activity.StartDateTime).format('YYYY-MM-DD')).isBefore(moment(startDate)) &&
        ////            moment(moment(activity.EndDateTime).format('YYYY-MM-DD')).isAfter(moment(startDate))){

        ////            if (activity.BookingId == 37) {
        ////                console.log('3');
        ////            }
        ////            arrActivities.push(activity);
        ////        }
        ////    }
        ////}
        
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

            dataHTML += '<div class="main-ac-data ' + widthClass + ' cur-default draggable" style="background:' + backColor + ';" data-booking-id=' + arrActivities[k].BookingId + '>';

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

        $element.append($('<td data-emp-id=' + empId + ' data-emp-name="' + employee + '" data-date="' + startDate + ' ' + moment(currTime, "HH:mm:ss").format('HH:mm') + '" />').append(dataHTML).addClass("droppable day-data-td " + todayClass));
    }
}

function bindMonthViewDataCustomerJob(lastDate, year, month, activity, $element) {
    var monthHeaderth = $('tr#monthtbldate th.month-date-ths');

    for (var i = 0; i < lastDate; i++) {

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

        var dataHTML = '';

        if (moment(moment(activity.EndDateTime).format('YYYY-MM-DD')) >= moment(currDate) && moment(moment(activity.StartDateTime).format('YYYY-MM-DD')) <= moment(currDate)) {
            for (var k = 0; k < activity.EmployeeList.length; k++) {
                var backColor = actTypes[activity.ActivityType];
                var activityTime = moment(activity.EndDateTime).diff(activity.StartDateTime, 'minutes');
                var widthClass = 'month-full-width';
                if (activityTime <= 240) {
                    widthClass = 'month-half-width';
                }

                dataHTML += '<div class="main-ac-data ' + widthClass + ' cur-default draggable" style="background:' + backColor + ';" data-booking-id=' + activity.BookingId + '>';
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
            todayClass = 'today-color droppable';
        }
        else if ($(monthHeaderth[i]).hasClass('weekend-color')) {
            todayClass = 'weekend-color droppable';
        }
        else {
            todayClass = 'droppable';
        }

        $element.append($('<td data-date="' + currDate + '" />').append(dataHTML).addClass(todayClass));
    }
}

function bindWeekViewDataCustomerJob(arrDates, activity, $element) {
    var weekHeaderth = $('tr#weektbldaystr th.week-day-ths');

    for (var i = 0; i < arrDates.length; i++) {
        var todayClass = '';
        if ($(weekHeaderth[i]).hasClass('today-color')) {
            todayClass = 'today-color';
        }
        else if ($(weekHeaderth[i]).hasClass('weekend-color')) {
            todayClass = 'weekend-color';
        }

        var dataHTML = '';

        if (moment(moment(activity.EndDateTime).format('YYYY-MM-DD')) >= moment(arrDates[i]) && moment(moment(activity.StartDateTime).format('YYYY-MM-DD')) <= moment(arrDates[i])) {
            for (var k = 0; k < activity.EmployeeList.length; k++) {
                var backColor = actTypes[activity.ActivityType];
                var activityTime = moment(activity.EndDateTime).diff(activity.StartDateTime, 'minutes');
                var widthClass = 'week-full-width';
                if (activityTime <= 120) {
                    widthClass = 'week-half-width';
                }

                dataHTML += '<div class="main-ac-data ' + widthClass + ' cur-default draggable" style="background:' + backColor + ';" data-booking-id=' + activity.BookingId + '>';

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

        $element.append($('<td data-date="' + moment(arrDates[i]).format('YYYY-MM-DD') + '" />').append(dataHTML).addClass("droppable week-data-td " + todayClass));
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
        var currTime = i + ':00:00';
        if (i < 10) {
            currTime = '0' + currTime;
        }

        var dataHTML = '';

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

                    dataHTML += '<div class="main-ac-data ' + widthClass + ' cur-default draggable" style="background:' + backColor + ';" data-booking-id=' + activity.BookingId + '>';

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

        $element.append($('<td data-date="' + startDate + ' ' + moment(currTime, "HH:mm:ss").format('HH:mm') + '" />').append(dataHTML).addClass("droppable day-data-td " + todayClass));
    }
}

function arrangeMonthViewData() {
    $('#monthtbl tbody').find('.month-half-width').each(function (index, item) {
        $(item).parent().append($(item));
    });

    applyDraggable();
    applyContextMenu();
}

function arrangeWeekViewData() {
    $('#weektbl tbody').find('.week-half-width').each(function (index, item) {
        $(item).parent().append($(item));
    });

    applyDraggable();
    applyContextMenu();
}

function arrangeDayViewData() {
    $('#daytbl tbody').find('.day-half-width').each(function (index, item) {
        $(item).parent().append($(item));
    });

    applyDraggable();
    applyContextMenu();
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

function setFitToScreen(cntrl) {
    if ($(cntrl).is(':checked')) {
        $('#headerdiv').addClass('table-responsive brd-table');
    }
    else {
        $('#headerdiv').removeClass('table-responsive brd-table');
    }
}

function showLoading() {
    $('.page-loader').show();
}

function hideLoading() {
    $('.page-loader').hide();
}

function applyDraggable() {
    $(".draggable").draggable({
        revert: "invalid",
        start: function (event, ui) {
            $(ui.helper[0]).css('z-index', 9);
            $(ui.helper[0]).css('cursor', 'move');
        },
        stop: function (event, ui) {
            $(ui.helper[0]).css('z-index', '');
            $(ui.helper[0]).css('cursor', '');
        }
    });
    $(".droppable").droppable({
        accept: '.draggable',
        drop: function (event, ui) {
            var draggedBookingId = $(ui.draggable[0]).data('booking-id');
            var draggedEmpId = $(ui.draggable[0]).parent().data('emp-id');
            var droppedDate = $(this).data('date');
            var droppedEmpId = $(this).data('emp-id');
            var droppedEmpName = $(this).data('emp-name');

            if (draggedEmpId == undefined) {
                draggedEmpId = 0;
            }

            if (droppedEmpId == undefined) {
                droppedEmpId = 0;
            }

            if (droppedEmpName == undefined) {
                droppedEmpName = '';
            }

            showLoading();

            setTimeout(function () {
                var isDateOnly = false;
                if (currentView != 'day') {
                    isDateOnly = true;
                }

                $.ajax({
                    type: "POST",
                    url: "ressplan_2017.aspx/UpdateBooking",
                    data: JSON.stringify({ 'bookingId': draggedBookingId, 'droppedDate': droppedDate, 'isDateOnly': isDateOnly, 'draggedEmpId': draggedEmpId, 'droppedEmpId': droppedEmpId }),
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    success: function (response) {
                        hideLoading();

                        if (response.d > 0) {
                            resourceChange();
                        }
                        else if (response.d == -1) {
                            alert('Error! The booking is already exists on ' + droppedEmpName);
                            return ui.draggable.animate({ left: 0, top: 0 }, 500);
                        }
                        else {
                            alert('Error! while updating activity.');
                        }
                    },
                    error: function (jqXHR, exception) {
                        hideLoading();
                        alert(jqXHR.responseText);
                        //jqXHR.status
                        //jqXHR.responseText
                        //exception
                    }
                });
            }, 1500);
        }
    });
}

function applyContextMenu() {
    $.contextMenu({
        selector: '.draggable',
        callback: function (key, options) {
            console.log(key);
            console.log(options);
        },
        items: {
            "edit": {
                name: "Edit",
                icon: "fa-pencil-square-o",
                callback: function (key, options) {
                    var bookingId = $(this).data('booking-id');
                    if (bookingId != undefined) {
                        editActivity(bookingId);
                    }
                }
            }
        }
    });
}

function enableSplitBooking(cntrl) {
    if ($(cntrl).prop('checked')) {
        $('#lnkSplitBooking').show();
    }
    else {
        $('#lnkSplitBooking').hide();
        $('#main_split_Bookings').html('');
        splitCounter = 0;
    }
}

function loadSplitBooking() {
    splitCounter = splitCounter + 1;
    var splitHTML = getSplitBookingHTML();
    splitHTML = splitHTML.replace(/\##ID##/g, splitCounter);
    $('#main_split_Bookings').append(splitHTML);

    var $options = $("#empchk > option").clone();
    $('#empSplit_' + splitCounter).append($options);
    $('#empSplit_' + splitCounter + ' option[value="' + $("#empchk").val() + '"]').remove();

    $('#dtstart_' + splitCounter).val($('#dtstart_0').val());
    $('#dtstart_' + splitCounter).attr('date', $('#dtstart_0').attr('date'));

    $('#dtend_' + splitCounter).val($('#dtend_0').val());
    $('#dtend_' + splitCounter).attr('date', $('#dtend_0').attr('date'));

    $('#fromtime_' + splitCounter).val($('#fromtime_0').val());
    $('#totime_' + splitCounter).val($('#totime_0').val());

    afterModelShown(splitCounter);
}

function removeSplitBooking(id) {
    $('#split_booking_' + id).remove();

    if ($('#main_split_Bookings > div').length == 0) {
        splitCounter = 0;
    }
}

function getSplitBookingHTML() {
    return '<div class="split-booking row" id="split_booking_##ID##">\
                <div class="remove-booking"><a href="javascript:void(0)" onclick="removeSplitBooking(##ID##)">×</a></div>\
                <div class="col-lg-12">\
                    <div class="row">\
                        <div class="col-lg-2 lh-34">Employee:</div>\
                        <div class="col-lg-4">\
                            <select id="empSplit_##ID##" name="empSplit_##ID##" class="form-control input-small"></select>\
                        </div>\
                    </div>\
                    <div class="row">\
                        <div class="col-lg-2 lh-34">Start:</div>\
                        <div class="col-lg-4">\
                            <div class="input-group date start_##ID##">\
                                <input type="text" id="dtstart_##ID##" data-counter="##ID##" class="form-control input-small" name="dtstart_##ID##" placeholder="dd-mm-yyyy" />\
                                <span class="input-group-addon input-small">\
                                    <span class="fa fa-calendar open-datetimepicker"></span>\
                                </span>\
                            </div>\
                        </div>\
                        <div class="col-lg-2 lh-34"></div>\
                        <div class="col-lg-4">\
                            <input type="time" id="fromtime_##ID##" data-counter="##ID##" onchange="fromTimeChange(this)" name="fromtime_##ID##" class="form-control input-small" placeholder="00:00" />\
                        </div>\
                    </div>\
                    <div class="row">\
                        <div class="col-lg-2 lh-34">End:</div>\
                        <div class="col-lg-4">\
                            <div class="input-group date end_##ID##">\
                                <input type="text" id="dtend_##ID##" data-counter="##ID##" class="form-control input-small" name="dtend_##ID##" value="" placeholder="dd-mm-yyyy" />\
                                <span class="input-group-addon input-small">\
                                    <span class="fa fa-calendar"></span>\
                                </span>\
                            </div>\
                            <p class="error-txt" style="display: none;">End Date should be more than start date!</p>\
                        </div>\
                        <div class="col-lg-2 lh-34"></div>\
                        <div class="col-lg-4">\
                            <input type="time" id="totime_##ID##" data-counter="##ID##" onchange="toTimeChange(this)" name="totime_##ID##" class="form-control input-small" placeholder="00:00" />\
                        </div>\
                    </div>\
                </div>\
            </div>';
}