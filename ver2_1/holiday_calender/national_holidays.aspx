<%@ Page Language="C#" AutoEventWireup="true" CodeFile="national_holidays.aspx.cs" Inherits="national_holidays" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        th, td {
            border: 1px solid white;
            padding: 2px;
            text-align: center;            
        }

        th {
            background-color: #5B9BD5;
            line-height: 30px;
            color: white;
            font-weight: bold;
            font-size: 15px;
        }

        td {
            height: 25px;
        }

        tr {
            background-color: #EAEFF7;
        }

        tr:nth-child(even) {
            background-color: #D2DEEF;
        }

        .duration {
            width: 13%;
        }

        .name {
            width: 27%;
            text-align: left;
            padding-left: 5px;
        }

        .year {
            width: 6%;
        }

        .name input[type=text] {
            text-align: left !important;
            width: 97%;
        }

        select {
            height: 21px;
            width: 95%;
        }

        input[type=text] {
            width: 85%;
            text-align: center;
        }

        .btn{
            background-color: #5B9BD5;
            color:white;
            padding: 7px 15px;
            float: right;
            margin-bottom: 10px;
            margin-top: 10px;            
            font-size: 14px;
            font-weight:bold;
            border:none;
            cursor:pointer;
        }

        .maindiv{
            width: 1050px; 
            margin: 0 auto;
            font-family: 'Open Sans', arial, verdana, sans-serif !important;
            font-size: 13px;
        }

        .left{
            float:left;
        }

        .right{
            float:right;
        }

        .error{
            vertical-align: bottom;
            line-height: 50px;
            color: red;
            font-weight: bold;
        }

        .success{
            vertical-align: bottom;
            line-height: 50px;
            color: green;
            font-weight: bold;
        }
    </style>

    <script type="text/javascript" src="js/jquery-2.2.3.js"></script>
    <script type="text/javascript" src="js/jquery.inputmask.bundle.js"></script>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <div class="maindiv">
                <div>
                    <div class="left">
                        <label class="error" style="display:none;">Error!</label>
                        <label class="success" style="display:none;">Success!</label>
                    </div>
                    <div class="right">
                        <a class="btn add-new">Add New</a>
                    </div>                    
                </div>
                <div>
                    <asp:Literal ID="ltrHolidays" runat="server"></asp:Literal>
                </div>
                <div>
                    <a class="btn save-holiday" onclick="saveHolidays()" style="display:none;">Save</a>
                </div>
            </div>
        </div>        
    </form>

    <script type="text/javascript">
        var newHolidayHTML = '<tr id="tr_#NO#"><td class="duration"><select id="ddl_Duration_#NO#"><option value="1" selected="selected">Open for timerec</option><option value="0">Close for timerec</option></select></td><td class="name"><input type="text" id="txt_Name_#NO#" /></td><td class="year"><input type="text" id="txt_Year_#NO#_1" /></td><td class="year"><input type="text" id="txt_Year_#NO#_2" /></td><td class="year"><input type="text" id="txt_Year_#NO#_3" /></td><td class="year"><input type="text" id="txt_Year_#NO#_4" /></td><td class="year"><input type="text" id="txt_Year_#NO#_5" /></td><td class="year"><input type="text" id="txt_Year_#NO#_6" /></td><td class="year"><input type="text" id="txt_Year_#NO#_7" /></td><td class="year"><input type="text" id="txt_Year_#NO#_8" /></td><td class="year"><input type="text" id="txt_Year_#NO#_9" /></td><td class="year"><input type="text" id="txt_Year_#NO#_10" /></td></tr>';
        var counter = 0;

        $(document).ready(function () {
            $('.error, .success').hide();
            counter = $('#tbl_Holidays tr td').parent().length;

            if ($('#tbl_Holidays tr td').length == 0) {
                $('.save-holiday').hide();
            }
            else {
                $('.save-holiday').show();
                setInputDateMask();
            }            

            $('.add-new').click(function () {
                counter++;

                var holidayHTML = newHolidayHTML.replace(/\#NO#/g, '0' + counter);
                $('#tbl_Holidays').append(holidayHTML);
                setInputDateMask();
                $('.save-holiday').show();
            });
        });

        function setInputDateMask() {
            $('.year input').inputmask("99/99");
        }

        function saveHolidays() {
            var strHolidays = '';

            if ($('#tbl_Holidays tr td').parent().length > 0) {
                for (var i = 1; i <= counter; i++) {
                    var $nameCntrl = $('#txt_Name_0' + i);
                    if ($nameCntrl.val() != '') {                        
                        var $durationCntrl = $('#ddl_Duration_0' + i);

                        for (var j = 1; j <= 10; j++) {
                            var dateCntrl = $('#txt_Year_0' + i + '_' + j);

                            if (dateCntrl.val() != '') {
                                var $td = dateCntrl.parent();
                                var $th = $td.closest('table').find('th').eq($td.index());
                                var holidayDate = $th.text() + '/' + dateCntrl.val();

                                strHolidays += holidayDate;
                                strHolidays += '###' + $durationCntrl.val();
                                strHolidays += '###' + $nameCntrl.val();
                                strHolidays += '@@@';
                            }
                        }
                    }                    
                }
            }

            if (strHolidays != '') {                
                $.ajax({
                    type: "POST",
                    url: "national_holidays.aspx/SaveHolidays",
                    data: '{holidayData: "' + strHolidays + '" }',
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    success: function (response) {
                        if (response.d > 0) {
                            $('.success').show();
                            $('.success').text('Holidays saved successfully.');
                            $('.hide').show();
                        }
                        else {
                            $('.success').hide();
                            $('.error').show();
                            $('.error').text('Error! while saving holidays');
                        }
                    },
                    failure: function (response) {
                        $('.success').hide();
                        $('.error').show();
                        $('.error').text('Error! while saving holidays');
                    }
                });
            }
        }
    </script>
</body>
</html>
