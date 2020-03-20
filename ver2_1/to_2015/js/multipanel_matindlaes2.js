
$(document).ready(function () {

    $('#oprmat').click(function () {

        
        CreateMat();

    });


    $('#FM_matdato').keypress(function (event) {
        var keycode = (event.keyCode ? event.keyCode : event.which);
        if (keycode == 13) {
            CreateMat();
        }
    });

    $('#FM_matnavn').keypress(function (event) {
        var keycode = (event.keyCode ? event.keyCode : event.which);
        if (keycode == 13) {
            CreateMat();
        }
    });

    $('#FM_matantal').keypress(function (event) {
        var keycode = (event.keyCode ? event.keyCode : event.which);
        if (keycode == 13) {
            CreateMat();
        }
    });

    $('#FM_matpris').keypress(function (event) {
        var keycode = (event.keyCode ? event.keyCode : event.which);
        if (keycode == 13) {
            CreateMat();
        }
    });

    function CreateMat() {

        endpoint = 'convert'
        access_key = 'de1b777d2882c4fe895b0ade03dbb001';

        from = $('#FM_valutakode').val();
        to = $('#basic_valuta').val();
        amount = $('#FM_matpris').val();
        amount = amount.replace(/\,/g, '.')

        $.ajax({
            url: 'https://apilayer.net/api/' + endpoint + '?access_key=' + access_key + '&from=' + from + '&to=' + to + '&amount=' + amount + "&format=1",
            dataType: 'jsonp',
            success: function (json) {

                basic_valuta = from
                basic_kurs = json.result / parseFloat(amount)
                basic_belob = json.result

                if (amount == 0) {
                    basic_kurs = 0
                    basic_belob = 0
                }

               // alert(basic_valuta + " " + basic_kurs + " " + basic_belob);

                $('#basic_valuta').val(basic_valuta)
                $('#basic_kurs').val(basic_kurs)
                $('#basic_belob').val(basic_belob)

                $('#oprmatform').submit();

            },
            error: function (json) {

                alert("error")

            }
        });

    }

});