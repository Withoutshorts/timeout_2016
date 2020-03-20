
$(document).ready(function () {

    $('#oprmat').click(function () {

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

                $('#basic_valuta').val(basic_valuta)
                $('#basic_kurs').val(basic_kurs)
                $('#basic_belob').val(basic_belob)

                $('#oprmatform').submit();

            },
            error: function (json) {

                alert("error")

            }
        });


    });

});