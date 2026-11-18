$(function () {
    $("#choose").change(function () {
        var selectedValue = $(this).val();
        $.ajax({
            url: yearUrl,
            type: 'post',
            cache: false,
            async: false,
            dataType: 'json',
            data: { PID: selectedValue },
            success: function (data1) {
                var data = JSON.parse(data1);
                if (data.length > 0) {
                    $('#SYR').empty();
                    $('#SYR').append($('<option value="" disabled selected></option>').val('').text('請選擇'));
                    $.each(data, function (i, item) {
                        $('#SYR').append($('<option></option>').val(item.SYR).text(item.SYR));
                    })
                }
                else {
                    $('#SYR').empty();
                    $('#SYR').append($('<option></option>').val('').text('------'));
                }
            }
        });
    });
});

$(function () {
    $("#choosestate").change(function () {
        var selectedValue = $(this).val();
        $.ajax({
            url: stateUrl,
            type: 'post',
            cache: false,
            async: false,
            data: { value: selectedValue },
            success: function (result) {
                window.location.href = result;
            }
        });
    });
});