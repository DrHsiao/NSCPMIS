$(document).ready(function () {
    $('#Year').change(function () { ChangeCustomer(); });
});

function ChangeCustomer() {
    var selectedValue = $('#Year option:selected').val();
    if ($.trim(selectedValue).length > 0) {
        GetOrders(selectedValue);
    }
}

function GetOrders(Year) {
    $.ajax({
        url: orderUrl,
        data: { selectYear: Year },
        type: 'post',
        cache: false,
        async: false,
        dataType: 'json',
        success: function (data1) {
            var data = JSON.parse(data1);
            if (data.length > 0) {
                $('#Value').empty();
                $('#Value').append($('<option value="" disabled selected></option>').val('').text('請選擇'));
                $.each(data, function (i, item) {
                    $('#Value').append($('<option></option>').val(item.Value).text(item.Text));
                })
            }
            else {
                $('#Value').empty();
                $('#Value').append($('<option></option>').val('').text('------'));
            }
        }
    });
}

$(function () {
    $("#Value").change(function () {
        var selectedValue = $(this).val();
        $.ajax({
            url: pjYearUrl,
            type: 'post',
            cache: false,
            async: false,
            data: { selectedPID: selectedValue },
            dataType: 'json',
            success: function (data1) {
                var data = JSON.parse(data1);
                if (data.length > 0) {
                    $('#PJYear').empty();
                    $('#PJYear').append($('<div></div>').val('').text('計畫年度'));
                    $.each(data, function (i, item) {
                        $('#PJYear').append($('<label></label>').append($('<input>', {
                            type: 'checkbox',
                            name: 'PJYear',
                            value: item.SYR
                        })).append(item.SYR));
                    });
                }
            }
        });
    });
});

$(function () {
    $("#choose").change(function () {
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