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

$(document).ready(function () {
    $('#Year').change(function () { ChangeCustomer(); });
});

function ChangeCustomer() {
    var selectedValue = $('#Year option:selected').val();
    if ($.trim(selectedValue).length > 0) {
        if (currentUser != null) {
            GetOrders2(selectedValue);
        }
        else {
            GetOrders(selectedValue);
        }
    }
}

function GetOrders(Year) {
    $.ajax({
        url: orderUrl,
        data: { selectYear: Year, type: '污水使用' },
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
                    if (item.Text.split('-')[0].trim() === '') {
                        // 如果 A 为空值，跳过当前迴圈
                        return true; // 相当于 continue
                    }
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

function GetOrders2(Year) {
    $.ajax({
        url: orderUrl,
        data: { selectYear: Year, type: '雨水使用' },
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
                    if (item.Text.split('-')[0].trim() === '') {
                        // 如果 A 为空值，跳过当前迴圈
                        return true; // 相当于 continue
                    }
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