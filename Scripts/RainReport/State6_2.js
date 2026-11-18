function toggleAll(source) {
    var checkboxes = document.querySelectorAll('input[type="checkbox"]');
    checkboxes.forEach(function (checkbox) {
        if (checkbox !== source) {
            checkbox.checked = source.false;
        }
    });
}

function toggleOnly(source) {
    var allcheckbox = document.getElementById("allcheck");
    if (allcheckbox !== source) {
        allcheckbox.checked = source.false;
    }
}

$(function () {
    $("#Value").change(function () {
        var selectedValue = $(this).val();
        var selectedValue2 = $('#Year option:selected').val();
        $.ajax({
            url: periodUrl,
            type: 'post',
            cache: false,
            async: false,
            data: { selectedPID: selectedValue },
            dataType: 'json',
            success: function (data1) {
                var data = JSON.parse(data1);
                if (data.length > 0) {
                    $('#Period').empty();
                    $.each(data, function (i, item) {
                        if (item.SEQ_NO === 0) {
                            $('#Period').append($('<option></option>').val('0').text('原預算'));
                        }
                        else {
                            $('#Period').append($('<option></option>').val(item.SEQ_NO).text('第' + item.SEQ_NO + '期'));
                        }
                    });
                    //$('#Period').append($('<option></option>').val('0').text('原預算'));
                }
                else {
                    $('#Period').empty();
                    $('#Period').append($('<option></option>').val('0').text('原預算'));
                }
            }
        });

        $.ajax({
            url: subjectUrl,
            type: 'post',
            cache: false,
            async: false,
            data: { selectedPID: selectedValue, SYR: selectedValue2 },
            dataType: 'json',
            success: function (data1) {
                var data = JSON.parse(data1);
                if (data.length > 0) {
                    $('#Subject').empty();
                    $('#Subject').append($('<div></div>').val('').text('選擇會計科目:'));
                    $.each(data, function (i, item) {
                        $('#Subject').append($('<div></div>').append($('<input>', {
                            type: 'checkbox',
                            name: 'Subject',
                            value: item.ACC_ID
                        })).append(item.ACC_DESC));
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
                        return true;
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
                        return true;
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